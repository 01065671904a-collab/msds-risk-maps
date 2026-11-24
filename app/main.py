import io
import os
import re
import xml.etree.ElementTree as ET

import fitz
import pytesseract
import requests
from bs4 import BeautifulSoup
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

app = FastAPI(title="MSDS 유해화학물질 통합관리")
app.mount("/static", StaticFiles(directory="app/static"), name="static")

BASE_API = "https://apis.data.go.kr/B552468/msdschem1"
KOSHA_WEB_BASE = "https://msds.kosha.or.kr/MSDSInfo"
SERVICE_KEY = os.getenv("KOSHA_SERVICE_KEY", "").strip()

CAS_RE = re.compile(r"(?<!\d)(\d{2,7}-\d{2}-\d)(?!\d)")
FLEX_CAS_RE = re.compile(r"(?<!\d)(\d{2,7})\s*[-–—‐‑‒−]\s*(\d{2})\s*[-–—‐‑‒−]\s*(\d)(?!\d)")
H_CODE_RE = re.compile(
    r"\bH(?:300|301|302|304|310|311|312|314|315|317|318|319|320|330|331|332|334|335|336|340|341|350|351|360|361|362|370|371|372|373)\b"
)

H_GRADE_MAP = {
    "H340": 4, "H341": 4, "H350": 4, "H351": 4, "H360": 4, "H361": 4, "H362": 4, "H370": 4,
    "H300": 3, "H301": 3, "H310": 3, "H311": 3, "H330": 3, "H331": 3,
    "H314": 3, "H318": 3, "H317": 3, "H334": 3, "H372": 3, "H373": 3,
    "H302": 2, "H312": 2, "H332": 2, "H304": 2, "H371": 2,
    "H315": 1, "H319": 1, "H320": 1, "H335": 1, "H336": 1,
}

INVALID_CHEM_NAMES = {
    "NORMAL SERVICE",
    "NORMAL SERVICE.",
    "00",
    "NULL",
    "NONE",
    "N/A",
    "-",
    "자료없음",
    "해당없음",
    "MSDS",
    "KOSHA",
}

CHEM_NAME_KEYS = {
    "chemnamekor",
    "chemnmkor",
    "chemname",
    "chemnm",
    "chemicalname",
    "msdsname",
    "prbtnm",
    "prdtnm",
    "물질명",
    "화학물질명",
    "국문명",
}


def normalize_cas(text: str) -> str:
    return FLEX_CAS_RE.sub(r"\1-\2-\3", text or "")


def is_valid_chem_name(val: str, cas: str = "") -> bool:
    s = re.sub(r"\s+", " ", str(val or "")).strip()
    if not s or len(s) < 2:
        return False
    up = s.upper()
    if up in INVALID_CHEM_NAMES or up.startswith("NORMAL SERVICE"):
        return False
    if cas and s == cas:
        return False
    if CAS_RE.fullmatch(s):
        return False
    if re.fullmatch(r"[0-9.\s~-]+", s):
        return False
    return bool(re.search(r"[가-힣A-Za-z]", s))


def clean_name(raw: str, cas: str = "") -> str:
    s = re.sub(r"/\s*KE-\d+", "", str(raw or ""), flags=re.I)
    s = re.sub(r"\s+", " ", s).strip(" |,:;-/")
    if not is_valid_chem_name(s, cas):
        return ""
    return s[:150]


def extract_pdf_items(data: bytes) -> list[dict]:
    doc = fitz.open(stream=data, filetype="pdf")
    pages_text = []
    for page in doc:
        txt = page.get_text("text") or ""
        if len(txt.strip()) < 25:
            pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            try:
                ocr = pytesseract.image_to_string(img, lang="kor+eng")
            except Exception:
                ocr = pytesseract.image_to_string(img, lang="eng")
            txt = f"{txt}\n{ocr}".strip()
        pages_text.append(normalize_cas(txt))

    full_text = "\n".join(pages_text)
    lines = [re.sub(r"\s+", " ", ln).strip() for ln in full_text.splitlines() if ln.strip()]

    items = []
    seen = set()
    for idx, line in enumerate(lines):
        for m in CAS_RE.finditer(line):
            cas = m.group(1)
            if cas in seen:
                continue
            seen.add(cas)
            before = line[: m.start()].strip()
            name = clean_name(before, cas)
            if not name and idx > 0:
                prev = lines[idx - 1]
                if not CAS_RE.search(prev):
                    name = clean_name(prev, cas)
            items.append({
                "original_name": name,
                "cas": cas,
            })
    return items


def xml_to_dict(node) -> dict:
    out = {}
    for child in list(node):
        tag = child.tag.split("}")[-1]
        val = xml_to_dict(child) if list(child) else (child.text or "").strip()
        if tag in out:
            if not isinstance(out[tag], list):
                out[tag] = [out[tag]]
            out[tag].append(val)
        else:
            out[tag] = val
    return out


def flatten_strings(obj) -> list[str]:
    res = []
    if isinstance(obj, dict):
        for v in obj.values():
            res.extend(flatten_strings(v))
    elif isinstance(obj, list):
        for v in obj:
            res.extend(flatten_strings(v))
    elif obj is not None:
        s = re.sub(r"\s+", " ", str(obj)).strip()
        if s and not s.upper().startswith("NORMAL SERVICE"):
            res.append(s)
    return res


def extract_named_chem_from_xml(obj, cas: str) -> str:
    found = []

    def walk(node):
        if isinstance(node, dict):
            for k, v in node.items():
                kl = str(k).lower()
                if kl in CHEM_NAME_KEYS or any(t in kl for t in ["chemname", "chemnm", "물질명", "국문명"]):
                    for s in flatten_strings(v):
                        cleaned = clean_name(s, cas)
                        if cleaned:
                            found.append(cleaned)
                walk(v)
        elif isinstance(node, list):
            for item in node:
                walk(item)

    walk(obj)
    for cand in found:
        if re.search(r"[가-힣]", cand):
            return cand
    return found[0] if found else ""


def kosha_api_get(operation: str, params: dict) -> dict:
    if not SERVICE_KEY:
        return {}
    try:
        r = requests.get(f"{BASE_API}/{operation}", params={"serviceKey": SERVICE_KEY, **params}, timeout=15)
        r.raise_for_status()
        root = ET.fromstring(r.content)
        return xml_to_dict(root)
    except Exception:
        return {}


def parse_kosha_web(cas: str) -> dict:
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer": f"{KOSHA_WEB_BASE}/",
    }
    info = {"chem_name": "", "text_blobs": []}
    try:
        r = requests.get(f"{KOSHA_WEB_BASE}/search/totalSearch.do", params={"searchWrd": cas}, headers=headers, timeout=12)
        if r.status_code != 200 or not r.text:
            return info
        soup = BeautifulSoup(r.text, "html.parser")
        all_txt = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))
        info["text_blobs"].append(all_txt)

        for tr in soup.select("table tr"):
            row_txt = re.sub(r"\s+", " ", tr.get_text(" ", strip=True))
            if cas in row_txt:
                cells = [re.sub(r"\s+", " ", td.get_text(" ", strip=True)) for td in tr.select("td")]
                for c in cells:
                    cleaned = clean_name(c, cas)
                    if cleaned:
                        info["chem_name"] = cleaned
                        break
                if info["chem_name"]:
                    break
        return info
    except Exception:
        return info


def parse_twa_values(combined_text: str) -> tuple[str, str]:
    ppm_match = re.search(
        r"(?:TWA\s*[:：-]?\s*)?([0-9]+(?:\.[0-9]+)?)\s*ppm\b",
        combined_text,
        re.I,
    )
    if ppm_match:
        return ppm_match.group(1), ""

    mg_match = re.search(
        r"(?:TWA\s*[:：-]?\s*)?([0-9]+(?:\.[0-9]+)?)\s*(?:mg/m3|mg/㎥|㎎/㎥|mg)\b",
        combined_text,
        re.I,
    )
    if mg_match:
        return "", mg_match.group(1)

    return "", ""


def parse_h_codes_and_grade(combined_text: str) -> tuple[str, str]:
    codes = sorted(set(H_CODE_RE.findall(combined_text.upper())))
    if not codes:
        return "", ""
    max_grade = max([H_GRADE_MAP.get(c, 0) for c in codes], default=0)
    grade_str = str(max_grade) if max_grade > 0 else ""
    return ", ".join(codes), grade_str


def _extract_cmr_cat(text: str, keywords: list[str]) -> str:
    for kw in keywords:
        for m in re.finditer(re.escape(kw), text, re.I):
            window = text[max(0, m.start() - 25) : min(len(text), m.end() + 60)]
            if "수유독성" in window or "수유에 독성" in window or re.search(r"수유\s*독성", window):
                return "수유독성"
            cat = re.search(r"(?:구분\s*)?(1A|1B|2|1)\b", window, re.I)
            if cat:
                return cat.group(1).upper()
    return ""


def parse_cmr(combined_text: str, h_codes_str: str) -> tuple[str, str, str]:
    c_val = _extract_cmr_cat(combined_text, ["발암성", "carcinogenicity"])
    m_val = _extract_cmr_cat(combined_text, ["생식세포 변이원성", "생식세포변이원성", "변이원성", "mutagenicity"])
    r_val = _extract_cmr_cat(combined_text, ["생식독성", "reproductive toxicity"])

    if not c_val and any(x in h_codes_str for x in ["H350", "H351"]):
        c_val = "2"
    if not m_val and any(x in h_codes_str for x in ["H340", "H341"]):
        m_val = "2"
    if not r_val:
        if "H362" in h_codes_str:
            r_val = "수유독성"
        elif any(x in h_codes_str for x in ["H360", "H361"]):
            r_val = "2"

    return c_val, m_val, r_val


def parse_boiling_point(combined_text: str) -> str:
    if not combined_text:
        return ""

    patterns = [
        r"초기\s*끓는점과\s*끓는점\s*범위",
        r"초기\s*끓는점",
        r"끓는점\s*범위",
        r"초기끓는점",
        r"끓는점",
    ]

    for pat in patterns:
        for m in re.finditer(pat, combined_text, re.I):
            after = combined_text[m.end() :]
            deg_m = re.search(r"℃", after)
            if not deg_m:
                continue
            segment = after[: deg_m.start()]

            range_m = re.search(
                r"(-?\d+(?:\.\d+)?)\s*(?:~|-|–|—|to)\s*(-?\d+(?:\.\d+)?)",
                segment,
                re.I,
            )
            if range_m:
                v1 = float(range_m.group(1))
                v2 = float(range_m.group(2))
                low = min(v1, v2)
                return str(int(low)) if low.is_integer() else str(low)

            single_m = re.search(r"(-?\d+(?:\.\d+)?)", segment)
            if single_m:
                val = float(single_m.group(1))
                return str(int(val)) if val.is_integer() else str(val)

    return ""


@app.get("/", response_class=HTMLResponse)
def index():
    with open("app/static/index.html", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/api/health")
def health():
    return JSONResponse(content={"ok": True, "kosha_key_configured": bool(SERVICE_KEY)})


@app.post("/api/extract")
async def extract_pdf(file: UploadFile = File(...)):
    if not file or not (file.filename or "").lower().endswith(".pdf"):
        return JSONResponse(status_code=400, content={"detail": "PDF 파일만 업로드해 주세요."})
    data = await file.read()
    if not data:
        return JSONResponse(status_code=400, content={"detail": "빈 파일입니다."})
    items = extract_pdf_items(data)
    return JSONResponse(content={"items": items})


@app.get("/api/lookup/{cas}")
def lookup_cas(cas: str):
    if not CAS_RE.fullmatch(cas):
        return JSONResponse(status_code=400, content={"detail": "CAS번호 형식이 올바르지 않습니다."})

    web_info = parse_kosha_web(cas)
    blobs = list(web_info.get("text_blobs") or [])
    chem_name = clean_name(web_info.get("chem_name") or "", cas)

    listing = kosha_api_get("getChemList001", {"searchWrd": cas, "searchCnd": "1", "numOfRows": "10", "pageNo": "1"})
    if not chem_name:
        chem_name = extract_named_chem_from_xml(listing, cas)
    blobs.extend(flatten_strings(listing))

    chem_id = ""
    for s in flatten_strings(listing):
        if re.fullmatch(r"[A-Za-z0-9_-]{4,30}", s) and s != cas and not s.upper().startswith("NORMAL"):
            chem_id = s
            break

    details = {}
    if chem_id:
        for op in ["getChemDetail011", "getChemDetail021", "getChemDetail031", "getChemDetail081", "getChemDetail091", "getChemDetail151"]:
            detail = kosha_api_get(op, {"chemId": chem_id})
            details[op] = detail
            blobs.extend(flatten_strings(detail))
            if not chem_name:
                chem_name = extract_named_chem_from_xml(detail, cas)

    if not chem_name:
        for s in blobs:
            cleaned = clean_name(s, cas)
            if cleaned:
                chem_name = cleaned
                break

    physical_text = " ".join(flatten_strings(details.get("getChemDetail091", {})))
    combined_text = " ".join(blobs)

    twa_ppm, twa_mg = parse_twa_values(combined_text)
    h_codes, h_grade = parse_h_codes_and_grade(combined_text)
    cmr_c, cmr_m, cmr_r = parse_cmr(combined_text, h_codes)

    boiling_point = ""
    if twa_ppm.strip():
        boiling_point = parse_boiling_point(physical_text) or parse_boiling_point(combined_text)

    return JSONResponse(content={
        "chem_name": chem_name,
        "cas": cas,
        "twa_ppm": twa_ppm,
        "twa_mg": twa_mg,
        "h_code": h_codes,
        "h_grade": h_grade,
        "cmr_c": cmr_c,
        "cmr_m": cmr_m,
        "cmr_r": cmr_r,
        "boiling_point": boiling_point,
    })


EXPORT_HEADERS = [
    "공정명",
    "제품명",
    "원본 MSDS 물질명",
    "화학물질명(필수)",
    "CAS번호",
    "TWA(ppm)값",
    "TWA(mg)값",
    "H-CODE값",
    "H-CODE등급",
    "CMR_C값",
    "CMR_M값",
    "CMR_R값",
    "끓는점",
]

EXPORT_KEYS = [
    "process",
    "product",
    "original_name",
    "chem_name",
    "cas",
    "twa_ppm",
    "twa_mg",
    "h_code",
    "h_grade",
    "cmr_c",
    "cmr_m",
    "cmr_r",
    "boiling_point",
]


@app.post("/api/export/xlsx")
async def export_xlsx(payload: dict):
    rows = (payload or {}).get("rows", [])
    wb = Workbook()
    ws = wb.active
    ws.title = "MSDS 조회결과"
    ws.append(EXPORT_HEADERS)
    for r in rows:
        ws.append([str((r or {}).get(k, "") or "") for k in EXPORT_KEYS])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="18384B")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for col in ws.columns:
        letter = col[0].column_letter
        ws.column_dimensions[letter].width = min(max(max(len(str(c.value or "")) for c in col) + 2, 12), 36)
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=msds_result.xlsx"},
    )


@app.post("/api/export/pdf")
async def export_pdf(payload: dict):
    rows = (payload or {}).get("rows", [])
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=landscape(A4), rightMargin=18, leftMargin=18, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    font_name = "Helvetica"
    for path in ["/usr/share/fonts/truetype/nanum/NanumGothic.ttf", "/usr/share/fonts/truetype/nanum/NanumGothicCoding.ttf"]:
        if os.path.exists(path):
            pdfmetrics.registerFont(TTFont("Korean", path))
            font_name = "Korean"
            break
    style = ParagraphStyle("cell", parent=styles["BodyText"], fontName=font_name, fontSize=7, leading=8)
    data = [[Paragraph(h, style) for h in EXPORT_HEADERS]]
    for r in rows:
        data.append([
            Paragraph(str((r or {}).get(k, "") or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"), style)
            for k in EXPORT_KEYS
        ])
    widths = [55, 60, 95, 95, 58, 48, 48, 65, 45, 42, 42, 48, 45]
    t = Table(data, colWidths=widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#18384B")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    doc.build([Paragraph("MSDS 유해화학물질 조회 결과", styles["Title"]), Spacer(1, 8), t])
    bio.seek(0)
    return StreamingResponse(
        bio,
        media_type="application/pdf",
        headers={"Content-Disposition": "attachment; filename=msds_result.pdf"},
    )
