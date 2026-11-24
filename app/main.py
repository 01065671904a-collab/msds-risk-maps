import os, re, io, time, threading, xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import date
import requests
import fitz
import pytesseract
from PIL import Image
from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

app = FastAPI(title="MSDS 유해화학물질 통합관리")
app.mount("/static", StaticFiles(directory="app/static"), name="static")

BASE = "https://apis.data.go.kr/B552468/msdschem1"
KEY = os.getenv("KOSHA_SERVICE_KEY", "").strip()
_CACHE = {}
_CACHE_LOCK = threading.Lock()
_IP_CALLS = defaultdict(list)
_DAILY_LOOKUPS = {"day": date.today().isoformat(), "count": 0}
MAX_NEW_CAS_LOOKUPS_PER_DAY = int(os.getenv("MAX_NEW_CAS_LOOKUPS_PER_DAY", "250"))
CAS_RE = re.compile(r"(?<!\d)(\d{2,7}-\d{2}-\d)(?!\d)")
FLEX_CAS_RE = re.compile(r"(?<!\d)(\d{2,7})\s*[-–—‐‑‒−]\s*(\d{2})\s*[-–—‐‑‒−]\s*(\d)(?!\d)")
H_RE = re.compile(r"\bH(?:300|301|302|304|310|311|312|314|315|317|318|319|320|330|331|332|334|335|336|340|341|350|351|360|361|362|370|371|372|373)\b")

# 4등급 > 3등급 > 2등급 > 1등급
H_LEVEL = {
    "H340": 4, "H341": 4, "H350": 4, "H351": 4, "H360": 4, "H361": 4, "H362": 4, "H370": 4,
    "H300": 3, "H301": 3, "H310": 3, "H311": 3, "H330": 3, "H331": 3,
    "H314": 3, "H318": 3, "H317": 3, "H334": 3, "H372": 3, "H373": 3,
    "H302": 2, "H312": 2, "H332": 2, "H304": 2, "H371": 2,
    "H315": 1, "H319": 1, "H320": 1, "H335": 1, "H336": 1,
}

NOISE_TOKENS = {"00", "NORMAL SERVICE.", "NORMAL SERVICE", "NULL", "NONE", "-", "N/A"}


def xml_dict(node):
    out = {}
    for child in list(node):
        tag = child.tag.split("}")[-1]
        if list(child):
            val = xml_dict(child)
        else:
            val = (child.text or "").strip()
        if tag in out:
            if not isinstance(out[tag], list):
                out[tag] = [out[tag]]
            out[tag].append(val)
        else:
            out[tag] = val
    return out


def flatten_text(obj):
    vals = []
    if isinstance(obj, dict):
        for v in obj.values():
            vals.extend(flatten_text(v))
    elif isinstance(obj, list):
        for v in obj:
            vals.extend(flatten_text(v))
    elif obj is not None:
        s = re.sub(r"\s+", " ", str(obj)).strip()
        if s:
            vals.append(s)
    return vals


def clean_values(vals, allow_no_data=False):
    out = []
    for v in vals:
        s = re.sub(r"\s+", " ", str(v or "")).strip()
        if not s or s.upper() in NOISE_TOKENS:
            continue
        if not allow_no_data and s in ("자료없음", "해당없음"):
            continue
        out.append(s)
    return list(dict.fromkeys(out))


def parse_kosha_xml(content):
    try:
        root = ET.fromstring(content)
        return xml_dict(root)
    except Exception:
        return {"raw": content.decode("utf-8", "ignore")[:12000]}


def api_get(operation, params):
    if not KEY:
        raise HTTPException(400, "KOSHA_SERVICE_KEY가 설정되지 않았습니다. Render 환경변수에 공공데이터포털 인증키를 등록하세요.")
    q = {"serviceKey": KEY, **params}
    try:
        r = requests.get(f"{BASE}/{operation}", params=q, timeout=25)
        r.raise_for_status()
        raw = r.text
        error_codes = {
            "SERVICE_KEY_IS_NOT_REGISTERED_ERROR": "등록되지 않은 인증키입니다.",
            "SERVICE_ACCESS_DENIED_ERROR": "해당 API의 승인·이용 권한을 확인해 주세요.",
            "LIMITED_NUMBER_OF_SERVICE_REQUESTS_EXCEEDS_ERROR": "일일 API 호출 한도를 초과했습니다.",
            "LIMITED_NUMBER_OF_SERVICE_REQUESTS_PER_SECOND_EXCEEDS_ERROR": "API 초당 호출 한도를 초과했습니다. 잠시 후 다시 시도하세요.",
            "DEADLINE_HAS_EXPIRED_ERROR": "API 인증키 사용기간이 만료되었습니다.",
            "INVALID_REQUEST_PARAMETER_ERROR": "API 요청변수 형식이 올바르지 않습니다.",
        }
        for code, message in error_codes.items():
            if code in raw:
                raise HTTPException(502, f"KOSHA API 오류({code}): {message}")
        return parse_kosha_xml(r.content)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(502, f"KOSHA API 호출 실패: {e}")


def protect_public_kosha_request(client_ip, cas):
    now = time.time()
    recent = [t for t in _IP_CALLS[client_ip] if now - t < 60]
    if len(recent) >= 30:
        raise HTTPException(429, "요청이 너무 많습니다. 1분 후 다시 시도해 주세요.")
    recent.append(now)
    _IP_CALLS[client_ip] = recent
    today = date.today().isoformat()
    if _DAILY_LOOKUPS["day"] != today:
        _DAILY_LOOKUPS.update({"day": today, "count": 0})
    with _CACHE_LOCK:
        if cas in _CACHE:
            return _CACHE[cas]
        if _DAILY_LOOKUPS["count"] >= MAX_NEW_CAS_LOOKUPS_PER_DAY:
            raise HTTPException(429, "오늘의 공개 KOSHA 조회 한도에 도달했습니다. 내일 다시 이용해 주세요.")
        _DAILY_LOOKUPS["count"] += 1
    return None


def find_values(obj, keys):
    found = []
    lower_keys = {x.lower() for x in keys}
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k.lower() in lower_keys:
                found.extend(flatten_text(v))
            found.extend(find_values(v, keys))
    elif isinstance(obj, list):
        for v in obj:
            found.extend(find_values(v, keys))
    return found


def first_value(obj, keys, allow_no_data=False):
    vals = clean_values(find_values(obj, keys), allow_no_data=allow_no_data)
    return " / ".join(vals)


def extract_chem_name(listing, detail01, detail03, cas):
    candidate_keys = [
        "chemNameKor", "chemNmKor", "chemName", "chemNm", "chemicalName",
        "msdsName", "prdtNm", "물질명", "화학물질명", "국문명", "제품명"
    ]

    def search_items(node):
        if isinstance(node, dict):
            vals_str = " ".join(flatten_text(node))
            if cas in vals_str:
                for k in candidate_keys:
                    for nk, nv in node.items():
                        if nk.lower() == k.lower():
                            c = clean_values(flatten_text(nv))
                            for item in c:
                                if item != cas and not CAS_RE.fullmatch(item):
                                    return item
            for v in node.values():
                res = search_items(v)
                if res:
                    return res
        elif isinstance(node, list):
            for v in node:
                res = search_items(v)
                if res:
                    return res
        return None

    for src in (listing, detail01, detail03):
        hit = search_items(src)
        if hit:
            return hit

    for src in (listing, detail01, detail03):
        vals = clean_values(find_values(src, candidate_keys))
        for v in vals:
            if v != cas and not CAS_RE.fullmatch(v) and "MSDS" not in v.upper():
                return v

    def search_label_value(node):
        if isinstance(node, dict):
            flat = clean_values(flatten_text(node))
            for idx, txt in enumerate(flat):
                if any(lbl in txt for lbl in ["제품명", "화학물질명", "물질명"]) and idx + 1 < len(flat):
                    nxt = flat[idx + 1]
                    if nxt != cas and not CAS_RE.fullmatch(nxt):
                        return nxt
            for v in node.values():
                res = search_label_value(v)
                if res:
                    return res
        elif isinstance(node, list):
            for v in node:
                res = search_label_value(v)
                if res:
                    return res
        return None

    for src in (detail01, detail03):
        hit = search_label_value(src)
        if hit:
            return hit

    return "미확인"


def extract_exposure_text(detail_obj):
    domestic_texts = []

    def collect_domestic(node):
        if isinstance(node, dict):
            flat = clean_values(flatten_text(node))
            joined = " ".join(flat)
            if "국내규정" in joined:
                for idx, txt in enumerate(flat):
                    if "국내규정" in txt:
                        domestic_texts.extend(flat[idx : idx + 4])
                return
            for k, v in node.items():
                if "국내규정" in str(k):
                    domestic_texts.extend(clean_values(flatten_text(v)))
                else:
                    collect_domestic(v)
        elif isinstance(node, list):
            for v in node:
                collect_domestic(v)

    collect_domestic(detail_obj)
    source_texts = domestic_texts if domestic_texts else clean_values(flatten_text(detail_obj))

    filtered_chunks = []
    for t in source_texts:
        if any(foreign in t.upper() for foreign in ["ACGIH", "OSHA", "NIOSH", "국외규정", "생물학적 노출기준"]):
            break
        filtered_chunks.append(t)

    combined = " \n ".join(filtered_chunks)
    if not combined or ("자료없음" in combined and not re.search(r"(TWA|STEL|CEILING|(?<![A-Z])C(?![A-Z]))", combined, re.I)):
        return "미확인"

    found = {}
    patterns = [
        ("TWA", r"TWA\s*[:：-]?\s*([0-9]+(?:\.[0-9]+)?\s*(?:ppm|mg/m3|mg/㎥|㎎/㎥))"),
        ("STEL", r"STEL\s*[:：-]?\s*([0-9]+(?:\.[0-9]+)?\s*(?:ppm|mg/m3|mg/㎥|㎎/㎥))"),
        ("C", r"(?:(?<![A-Z0-9])C|CEILING|최고노출기준)\s*[:：-]?\s*([0-9]+(?:\.[0-9]+)?\s*(?:ppm|mg/m3|mg/㎥|㎎/㎥))"),
    ]
    for label, pat in patterns:
        m = re.search(pat, combined, re.I)
        if m:
            val = re.sub(r"\s+", "", m.group(1))
            val = val.replace("mg/m3", "mg/㎥").replace("㎎/㎥", "mg/㎥")
            found[label] = f"{label} : {val}"

    ordered = [found[k] for k in ["TWA", "STEL", "C"] if k in found]
    return " / ".join(ordered) if ordered else "미확인"


def extract_boiling_text(detail_obj):
    bp_texts = []

    def collect_bp(node):
        if isinstance(node, dict):
            flat = clean_values(flatten_text(node), allow_no_data=True)
            for idx, txt in enumerate(flat):
                if any(k in txt for k in ["초기 끓는점", "끓는점 범위", "끓는점", "초기끓는점"]):
                    bp_texts.extend(flat[idx : idx + 3])
                    return
            for k, v in node.items():
                if any(tag in str(k).lower() for tag in ["끓는점", "boiling"]):
                    bp_texts.extend(clean_values(flatten_text(v), allow_no_data=True))
                else:
                    collect_bp(v)
        elif isinstance(node, list):
            for v in node:
                collect_bp(v)

    collect_bp(detail_obj)
    if not bp_texts:
        bp_texts = clean_values(flatten_text(detail_obj), allow_no_data=True)

    joined = " ".join(bp_texts)
    if not joined:
        return "미확인"

    range_m = re.search(
        r"(-?\d+(?:\.\d+)?)\s*(?:~|-|–|to)\s*(-?\d+(?:\.\d+)?)\s*(?:°\s*C|℃)",
        joined,
        re.I,
    )
    if range_m:
        v1 = float(range_m.group(1))
        v2 = float(range_m.group(2))
        low = min(v1, v2)
        if low < 0:
            return "미확인"
        return str(int(low)) if low.is_integer() else str(low)

    single_m = re.search(r"(-?\d+(?:\.\d+)?)\s*(?:°\s*C|℃)", joined, re.I)
    if single_m:
        val = float(single_m.group(1))
        if val < 0:
            return "미확인"
        return str(int(val)) if val.is_integer() else str(val)

    if "자료없음" in joined:
        return "자료없음"

    return "미확인"


def industrial_safety_law_text(obj):
    hits = []

    def walk(node):
        if isinstance(node, dict):
            subtree = " ".join(flatten_text(node))
            if "산업안전보건법" in subtree:
                child_hits = [v for v in node.values() if "산업안전보건법" in " ".join(flatten_text(v))]
                if child_hits:
                    for v in child_hits:
                        walk(v)
                else:
                    s = re.sub(r"\s+", " ", " | ".join(clean_values(flatten_text(node)))).strip()
                    if s:
                        hits.append(s)
            else:
                for v in node.values():
                    walk(v)
        elif isinstance(node, list):
            for v in node:
                walk(v)
        elif node is not None and "산업안전보건법" in str(node):
            s = re.sub(r"\s+", " ", str(node)).strip()
            if s:
                hits.append(s)

    walk(obj)
    return " / ".join(dict.fromkeys(hits))


def infer_law_flags(reg_text, full_reg_obj=None):
    combined = " ".join([reg_text or "", " ".join(clean_values(flatten_text(full_reg_obj or {})))])
    compact = re.sub(r"\s+", "", combined)
    if not compact or compact == "미확인":
        return {
            "work_measurement": "미확인",
            "special_health_exam": "미확인",
            "managed_substance": "미확인",
            "specially_managed": "미확인",
        }
    return {
        "work_measurement": "Y" if ("작업환경측정" in compact or "측정대상" in compact) else "N",
        "special_health_exam": "Y" if ("특수건강진단" in compact or "특수검진" in compact) else "N",
        "managed_substance": "Y" if ("관리대상유해물질" in compact or "관리대상물질" in compact) else "N",
        "specially_managed": "Y" if "특별관리물질" in compact else "N",
    }


def normalize_cas_text(text):
    return FLEX_CAS_RE.sub(r"\1-\2-\3", text or "")


def _clean_chem_name(raw_name):
    s = re.sub(r"/\s*KE-\d+", "", str(raw_name or ""), flags=re.I)
    s = re.sub(
        r"^(?:화학물질명|물질명|관용명\s*및\s*이명(?:\(異名\))?|CAS\s*번호\s*또는\s*식별번호|함유량\s*\(?%?\)?)[\s|:.-]*",
        "",
        s,
        flags=re.I,
    )
    s = re.sub(r"\s+", " ", s).strip(" |,:;-/")
    return s[:160] if s else "미확인"


def _parse_pct_from_text(text):
    s = str(text or "").strip()
    if not s:
        return ""
    pm = re.search(r"(\d+(?:\.\d+)?\s*(?:~|[-–]\s*)?\s*\d*(?:\.\d+)?\s*%)", s)
    if pm:
        return re.sub(r"\s+", "", pm.group(1))
    pm2 = re.search(r"(\d+(?:\.\d+)?\s*(?:~|[-–]|\s+)\s*\d+(?:\.\d+)?)", s)
    if pm2:
        return re.sub(r"\s+", "~", pm2.group(1).strip()) + "%"
    pm3 = re.search(r"^\s*(\d+(?:\.\d+)?)\s*$", s)
    if pm3:
        return pm3.group(1) + "%"
    return ""


def _make_ingredient_item(name, cas, pct):
    return {
        "selected": True,
        "process": "",
        "product": "",
        "original_name": _clean_chem_name(name),
        "cas": cas,
        "content": pct or "미확인",
        "kosha_name": "미확인",
        "cmr": "",
        "exposure_limit": "미확인",
        "boiling_point": "미확인",
        "work_measurement": "미확인",
        "special_health_exam": "미확인",
        "managed_substance": "미확인",
        "specially_managed": "미확인",
        "h_codes": "",
        "grade": "미확인",
        "kosha_status": "조회 전",
        "legal_basis": "미확인",
    }


def extract_pdf_ingredients_and_text(data):
    """
    PDF의 표 구조(find_tables)와 단어 좌표(get_text('words'))를 이용해
    열 순서가 뒤섞이는 PDF에서도 CAS No., 물질명, 함유량을 정확히 추출합니다.
    """
    doc = fitz.open(stream=data, filetype="pdf")
    chunks = []
    structured_items = []
    seen_cas = set()

    for page in doc:
        # 1) PyMuPDF 표 탐지(find_tables) 기반 추출 (가장 정확)
        try:
            tabs = page.find_tables()
            for tb in getattr(tabs, "tables", []):
                for row in tb.extract():
                    cells = [normalize_cas_text(re.sub(r"\s+", " ", (c or "")).strip()) for c in row]
                    if not any(cells):
                        continue
                    chunks.append(" | ".join(cells))
                    cas_col_idx = None
                    cas_val = None
                    for idx, c in enumerate(cells):
                        m = CAS_RE.search(c)
                        if m:
                            cas_col_idx = idx
                            cas_val = m.group(1)
                            break
                    if cas_val and cas_val not in seen_cas:
                        # 물질명: CAS 열 이전의 첫 번째 유효 텍스트 셀
                        name_val = ""
                        for idx in range(cas_col_idx):
                            cand = cells[idx]
                            if cand and not any(h in cand for h in ["화학물질명", "물질명", "관용명", "식별번호"]):
                                name_val = cand
                                break
                        # 함유량: CAS 열 이후의 첫 번째 숫자/범위 셀
                        pct_val = ""
                        for idx in range(cas_col_idx + 1, len(cells)):
                            pct_val = _parse_pct_from_text(cells[idx])
                            if pct_val:
                                break
                        seen_cas.add(cas_val)
                        structured_items.append(_make_ingredient_item(name_val, cas_val, pct_val))
        except Exception:
            pass

        # 2) 단어 좌표(words) 기반 Y축 행 정렬 추출 (표 테두리가 없거나 find_tables가 놓친 경우 보완)
        word_lines = []
        try:
            words = page.get_text("words") or []
            if words:
                # 같은 행(Y좌표 차이 8pt 이내)으로 묶은 뒤 X좌표 오름차순 정렬
                words.sort(key=lambda w: (w[1], w[0]))
                rows_y = []
                for w in words:
                    if not rows_y or abs(w[1] - rows_y[-1]["y"]) > 8.0:
                        rows_y.append({"y": w[1], "words": [w]})
                    else:
                        rows_y[-1]["words"].append(w)

                # 표 헤더 위치(화학물질명 / 관용명 / CAS / 함유량 X좌표) 탐지
                col_bounds = None
                for r in rows_y:
                    r["words"].sort(key=lambda x: x[0])
                    line_txt = normalize_cas_text(" ".join(x[4] for x in r["words"]))
                    word_lines.append(line_txt)
                    if "CAS" in line_txt.upper() and ("함유량" in line_txt or "화학물질명" in line_txt or "물질명" in line_txt):
                        cas_x = next((x[0] for x in r["words"] if "CAS" in x[4].upper()), None)
                        pct_x = next((x[0] for x in r["words"] if "함유량" in x[4]), None)
                        alias_x = next((x[0] for x in r["words"] if "관용명" in x[4] or "이명" in x[4]), None)
                        if cas_x:
                            col_bounds = {"alias_x": alias_x, "cas_x": cas_x, "pct_x": pct_x}

                for r in rows_y:
                    line_txt = normalize_cas_text(" ".join(x[4] for x in r["words"]))
                    m = CAS_RE.search(line_txt)
                    if not m:
                        continue
                    cas_val = m.group(1)
                    if cas_val in seen_cas:
                        continue

                    if col_bounds:
                        name_words = [
                            x[4] for x in r["words"]
                            if x[0] < (col_bounds["alias_x"] or col_bounds["cas_x"] - 20)
                        ]
                        pct_words = [
                            x[4] for x in r["words"]
                            if x[0] >= ((col_bounds["pct_x"] - 25) if col_bounds["pct_x"] else (col_bounds["cas_x"] + 90))
                        ]
                        name_val = " ".join(name_words)
                        pct_val = _parse_pct_from_text(" ".join(pct_words))
                    else:
                        cas_word_x = next((x[0] for x in r["words"] if CAS_RE.search(normalize_cas_text(x[4]))), 300)
                        name_val = " ".join(x[4] for x in r["words"] if x[0] < cas_word_x - 10)
                        pct_val = _parse_pct_from_text(" ".join(x[4] for x in r["words"] if x[0] > cas_word_x + 20))

                    seen_cas.add(cas_val)
                    structured_items.append(_make_ingredient_item(name_val, cas_val, pct_val))
        except Exception:
            pass

        txt = page.get_text("text") or ""
        page_combined = "\n".join([x for x in ["\n".join(word_lines), txt] if x.strip()])

        # 3) 스캔 이미지 PDF 등 텍스트가 거의 없으면 OCR 수행
        if len(txt.strip()) < 30:
            pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5), alpha=False)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            try:
                ocr_txt = pytesseract.image_to_string(img, lang="kor+eng")
            except Exception:
                ocr_txt = pytesseract.image_to_string(img, lang="eng")
            page_combined = (page_combined + "\n" + ocr_txt).strip()

        chunks.append(page_combined)

    full_text = "\n".join(chunks)
    if not structured_items:
        structured_items = parse_ingredients_fallback(full_text)
    return structured_items, full_text


def parse_ingredients_fallback(text):
    """좌표/표 추출로 못 잡힌 경우(스캔 OCR 등)를 위한 텍스트 기반 폴백 파서"""
    norm = normalize_cas_text(text)
    lines = [re.sub(r"\s+", " ", x).strip() for x in norm.splitlines() if x.strip()]
    found = []
    seen = set()
    for i, line in enumerate(lines):
        for m in CAS_RE.finditer(line):
            cas = m.group(1)
            if cas in seen:
                continue
            window = [
                lines[j] for j in range(max(0, i - 2), min(len(lines), i + 3))
                if j == i or not CAS_RE.search(lines[j])
            ]
            row = " ".join(window)
            pct = _parse_pct_from_text(row)
            name_part = line[:m.start()].strip() or (lines[i - 1] if i > 0 and not CAS_RE.search(lines[i - 1]) else "")
            seen.add(cas)
            found.append(_make_ingredient_item(name_part, cas, pct))
    return found


@app.get("/", response_class=HTMLResponse)
def home():
    with open("app/static/index.html", encoding="utf-8") as f:
        return f.read()


@app.get("/api/health")
def health():
    return {
        "ok": True,
        "kosha_key_configured": bool(KEY),
        "anonymous_access": True,
        "daily_lookup_limit": MAX_NEW_CAS_LOOKUPS_PER_