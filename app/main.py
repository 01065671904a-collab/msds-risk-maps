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
    """KOSHA 검색 결과에서 물질명(국문명 우선)을 정확히 추출합니다."""
    candidate_keys = [
        "chemNameKor", "chemNmKor", "chemName", "chemNm", "chemicalName",
        "msdsName", "prdtNm", "물질명", "화학물질명", "국문명", "제품명"
    ]
    # 1. listing items 중 CAS가 일치하는 항목 우선 탐색
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

    # 011 상세의 "가. 제품명" 항목 구조 대응 (itemDetail / msdsItemDetail 등)
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
    """국내규정에서 TWA, STEL, C 값만 추출합니다. ACGIH/OSHA 등 국외기준은 제외하고 국내규정 첫 유효값만 사용합니다."""
    domestic_texts = []

    def collect_domestic(node):
        if isinstance(node, dict):
            flat = clean_values(flatten_text(node))
            joined = " ".join(flat)
            if "국내규정" in joined:
                # 국내규정 항목과 그 형제/하위 값 수집
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

    # 해외 기준(ACGIH, OSHA, NIOSH 등) 구간 이전까지만 자르기
    filtered_chunks = []
    for t in source_texts:
        if any(foreign in t.upper() for foreign in ["ACGIH", "OSHA", "NIOSH", "국외규정", "생물학적 노출기준"]):
            break
        filtered_chunks.append(t)

    combined = " \n ".join(filtered_chunks)
    if not combined or "자료없음" in combined and not re.search(r"(TWA|STEL|CEILING|(?<![A-Z])C(?![A-Z]))", combined, re.I):
        return "미확인"

    # 국내규정 텍스트에서 TWA, STEL, C 각각의 첫 번째 수치+단위(ppm 또는 mg/m3)만 깔끔하게 추출
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
    """
    끓는점 추출 규칙:
    1) '자료없음'이면 '자료없음' 반환
    2) 범위(예: 140~150℃)이면 최소값(140)만 반환
    3) ℃ 앞 숫자만 추출 (단위 제외)
    4) 0℃ 미만(음수, 예: -2℃)이면 '미확인' 반환
    """
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

    # 범위 패턴 우선 탐색 (예: 140~150℃, 159.5 ~ 160 °C, -5~10℃)
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

    # 단일 온도 패턴 탐색 (예: 56.1 ℃, 235 ℃)
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


def extract_pdf_text(data):
    doc = fitz.open(stream=data, filetype="pdf")
    chunks = []
    for page in doc:
        txt = page.get_text("text") or ""
        if len(txt.strip()) < 30:
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            try:
                txt = pytesseract.image_to_string(img, lang="kor+eng")
            except Exception:
                txt = pytesseract.image_to_string(img, lang="eng")
        chunks.append(txt)
    return "\n".join(chunks)


def section_three(text):
    start_patterns = [
        r"3\s*[.\-]\s*구성\s*성분\s*의\s*명칭\s*및\s*함유량",
        r"3\s*[.\-]\s*구성성분의\s*명칭\s*및\s*함유량",
        r"3\s*[.\-]\s*구성성분",
        r"3\s*[.\-]\s*Composition\s*/\s*information",
    ]
    starts = [m.start() for p in start_patterns for m in re.finditer(p, text, re.I)]
    if starts:
        start = min(starts)
        tail = text[start:]
        end = re.search(r"\n\s*4\s*[.\-]\s*(응급조치|First[- ]aid)", tail, re.I)
        return tail[: end.start()] if end else tail[:16000]
    return text


def parse_ingredients(text):
    sec = section_three(text)
    lines = [re.sub(r"\s+", " ", x).strip() for x in sec.splitlines() if x.strip()]
    found = []
    for i, line in enumerate(lines):
        for m in CAS_RE.finditer(line):
            cas = m.group(1)
            row = " ".join(lines[max(0, i - 2) : min(len(lines), i + 3)])
            pct = ""
            pm = re.search(r"(\d+(?:\.\d+)?\s*(?:~|[-–]\s*)?\s*\d*(?:\.\d+)?\s*%)", row)
            if pm:
                pct = pm.group(1)
            clean = row.replace(cas, " ")
            clean = re.sub(r"\s+", " ", clean).strip(" |,:;")
            clean = re.sub(r"(화학물질명|물질명|CAS\s*No\.?|함유량|중량\s*%)", "", clean, flags=re.I).strip()
            name = clean[:160] if clean else "미확인"
            item = {
                "selected": True,
                "process": "",
                "product": "",
                "original_name": name,
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
            if not any(x["cas"] == cas for x in found):
                found.append(item)
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
        "daily_lookup_limit": MAX_NEW_CAS_LOOKUPS_PER_DAY,
    }


@app.post("/api/extract")
async def extract(file: UploadFile = File(...), request: Request = None):
    client_ip = getattr(getattr(request, "client", None), "host", "unknown")
    now = time.time()
    uploads = [t for t in _IP_CALLS["upload:" + client_ip] if now - t < 60]
    if len(uploads) >= 10:
        raise HTTPException(429, "PDF 분석 요청이 많습니다. 1분 후 다시 시도해 주세요.")
    uploads.append(now)
    _IP_CALLS["upload:" + client_ip] = uploads
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "PDF 파일만 업로드할 수 있습니다.")
    data = await file.read()
    if len(data) > 20 * 1024 * 1024:
        raise HTTPException(413, "파일은 20MB 이하로 업로드해 주세요.")
    try:
        txt = extract_pdf_text(data)
        ingredients = parse_ingredients(txt)
        return {
            "filename": file.filename,
            "ingredients": ingredients,
            "text_preview": txt[:1200],
            "ocr_used": len(txt.strip()) > 0,
            "warning": "자동 추출값은 원본 MSDS 3번 항목과 대조해야 합니다.",
        }
    except Exception as e:
        raise HTTPException(400, f"PDF 분석 실패: {e}")


@app.get("/api/kosha/{cas}")
def kosha_lookup(cas: str, request: Request):
    if not CAS_RE.fullmatch(cas):
        raise HTTPException(400, "CAS No. 형식이 올바르지 않습니다.")
    client_ip = getattr(getattr(request, "client", None), "host", "unknown")
    cached = protect_public_kosha_request(client_ip, cas)
    if cached is not None:
        return cached
    listing = api_get("getChemList001", {"searchWrd": cas, "searchCnd": "1", "numOfRows": "20", "pageNo": "1"})
    chem_id = first_value(listing, ["chemId", "chemID", "chemicalId", "화학물질ID"])
    if not chem_id:
        raw = " ".join(clean_values(flatten_text(listing)))
        result = {
            "cas": cas,
            "found": False,
            "kosha_name": "미확인",
            "h_codes": "",
            "grade": "미확인",
            "cmr": "",
            "exposure_limit": "미확인",
            "boiling_point": "미확인",
            "work_measurement": "미확인",
            "special_health_exam": "미확인",
            "managed_substance": "미확인",
            "specially_managed": "미확인",
            "regulation_text": "미확인",
            "status": "KOSHA 목록에서 검색되지 않음",
            "raw": raw[:2500],
        }
        with _CACHE_LOCK:
            _CACHE[cas] = result
        return result

    details = {}
    for code, op in [
        ("product", "getChemDetail011"),
        ("hazard", "getChemDetail021"),
        ("ingredients", "getChemDetail031"),
        ("exposure", "getChemDetail081"),
        ("physical", "getChemDetail091"),
        ("regulation", "getChemDetail151"),
    ]:
        try:
            details[code] = api_get(op, {"chemId": chem_id})
        except Exception as e:
            details[code] = {"error": str(e)}

    chem_name = extract_chem_name(listing, details.get("product", {}), details.get("ingredients", {}), cas)

    hazard_text = " ".join(flatten_text(details.get("hazard", {})))
    hcodes = sorted(set(H_RE.findall(hazard_text)))
    level = max([H_LEVEL.get(x, 0) for x in hcodes], default=0)
    grade_str = f"{level}등급" if level else "미확인"

    exposure = extract_exposure_text(details.get("exposure", {}))
    boiling = extract_boiling_text(details.get("physical", {}))

    cmr = []
    if any(x in hcodes for x in ["H340", "H341"]):
        cmr.append("생식세포 변이원성(구분 확인 필요)")
    if any(x in hcodes for x in ["H350", "H351"]):
        cmr.append("발암성(구분 확인 필요)")
    if any(x in hcodes for x in ["H360", "H361", "H362"]):
        cmr.append("생식독성(구분 확인 필요)")

    reg_text = industrial_safety_law_text(details.get("regulation", {}))[:4000] or "미확인"
    flags = infer_law_flags(reg_text, details.get("regulation", {}))

    result = {
        "cas": cas,
        "found": True,
        "chem_id": chem_id,
        "kosha_name": chem_name or "미확인",
        "h_codes": ", ".join(hcodes),
        "grade": grade_str,
        "cmr": ", ".join(cmr),
        "exposure_limit": exposure,
        "boiling_point": boiling,
        "work_measurement": flags["work_measurement"],
        "special_health_exam": flags["special_health_exam"],
        "managed_substance": flags["managed_substance"],
        "specially_managed": flags["specially_managed"],
        "regulation_text": reg_text,
        "raw_details": details,
        "status": "KOSHA 조회 완료",
        "notice": "CMR 구분은 H-code만으로 1A/1B를 확정할 수 없어 원문 확인이 필요합니다. H-code 등급은 제공된 기준표로 산정했습니다.",
    }
    with _CACHE_LOCK:
        _CACHE[cas] = result
    return result


def normalize_rows(rows):
    if not isinstance(rows, list) or len(rows) > 20000:
        raise HTTPException(400, "데이터 형식이 올바르지 않습니다.")
    return rows


@app.post("/api/export/xlsx")
async def export_xlsx(payload: dict):
    rows = normalize_rows(payload.get("rows", []))
    wb = Workbook()
    ws = wb.active
    ws.title = "MSDS 통합관리"
    headers = [
        "공정명", "제품명", "원본 MSDS 물질명", "KOSHA 조회 물질명", "CAS No.", "함유량", "CMR 분류",
        "노출기준", "끓는점 (℃)", "작업환경측정 대상", "특수건강진단 대상", "관리대상유해물질",
        "특별관리물질", "H-Code", "성분 위험등급", "KOSHA 조회상태", "산업안전보건법 규제사항",
    ]
    ws.append(headers)
    keys = [
        "process", "product", "original_name", "kosha_name", "cas", "content", "cmr",
        "exposure_limit", "boiling_point", "work_measurement", "special_health_exam", "managed_substance",
        "specially_managed", "h_codes", "grade", "kosha_status", "legal_basis",
    ]
    for r in rows:
        ws.append([r.get(k, "") for k in keys])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="193B50")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        letter = col[0].column_letter
        ws.column_dimensions[letter].width = min(max(max(len(str(c.value or "")) for c in col) + 2, 12), 38)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=msds_integrated.xlsx"},
    )


@app.post("/api/export/pdf")
async def export_pdf(payload: dict):
    rows = normalize_rows(payload.get("rows", []))
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=landscape(A4), rightMargin=18, leftMargin=18, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    font_name = "Helvetica"
    candidates = ["/usr/share/fonts/truetype/nanum/NanumGothic.ttf", "/usr/share/fonts/truetype/nanum/NanumGothicCoding.ttf"]
    for path in candidates:
        if os.path.exists(path):
            pdfmetrics.registerFont(TTFont("Korean", path))
            font_name = "Korean"
            break
    style = ParagraphStyle("small", parent=styles["BodyText"], fontName=font_name, fontSize=6, leading=7)
    headers = [
        "공정명", "제품명", "원본 물질명", "KOSHA 물질명", "CAS No.", "함유량", "CMR", "노출기준",
        "끓는점\n(℃)", "작업환경\n측정", "특수건강\n진단", "관리대상", "특별관리", "H-Code", "성분\n위험등급",
        "조회상태", "법령근거",
    ]
    data = [[Paragraph(x.replace("\n", "<br/>"), style) for x in headers]]
    keys = [
        "process", "product", "original_name", "kosha_name", "cas", "content", "cmr",
        "exposure_limit", "boiling_point", "work_measurement", "special_health_exam",
        "managed_substance", "specially_managed", "h_codes", "grade", "kosha_status", "legal_basis",
    ]
    for r in rows:
        data.append([
            Paragraph(str(r.get(k, "") or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"), style)
            for k in keys
        ])
    widths = [54, 64, 96, 96, 54, 44, 54, 66, 52, 42, 42, 42, 42, 50, 38, 58, 72]
    t = Table(data, colWidths=widths, repeatRows=1)
    t.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#193B50")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ])
    )
    doc.build([Paragraph("MSDS 유해화학물질 통합관리", styles["Title"]), Spacer(1, 8), t])
    bio.seek(0)
    return StreamingResponse(
        bio,
        media_type="application/pdf",
        headers={"Content-Disposition": "attachment; filename=msds_integrated.pdf"},
    )
