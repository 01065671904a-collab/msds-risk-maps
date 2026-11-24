import os
import re
import io
import time
import threading
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import date
from urllib.parse import quote

import requests
from bs4 import BeautifulSoup
import fitz
import pytesseract
from PIL import Image
from fastapi import FastAPI, UploadFile, File, Request
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
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
KOSHA_WEB_BASE = "https://msds.kosha.or.kr/MSDSInfo"
KEY = os.getenv("KOSHA_SERVICE_KEY", "").strip()
_CACHE = {}
_CACHE_LOCK = threading.Lock()
_IP_CALLS = defaultdict(list)
_DAILY_LOOKUPS = {"day": date.today().isoformat(), "count": 0}
MAX_NEW_CAS_LOOKUPS_PER_DAY = int(os.getenv("MAX_NEW_CAS_LOOKUPS_PER_DAY", "250"))

CAS_RE = re.compile(r"(?<!\d)(\d{2,7}-\d{2}-\d)(?!\d)")
FLEX_CAS_RE = re.compile(r"(?<!\d)(\d{2,7})\s*[-–—‐‑‒−]\s*(\d{2})\s*[-–—‐‑‒−]\s*(\d)(?!\d)")
H_RE = re.compile(
    r"\bH(?:300|301|302|304|310|311|312|314|315|317|318|319|320|330|331|332|334|335|336|340|341|350|351|360|361|362|370|371|372|373)\b"
)

# 4등급 > 3등급 > 2등급 > 1등급
H_LEVEL = {
    "H340": 4, "H341": 4, "H350": 4, "H351": 4, "H360": 4, "H361": 4, "H362": 4, "H370": 4,
    "H300": 3, "H301": 3, "H310": 3, "H311": 3, "H330": 3, "H331": 3,
    "H314": 3, "H318": 3, "H317": 3, "H334": 3, "H372": 3, "H373": 3,
    "H302": 2, "H312": 2, "H332": 2, "H304": 2, "H371": 2,
    "H315": 1, "H319": 1, "H320": 1, "H335": 1, "H336": 1,
}

NOISE_TOKENS = {"00", "NORMAL SERVICE.", "NORMAL SERVICE", "NULL", "NONE", "-", "N/A"}


def json_err(status_code: int, message: str):
    return JSONResponse(status_code=status_code, content={"detail": message, "error": message})


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
        raise RuntimeError("KOSHA_SERVICE_KEY가 설정되지 않았습니다. Render 환경변수에 공공데이터포털 인증키를 등록하세요.")
    q = {"serviceKey": KEY, **params}
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
            raise RuntimeError(f"KOSHA API 오류({code}): {message}")
    return parse_kosha_xml(r.content)


def protect_public_kosha_request(client_ip, cas):
    now = time.time()
    recent = [t for t in _IP_CALLS[client_ip] if now - t < 60]
    if len(recent) >= 30:
        raise RuntimeError("요청이 너무 많습니다. 1분 후 다시 시도해 주세요.")
    recent.append(now)
    _IP_CALLS[client_ip] = recent
    today = date.today().isoformat()
    if _DAILY_LOOKUPS["day"] != today:
        _DAILY_LOOKUPS.update({"day": today, "count": 0})
    with _CACHE_LOCK:
        if cas in _CACHE:
            return _CACHE[cas]
        if _DAILY_LOOKUPS["count"] >= MAX_NEW_CAS_LOOKUPS_PER_DAY:
            raise RuntimeError("오늘의 공개 KOSHA 조회 한도에 도달했습니다. 내일 다시 이용해 주세요.")
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
    if not combined or (
        "자료없음" in combined
        and not re.search(r"(TWA|STEL|CEILING|(?<![A-Z])C(?![A-Z]))", combined, re.I)
    ):
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


def _is_flash_or_ignition_only(text):
    low = (text or "").lower()
    if any(k in low for k in ["인화점", "발화점", "flash", "ignition", "autoignition"]):
        if not any(b in low for b in ["끓는점", "boiling"]):
            return True
    return False


def extract_boiling_text(detail_obj):
    """끓는점 정보만 추출하며 인화점/발화점과 혼동하지 않습니다."""
    bp_texts = []

    def collect_bp(node):
        if isinstance(node, dict):
            for k, v in node.items():
                ks = str(k).lower()
                if any(bad in ks for bad in ["flash", "ignition", "인화", "발화"]):
                    continue
                if any(tag in ks for tag in ["끓는점", "boiling", "boil"]):
                    bp_texts.extend(clean_values(flatten_text(v), allow_no_data=True))
                else:
                    collect_bp(v)
        elif isinstance(node, list):
            for v in node:
                collect_bp(v)

    collect_bp(detail_obj)

    if not bp_texts:
        def collect_by_label(node):
            if isinstance(node, dict):
                flat = clean_values(flatten_text(node), allow_no_data=True)
                for idx, txt in enumerate(flat):
                    if _is_flash_or_ignition_only(txt):
                        continue
                    if any(lbl in txt for lbl in ["초기 끓는점", "끓는점 범위", "초기끓는점", "끓는점"]):
                        segment = " ".join(flat[idx : idx + 3])
                        if not _is_flash_or_ignition_only(segment):
                            bp_texts.append(segment)
                for v in node.values():
                    collect_by_label(v)
            elif isinstance(node, list):
                for v in node:
                    collect_by_label(v)

        collect_by_label(detail_obj)

    joined = " ".join(bp_texts)
    if not joined or _is_flash_or_ignition_only(joined):
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


def _extract_category_near(text, keywords):
    for kw in keywords:
        for m in re.finditer(re.escape(kw), text):
            window = text[max(0, m.start() - 30) : min(len(text), m.end() + 60)]
            if "수유독성" in window or "수유에 독성" in window or "수유" in window and "독성" in window:
                return "수유독성"
            cat_m = re.search(r"(?:구분\s*)?(1A|1B|2|1)\b", window, re.I)
            if cat_m:
                return cat_m.group(1).upper()
    return ""


def extract_cmr_columns(hazard_obj, reg_obj, hcodes):
    """C(발암성), M(생식세포 변이원성), R(생식독성) 항목별 구분을 분리 반환합니다."""
    combined = " ".join(clean_values(flatten_text(hazard_obj)) + clean_values(flatten_text(reg_obj)))
    has_any_data = bool(combined.strip()) or bool(hcodes)

    c_cat = _extract_category_near(combined, ["발암성", "carcinogenicity"])
    m_cat = _extract_category_near(combined, ["생식세포 변이원성", "생식세포변이원성", "변이원성", "mutagenicity"])
    r_cat = _extract_category_near(combined, ["생식독성", "reproductive toxicity"])

    if not c_cat and any(x in hcodes for x in ["H350", "H351"]):
        c_cat = "구분확인필요"
    if not m_cat and any(x in hcodes for x in ["H340", "H341"]):
        m_cat = "구분확인필요"
    if not r_cat and any(x in hcodes for x in ["H360", "H361", "H362"]):
        r_cat = "수유독성" if "H362" in hcodes else "구분확인필요"

    default_val = "해당없음" if has_any_data else "미확인"
    return {
        "cmr_c": c_cat or default_val,
        "cmr_m": m_cat or default_val,
        "cmr_r": r_cat or default_val,
    }


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
@app.get("/api/health")
def health():
    return {"ok": True}
