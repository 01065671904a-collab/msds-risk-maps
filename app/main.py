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
            if "수유독성" in window or "수유에 독성" in window or ("수유" in window and "독성" in window):
                return "수유독성"
            cat_m = re.search(r"(?:구분\s*)?(1A|1B|2|1)\b", window, re.I)
            if cat_m:
                return cat_m.group(1).upper()
    return ""


def extract_cmr_columns(hazard_obj, reg_obj, hcodes):
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


def _is_filled_flag(val):
    s = re.sub(r"\s+", "", str(val or ""))
    if not s or s in {"-", "N", "NO", "해당없음", "미해당", "X", "×"}:
        return False
    if any(tok in s for tok in ["✔", "✓", "V", "Y", "O", "개월", "년", "주", "일", "대상"]):
        return True
    return len(s) > 0


def parse_kosha_web_regulation(cas):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer": f"{KOSHA_WEB_BASE}/",
    }
    try:
        search_url = f"{KOSHA_WEB_BASE}/search/totalSearch.do?searchWrd={quote(cas)}"
        resp = requests.get(search_url, headers=headers, timeout=12)
        if resp.status_code != 200 or not resp.text:
            return None
        soup = BeautifulSoup(resp.text, "html.parser")
        target_href = None
        for tr in soup.select("table tr"):
            row_txt = re.sub(r"\s+", " ", tr.get_text(" ", strip=True))
            if cas in row_txt and ("법규" in row_txt or "산업안전보건법" in row_txt or tr.find("a")):
                a = tr.find("a", href=True)
                if a:
                    href = a["href"]
                    if href.startswith("http"):
                        target_href = href
                    elif href.startswith("/"):
                        target_href = "https://msds.kosha.or.kr" + href
                    else:
                        target_href = f"{KOSHA_WEB_BASE}/{href.lstrip('./')}"
                    break
        if not target_href:
            return None
        detail_resp = requests.get(target_href, headers=headers, timeout=12)
        if detail_resp.status_code != 200 or not detail_resp.text:
            return None
        dsoup = BeautifulSoup(detail_resp.text, "html.parser")
        if cas not in re.sub(r"\s+", " ", dsoup.get_text(" ", strip=True)):
            return None

        flags = {
            "work_measurement": "N",
            "special_health_exam": "N",
            "managed_substance": "N",
            "specially_managed": "N",
        }
        for tbl in dsoup.select("table"):
            trs = tbl.select("tr")
            for i, tr in enumerate(trs):
                ths = [re.sub(r"\s+", " ", th.get_text(" ", strip=True)) for th in tr.select("th")]
                if not ths or i + 1 >= len(trs):
                    continue
                tds = [re.sub(r"\s+", " ", td.get_text(" ", strip=True)) for td in trs[i + 1].select("td")]
                if not tds:
                    continue
                for idx, h_lbl in enumerate(ths):
                    cell_val = tds[idx] if idx < len(tds) else ""
                    filled = _is_filled_flag(cell_val)
                    if "작업환경측정" in h_lbl:
                        flags["work_measurement"] = "Y" if filled else "N"
                    elif "특수건강진단" in h_lbl:
                        flags["special_health_exam"] = "Y" if filled else "N"
                    elif "관리대상유해물질" in h_lbl:
                        flags["managed_substance"] = "Y" if filled else "N"
                    elif "특별관리물질" in h_lbl:
                        flags["specially_managed"] = "Y" if filled else "N"
        return flags
    except Exception:
        return None


def infer_law_flags(reg_text, full_reg_obj=None, cas=None):
    web_flags = parse_kosha_web_regulation(cas) if cas else None
    if web_flags:
        return web_flags

    combined = " ".join([reg_text or "", " ".join(clean_values(flatten_text(full_reg_obj or {})))])
    compact = re.sub(r"\s+", "", combined)
    if not compact or compact == "미확인":
        return {
            "work_measurement": "미확인",
            "special_health_exam": "미확인",
            "managed_substance": "미확인",
            "specially_managed": "미확인",
        }

    def check_item(label_names):
        for lbl in label_names:
            for m in re.finditer(re.escape(lbl), compact):
                tail = compact[m.end() : m.end() + 12]
                if any(tok in tail for tok in ["Y", "✔", "✓", "V", "O", "개월", "년", "해당"]):
                    return "Y"
        return "N"

    return {
        "work_measurement": check_item(["작업환경측정대상유해인자", "작업환경측정"]),
        "special_health_exam": check_item(["특수건강진단대상유해인자", "특수건강진단"]),
        "managed_substance": check_item(["관리대상유해물질"]),
        "specially_managed": check_item(["특별관리물질"]),
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
        "cmr_c": "미확인",
        "cmr_m": "미확인",
        "cmr_r": "미확인",
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
    doc = fitz.open(stream=data, filetype="pdf")
    chunks = []
    structured_items = []
    seen_cas = set()

    for page in doc:
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
                        name_val = ""
                        for idx in range(cas_col_idx):
                            cand = cells[idx]
                            if cand and not any(h in cand for h in ["화학물질명", "물질명", "관용명", "식별번호"]):
                                name_val = cand
                                break
                        pct_val = ""
                        for idx in range(cas_col_idx + 1, len(cells)):
                            pct_val = _parse_pct_from_text(cells[idx])
                            if pct_val:
                                break
                        seen_cas.add(cas_val)
                        structured_items.append(_make_ingredient_item(name_val, cas_val, pct_val))
        except Exception:
            pass

        word_lines = []
        try:
            words = page.get_text("words") or []
            if words:
                words.sort(key=lambda w: (w[1], w[0]))
                rows_y = []
                for w in words:
                    if not rows_y or abs(w[1] - rows_y[-1]["y"]) > 8.0:
                        rows_y.append({"y": w[1], "words": [w]})
                    else:
                        rows_y[-1]["words"].append(w)

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
        return HTMLResponse(content=f.read())


@app.get("/api/health")
def health():
    return JSONResponse(
        content={
            "ok": True,
            "kosha_key_configured": bool(KEY),
            "anonymous_access": True,
            "daily_lookup_limit": MAX_NEW_CAS_LOOKUPS_PER_DAY,
        }
    )


@app.post("/api/extract")
async def extract(file: UploadFile = File(...), request: Request = None):
    try:
        client_ip = getattr(getattr(request, "client", None), "host", "unknown") if request else "unknown"
        now = time.time()
        uploads = [t for t in _IP_CALLS["upload:" + client_ip] if now - t < 60]
        if len(uploads) >= 10:
            return json_err(429, "PDF 분석 요청이 많습니다. 1분 후 다시 시도해 주세요.")
        uploads.append(now)
        _IP_CALLS["upload:" + client_ip] = uploads
        if not file or not (file.filename or "").lower().endswith(".pdf"):
            return json_err(400, "PDF 파일만 업로드할 수 있습니다.")
        data = await file.read()
        if not data:
            return json_err(400, "빈 파일입니다. 올바른 PDF를 선택해 주세요.")
        if len(data) > 20 * 1024 * 1024:
            return json_err(413, "파일은 20MB 이하로 업로드해 주세요.")

        ingredients, txt = extract_pdf_ingredients_and_text(data)
        return JSONResponse(
            content={
                "filename": file.filename,
                "ingredients": ingredients,
                "text_preview": txt[:1200],
                "ocr_used": len(txt.strip()) > 0,
                "warning": "자동 추출값은 원본 MSDS 3번 항목과 대조해야 합니다.",
            }
        )
    except Exception as e:
        return json_err(400, f"PDF 분석 실패: {e}")


@app.get("/api/kosha/{cas}")
def kosha_lookup(cas: str, request: Request):
    if not CAS_RE.fullmatch(cas):
        return json_err(400, "CAS No. 형식이 올바르지 않습니다.")
    try:
        client_ip = getattr(getattr(request, "client", None), "host", "unknown")
        cached = protect_public_kosha_request(client_ip, cas)
        if cached is not None:
            return JSONResponse(content=cached)

        listing = api_get("getChemList001", {"searchWrd": cas, "searchCnd": "1", "numOfRows": "20", "pageNo": "1"})
        chem_id = first_value(listing, ["chemId", "chemID", "chemicalId", "화학물질ID"])
        if not chem_id:
            web_flags = parse_kosha_web_regulation(cas)
            raw = " ".join(clean_values(flatten_text(listing)))
            result = {
                "cas": cas,
                "found": False,
                "kosha_name": "미확인",
                "h_codes": "",
                "grade": "미확인",
                "cmr_c": "미확인",
                "cmr_m": "미확인",
                "cmr_r": "미확인",
                "exposure_limit": "미확인",
                "boiling_point": "미확인",
                "work_measurement": web_flags["work_measurement"] if web_flags else "미확인",
                "special_health_exam": web_flags["special_health_exam"] if web_flags else "미확인",
                "managed_substance": web_flags["managed_substance"] if web_flags else "미확인",
                "specially_managed": web_flags["specially_managed"] if web_flags else "미확인",
                "regulation_text": "미확인",
                "status": "KOSHA 목록에서 검색되지 않음",
                "raw": raw[:2500],
            }
            with _CACHE_LOCK:
                _CACHE[cas] = result
            return JSONResponse(content=result)

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
        cmr_cols = extract_cmr_columns(details.get("hazard", {}), details.get("regulation", {}), hcodes)

        reg_text = industrial_safety_law_text(details.get("regulation", {}))[:4000] or "미확인"
        flags = infer_law_flags(reg_text, details.get("regulation", {}), cas=cas)

        result = {
            "cas": cas,
            "found": True,
            "chem_id": chem_id,
            "kosha_name": chem_name or "미확인",
            "h_codes": ", ".join(hcodes),
            "grade": grade_str,
            "cmr_c": cmr_cols["cmr_c"],
            "cmr_m": cmr_cols["cmr_m"],
            "cmr_r": cmr_cols["cmr_r"],
            "exposure_limit": exposure,
            "boiling_point": boiling,
            "work_measurement": flags["work_measurement"],
            "special_health_exam": flags["special_health_exam"],
            "managed_substance": flags["managed_substance"],
            "specially_managed": flags["specially_managed"],
            "regulation_text": reg_text,
            "status": "KOSHA 조회 완료",
        }
        with _CACHE_LOCK:
            _CACHE[cas] = result
        return JSONResponse(content=result)
    except Exception as e:
        return json_err(500, f"KOSHA 조회 실패: {e}")


def normalize_rows(rows):
    if not isinstance(rows, list) or len(rows) > 20000:
        return []
    return rows


@app.post("/api/export/xlsx")
async def export_xlsx(payload: dict):
    rows = normalize_rows((payload or {}).get("rows", []))
    wb = Workbook()
    ws = wb.active
    ws.title = "MSDS 통합관리"
    headers = [
        "공정명", "제품명", "원본 MSDS 물질명", "KOSHA 조회 물질명", "CAS No.", "함유량",
        "CMR-C", "CMR-M", "CMR-R", "노출기준", "끓는점 (℃)",
        "작업환경측정", "특수건강진단", "관리대상유해물질", "특별관리물질",
        "H-Code", "성분 위험등급", "KOSHA 조회상태", "산업안전보건법 규제사항",
    ]
    ws.append(headers)
    keys = [
        "process", "product", "original_name", "kosha_name", "cas", "content",
        "cmr_c", "cmr_m", "cmr_r", "exposure_limit", "boiling_point",
        "work_measurement", "special_health_exam", "managed_substance", "specially_managed",
        "h_codes", "grade", "kosha_status", "legal_basis",
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
    rows = normalize_rows((payload or {}).get("rows", []))
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
        "공정명", "제품명", "원본 물질명", "KOSHA 물질명", "CAS No.", "함유량",
        "CMR\n(C)", "CMR\n(M)", "CMR\n(R)", "노출기준", "끓는점\n(℃)",
        "작업환경\n측정", "특수건강\n진단", "관리대상", "특별관리",
        "H-Code", "성분\n위험등급", "조회상태", "법령근거",
    ]
    data = [[Paragraph(x.replace("\n", "<br/>"), style) for x in headers]]
    keys = [
        "process", "product", "original_name", "kosha_name", "cas", "content",
        "cmr_c", "cmr_m", "cmr_r", "exposure_limit", "boiling_point",
        "work_measurement", "special_health_exam", "managed_substance", "specially_managed",
        "h_codes", "grade", "kosha_status", "legal_basis",
    ]
    for r in rows:
        data.append([
            Paragraph(str(r.get(k, "") or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"), style)
            for k in keys
        ])
    widths = [48, 56, 84, 84, 50, 38, 36, 36, 42, 60, 44, 38, 38, 38, 38, 46, 34, 50, 66]
    t = Table(data, colWidths=widths, repeatRows=1)
    t.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#193B50")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ])
    )
    doc.build([Paragraph("MSDS 유해화학물질 통합관리", styles["Title"]), Spacer(1, 8), t])
    bio.seek(0)
    return StreamingResponse(
        bio,
        media_type="application/pdf",
        headers={"Content-Disposition": "attachment; filename=msds_integrated.pdf"},
    )
