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
        if cas in _