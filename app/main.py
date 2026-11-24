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
    grade_
