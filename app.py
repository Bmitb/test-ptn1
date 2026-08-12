"""
Pressure Calibration Report Parser
===================================
A Streamlit application that uses Google Gemini AI to extract structured data
from handwritten pressure calibration reports (PDF/Images) and appends the
parsed data into an existing Excel template with two sheets.

Author  : Senior Python Full-Stack Developer
Version : 1.0.0
"""

from __future__ import annotations

import base64
import io
import json
import os
import re
import tempfile
import traceback
from copy import copy
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG  (must be the very first Streamlit call)
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Pressure Calibration Parser",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(
    """
<style>
/* ── Google Font ── */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

/* ════════════════════════════════════════════════════════
   GLOBAL & BACKGROUND
   ════════════════════════════════════════════════════════ */
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif !important;
    color: #dde3f5 !important;
}
.stApp {
    background: linear-gradient(135deg, #0f0f1a 0%, #141428 40%, #0d1b2a 100%) !important;
}

/* ════════════════════════════════════════════════════════
   SIDEBAR
   ════════════════════════════════════════════════════════ */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0d1117 0%, #161b27 100%) !important;
    border-right: 1px solid rgba(255,255,255,0.07) !important;
}
/* All sidebar text */
[data-testid="stSidebar"],
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] div,
[data-testid="stSidebar"] label {
    color: #c8d4ee !important;
}

/* ════════════════════════════════════════════════════════
   LABELS & WIDGET HEADERS
   ════════════════════════════════════════════════════════ */
[data-testid="stWidgetLabel"] p,
[data-testid="stWidgetLabel"] span,
[data-testid="stWidgetLabel"],
.stTextInput label,
.stSelectbox label,
.stFileUploader label,
.stNumberInput label,
.stTextArea label,
.stMultiSelect label,
label {
    color: #c8d4ee !important;
    font-weight: 500 !important;
    font-size: 0.88rem !important;
}

/* ════════════════════════════════════════════════════════
   TEXT INPUT  (the white box problem)
   ════════════════════════════════════════════════════════ */
/* Wrapper */
.stTextInput > div,
.stTextInput > div > div,
[data-testid="stTextInput"] > div,
[data-testid="stTextInput"] > div > div {
    background: rgba(20, 26, 48, 0.9) !important;
    border-color: rgba(99,179,237,0.25) !important;
    border-radius: 10px !important;
}
/* Actual <input> element */
.stTextInput input,
[data-testid="stTextInput"] input,
input[type="text"],
input[type="password"],
input[type="number"],
input[type="email"],
input[type="search"] {
    color: #eef1fc !important;
    background: rgba(20, 26, 48, 0.9) !important;
    border-color: rgba(99,179,237,0.25) !important;
    border-radius: 10px !important;
    caret-color: #63b3ed !important;
}
.stTextInput input::placeholder,
input::placeholder {
    color: rgba(180, 195, 230, 0.4) !important;
}
.stTextInput input:focus,
input:focus {
    border-color: rgba(99,179,237,0.7) !important;
    box-shadow: 0 0 0 2px rgba(99,179,237,0.15) !important;
}

/* ════════════════════════════════════════════════════════
   TEXTAREA
   ════════════════════════════════════════════════════════ */
textarea,
.stTextArea textarea {
    color: #eef1fc !important;
    background: rgba(20, 26, 48, 0.9) !important;
    border-color: rgba(99,179,237,0.25) !important;
    border-radius: 10px !important;
}

/* ════════════════════════════════════════════════════════
   SELECTBOX  (the most problematic one)
   ════════════════════════════════════════════════════════ */
/* Outer wrapper */
.stSelectbox > div > div,
[data-testid="stSelectbox"] > div,
[data-testid="stSelectbox"] > div > div,
[data-baseweb="select"],
[data-baseweb="select"] > div {
    background: rgba(20, 26, 48, 0.9) !important;
    border-color: rgba(99,179,237,0.25) !important;
    border-radius: 10px !important;
    color: #eef1fc !important;
}
/* Selected value text */
[data-baseweb="select"] span,
[data-baseweb="select"] div,
[data-baseweb="select"] input,
.stSelectbox [data-baseweb="select"] *,
[data-testid="stSelectbox"] span,
[data-testid="stSelectbox"] div {
    color: #eef1fc !important;
    background: transparent !important;
}
/* Dropdown menu */
[data-baseweb="menu"],
[data-baseweb="popover"],
ul[data-baseweb="menu"],
[role="listbox"],
[role="option"] {
    background: #1a2035 !important;
    border: 1px solid rgba(99,179,237,0.2) !important;
    border-radius: 10px !important;
}
[role="option"],
[data-baseweb="menu"] li,
[data-baseweb="menu"] * {
    color: #c8d4ee !important;
    background: transparent !important;
}
[role="option"]:hover,
[data-baseweb="menu"] li:hover {
    background: rgba(99,179,237,0.12) !important;
    color: #ffffff !important;
}

/* ════════════════════════════════════════════════════════
   CHECKBOX
   ════════════════════════════════════════════════════════ */
.stCheckbox label p,
.stCheckbox label span,
[data-testid="stCheckbox"] label,
[data-testid="stCheckbox"] span,
[data-testid="stCheckbox"] p {
    color: #c8d4ee !important;
    font-size: 0.87rem !important;
}
/* The checkbox box itself */
[data-baseweb="checkbox"] span {
    border-color: rgba(99,179,237,0.5) !important;
    background: rgba(20,26,48,0.9) !important;
}

/* ════════════════════════════════════════════════════════
   FILE UPLOADER — force dark background
   ════════════════════════════════════════════════════════ */
[data-testid="stFileUploader"],
[data-testid="stFileUploaderDropzone"] {
    background: transparent !important;
}
[data-testid="stFileUploader"] > div,
[data-testid="stFileUploader"] section,
[data-testid="stFileUploaderDropzone"],
[data-testid="stFileUploaderDropzone"] > div {
    background: rgba(15, 20, 40, 0.92) !important;
    border: 2px dashed rgba(99,179,237,0.45) !important;
    border-radius: 14px !important;
    transition: all .3s ease !important;
}
[data-testid="stFileUploaderDropzone"]:hover,
[data-testid="stFileUploader"] > div:hover {
    border-color: rgba(99,179,237,0.85) !important;
    box-shadow: 0 0 20px rgba(99,179,237,0.1) !important;
}
[data-testid="stFileUploader"] *,
[data-testid="stFileUploaderDropzone"] * {
    color: #c8d4ee !important;
    background: transparent !important;
}
/* Browse files button */
[data-testid="stFileUploaderDropzone"] button,
[data-testid="stFileUploader"] button {
    background: rgba(99,179,237,0.15) !important;
    color: #63b3ed !important;
    border: 1px solid rgba(99,179,237,0.4) !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
}
[data-testid="stFileUploaderDropzone"] button:hover {
    background: rgba(99,179,237,0.28) !important;
}

/* ════════════════════════════════════════════════════════
   INFO / WARNING / SUCCESS / ERROR ALERTS
   ════════════════════════════════════════════════════════ */
[data-testid="stAlert"],
[data-testid="stAlert"] p,
[data-testid="stAlert"] span,
[data-testid="stAlert"] div,
.stSuccess, .stWarning, .stError, .stInfo {
    color: #eef1fc !important;
}

/* ════════════════════════════════════════════════════════
   CAPTIONS / SECONDARY TEXT
   ════════════════════════════════════════════════════════ */
.stCaption,
[data-testid="stCaptionContainer"],
[data-testid="stCaptionContainer"] p,
[data-testid="stCaptionContainer"] span,
small {
    color: #8a9bc4 !important;
}

/* ════════════════════════════════════════════════════════
   METRIC WIDGET
   ════════════════════════════════════════════════════════ */
[data-testid="stMetricLabel"] p,
[data-testid="stMetricLabel"] {
    color: #8a9bc4 !important;
}
[data-testid="stMetricValue"],
[data-testid="stMetricValue"] *{
    color: #eef1fc !important;
}

/* ════════════════════════════════════════════════════════
   EXPANDER
   ════════════════════════════════════════════════════════ */
details summary p,
details summary span,
.streamlit-expanderHeader p,
.streamlit-expanderHeader {
    color: #a78bfa !important;
    font-weight: 600 !important;
}
[data-testid="stExpander"] summary p {
    color: #a78bfa !important;
}

/* ════════════════════════════════════════════════════════
   MARKDOWN HEADINGS
   ════════════════════════════════════════════════════════ */
.stMarkdown h1, .stMarkdown h2,
.stMarkdown h3, .stMarkdown h4,
h1, h2, h3, h4 {
    color: #eef1fc !important;
}
.stMarkdown p, .stMarkdown li {
    color: #c8d4ee !important;
}
.stMarkdown a { color: #63b3ed !important; }

/* ════════════════════════════════════════════════════════
   DIVIDER
   ════════════════════════════════════════════════════════ */
hr { border-color: rgba(255,255,255,0.08) !important; }

/* ════════════════════════════════════════════════════════
   CUSTOM COMPONENTS
   ════════════════════════════════════════════════════════ */

/* Glass card */
.glass-card {
    background: rgba(255,255,255,0.04);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 16px;
    padding: 24px;
    margin-bottom: 20px;
    backdrop-filter: blur(12px);
    transition: border-color .3s ease;
}
.glass-card:hover { border-color: rgba(99,179,237,0.25); }

/* Section header */
.section-header {
    display: flex;
    align-items: center;
    gap: 10px;
    font-size: 1.05rem;
    font-weight: 600;
    color: #63b3ed;
    margin-bottom: 16px;
    padding-bottom: 10px;
    border-bottom: 1px solid rgba(99,179,237,0.2);
}

/* Status badges */
.badge-ok {
    display: inline-block;
    background: linear-gradient(135deg, #22543d, #276749);
    color: #68d391;
    border: 1px solid rgba(104,211,145,.4);
    padding: 3px 12px;
    border-radius: 20px;
    font-size: .78rem;
    font-weight: 600;
}
.badge-error {
    display: inline-block;
    background: linear-gradient(135deg, #742a2a, #9b2c2c);
    color: #fc8181;
    border: 1px solid rgba(252,129,129,.4);
    padding: 3px 12px;
    border-radius: 20px;
    font-size: .78rem;
    font-weight: 600;
}

/* Hero banner */
.hero-banner {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
    border: 1px solid rgba(99,179,237,0.2);
    border-radius: 20px;
    padding: 32px;
    text-align: center;
    margin-bottom: 28px;
    position: relative;
    overflow: hidden;
}
.hero-banner::before {
    content: '';
    position: absolute;
    top: -50%; left: -50%;
    width: 200%; height: 200%;
    background: radial-gradient(ellipse at center, rgba(99,179,237,0.06) 0%, transparent 70%);
    pointer-events: none;
}
.hero-title {
    font-size: 2rem;
    font-weight: 700;
    background: linear-gradient(135deg, #63b3ed, #a78bfa, #f687b3);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    margin-bottom: 8px;
}
.hero-sub {
    color: rgba(255,255,255,0.75) !important;
    font-size: .95rem;
    font-weight: 400;
}

/* Metric pill */
.metric-pill {
    background: rgba(255,255,255,0.05);
    border: 1px solid rgba(255,255,255,0.1);
    border-radius: 12px;
    padding: 14px 18px;
    text-align: center;
}
.metric-value { font-size: 1.4rem; font-weight: 700; color: #a78bfa !important; }
.metric-label { font-size: .75rem; color: #8a9bc4 !important; margin-top: 2px; }

/* Buttons */
.stButton > button {
    border-radius: 10px !important;
    font-weight: 600 !important;
    transition: all .25s ease !important;
}
.stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 24px rgba(0,0,0,0.4);
}

/* Data editor */
[data-testid="stDataFrame"],
[data-testid="data-grid-canvas"] {
    border-radius: 10px;
    overflow: hidden;
}

/* Scrollbar */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: rgba(255,255,255,.05); }
::-webkit-scrollbar-thumb { background: rgba(99,179,237,.4); border-radius: 3px; }
</style>
""",
    unsafe_allow_html=True,
)


# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
GEMINI_MODELS = [
    "gemini-2.5-flash",          # ⭐ Nhanh, rẻ — khuyến nghị
    "gemini-2.5-pro",            # Chính xác nhất
    "gemini-2.5-flash-lite",     # Nhẹ nhất
    "gemini-3-flash-preview",    # Thế hệ mới
    "gemini-3.5-flash",          # Mới nhất ổn định
    "gemini-flash-latest",       # Alias mới nhất
    "gemini-pro-latest",         # Pro mới nhất
]



def fetch_available_models(api_key: str) -> list[str]:
    """
    Query the Gemini API to list all models that support generateContent.
    Uses the new google-genai SDK.
    Returns a sorted list of model IDs (without the 'models/' prefix).
    """
    try:
        from google import genai
        client = genai.Client(api_key=api_key)
        models = [
            m.name.replace("models/", "")
            for m in client.models.list()
            if m.supported_actions and "generateContent" in m.supported_actions
        ]
        # fallback: if supported_actions not present, include all
        if not models:
            models = [
                m.name.replace("models/", "")
                for m in client.models.list()
            ]
        return sorted(models)
    except Exception as e:
        raise ValueError(f"Không thể lấy danh sách model: {e}")


SHEET1_HEADERS = [
    "GCN Số", "Mã ID", "Mã số nhận dạng", "Tên UUT",
    "Khách hàng", "Phiếu YCCV", "Người thực hiện", "P.pháp HC",
    "Ngày hiệu chuẩn", "Kết quả HC", "Tem hiệu chuẩn",
    "Ngày HC kế tiếp", "TB Chuẩn 1",
]

SHEET2_HEADERS = [
    "Mã Phụ", "GCN Số", "Mã QL / Mã ID", "Đ.vị",
    "Min", "Max", "Điểm HC", "Đơn vị P", "P",
    "Đơn vị Chuẩn P", "P c.tăng", "P c.giảm",
]

EXTRACTION_PROMPT = """You are an expert OCR assistant specialising in Vietnamese pressure calibration certificates.
Analyse the provided document image carefully and extract ALL handwritten and printed information.

Return ONLY a single valid JSON object — no markdown fences, no commentary.

JSON schema (strictly follow this):
{
  "gcn_so": "<Báo cáo số / GCN Số / Report No>",
  "ma_id": "<Mã/ID of the instrument being calibrated>",
  "ten_uut": "<Loại mẫu / instrument type, e.g. PG, PGPI, Pressure Gauge>",
  "khach_hang": "<Khách hàng / Customer name>",
  "nguoi_thuc_hien": "<Full name on Người thực hiện / Technician signature line>",
  "ngay_hc": "<Ngày HC / Calibration date in DD/MM/YYYY format>",
  "ket_qua": "<'OK' if Đạt is checked, else 'FAIL'>",
  "tem_hc": "<Tem hiệu chuẩn / Calibration label number>",
  "ngay_ke_tiep": "<Ngày tới hạn / Due date in DD/MM/YYYY format>",
  "tb_chuan_1": "<Mã số TB from the Chuẩn được sử dụng / Reference standard table>",
  "don_vi": "<unit of pressure, e.g. bar, MPa, kPa, psi>",
  "range_min": <numeric minimum of calibration range, e.g. 0>,
  "range_max": <numeric maximum of calibration range, e.g. 700>,
  "points": [
    {
      "point_id": "D1",
      "p_value": <numeric value of UUT/REF set point>,
      "p_tang":  <numeric reading during increasing stroke>,
      "p_giam":  <numeric reading during decreasing stroke>
    }
  ]
}

Rules:
- Extract ALL calibration data points (D1 … Dn).
- Use null for any field that cannot be determined.
- All numeric fields must be numbers, not strings.
- Return ONLY the JSON object, nothing else.
"""

# ─────────────────────────────────────────────────────────────────────────────
# HELPER UTILITIES
# ─────────────────────────────────────────────────────────────────────────────

def _safe_float(val: Any) -> float | None:
    """Convert any value to float, return None on failure."""
    if val is None:
        return None
    try:
        return float(str(val).replace(",", ".").strip())
    except (ValueError, TypeError):
        return None


def _parse_json_response(text: str) -> dict:
    """
    Extract and parse the JSON payload from Gemini's response text,
    even if it accidentally wraps it in markdown code fences.
    """
    text = text.strip()
    # Strip markdown fences
    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.MULTILINE)
    text = re.sub(r"\s*```$", "", text, flags=re.MULTILINE)
    text = text.strip()

    # Try direct parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Attempt to grab the outermost {...} block
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        return json.loads(match.group(0))

    raise ValueError(f"Could not parse JSON from response:\n{text[:500]}")


def _copy_row_style(ws, src_row: int, dest_row: int):
    """Copy cell styles from src_row to dest_row (best-effort)."""
    for col in range(1, ws.max_column + 1):
        src_cell = ws.cell(row=src_row, column=col)
        dest_cell = ws.cell(row=dest_row, column=col)
        if src_cell.has_style:
            dest_cell.font = copy(src_cell.font)
            dest_cell.fill = copy(src_cell.fill)
            dest_cell.border = copy(src_cell.border)
            dest_cell.alignment = copy(src_cell.alignment)
            dest_cell.number_format = src_cell.number_format


# ─────────────────────────────────────────────────────────────────────────────
# CORE FUNCTION 1: DATA EXTRACTION
# ─────────────────────────────────────────────────────────────────────────────

def extract_data_from_document(
    file_bytes: bytes,
    mime_type: str,
    api_key: str,
    model_name: str = "gemini-2.5-flash",
) -> dict:
    """
    Send a calibration document (PDF or image) to Google Gemini and return
    the extracted data as a Python dict.
    Uses the new `google-genai` SDK (replaces deprecated google-generativeai).
    """
    from google import genai
    from google.genai import types as genai_types

    client = genai.Client(api_key=api_key)

    # ── Render PDF → list of PNG bytes ────────────────────────────────────
    image_parts: list[genai_types.Part] = []

    if mime_type == "application/pdf":
        try:
            import fitz  # PyMuPDF
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            for page_num in range(min(len(doc), 5)):
                page = doc.load_page(page_num)
                pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
                img_bytes = pix.tobytes("png")
                image_parts.append(
                    genai_types.Part.from_bytes(data=img_bytes, mime_type="image/png")
                )
        except ImportError:
            try:
                from pdf2image import convert_from_bytes
                pages = convert_from_bytes(file_bytes, dpi=200)
                for page in pages[:5]:
                    buf = io.BytesIO()
                    page.save(buf, format="PNG")
                    image_parts.append(
                        genai_types.Part.from_bytes(data=buf.getvalue(), mime_type="image/png")
                    )
            except Exception as e:
                raise ValueError(
                    f"Không thể render PDF. Cài PyMuPDF: `pip install pymupdf`. Lỗi: {e}"
                )
        if not image_parts:
            raise ValueError("Không trích xuất được trang nào từ PDF.")
    else:
        # Image file (png/jpg/jpeg)
        image_parts.append(
            genai_types.Part.from_bytes(data=file_bytes, mime_type=mime_type)
        )

    # ── Call Gemini API with automatic retry & fallback ───────────────────
    import time
    contents = [EXTRACTION_PROMPT] + image_parts
    config = genai_types.GenerateContentConfig(
        temperature=0.0,
        max_output_tokens=4096,
    )

    # Candidate models to attempt if the primary one is 503 overloaded
    fallback_models = [model_name] + [
        m for m in ["gemini-2.5-flash", "gemini-1.5-flash", "gemini-2.5-pro", "gemini-flash-latest"]
        if m != model_name
    ]

    last_err = None
    for attempt_model in fallback_models:
        for attempt in range(3):  # 3 retries per model
            try:
                response = client.models.generate_content(
                    model=attempt_model,
                    contents=contents,
                    config=config,
                )
                result_text = response.text
                return _parse_json_response(result_text)
            except Exception as e:
                err_str = str(e)
                last_err = e
                # Check for 503 (server overloaded) or 429 (rate limit)
                if "503" in err_str or "UNAVAILABLE" in err_str or "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                    wait_time = (attempt + 1) * 2
                    time.sleep(wait_time)
                    continue
                else:
                    # Other non-retryable errors -> break retry loop and try next fallback model
                    break

    raise ValueError(
        f"Gemini API error (đã tự động thử lại nhưng máy chủ Google đang quá tải 503): {last_err}\n"
        f"💡 Gợi ý: Chọn model khác ở Sidebar (ví dụ: gemini-1.5-flash hoặc gemini-2.5-pro) hoặc thử lại sau 1-2 phút."
    )





# ─────────────────────────────────────────────────────────────────────────────
# CORE FUNCTION 2: EXCEL APPEND
# ─────────────────────────────────────────────────────────────────────────────

def append_to_excel(excel_path: str | Path, extracted_data: dict) -> bytes:
    """
    Append the extracted calibration data to the two sheets of the Excel
    template and return the modified workbook as bytes.

    Parameters
    ----------
    excel_path     : str or Path
        Path to the existing Excel template file.
    extracted_data : dict
        Validated extraction dict (matching the AI JSON schema).

    Returns
    -------
    bytes
        The in-memory workbook bytes ready for download.

    Raises
    ------
    ValueError
        When required sheets are not found.
    """
    wb = load_workbook(excel_path)

    # ── Resolve sheet references (by index or name) ────────────────────────
    if len(wb.sheetnames) < 1:
        raise ValueError("Excel template has no sheets.")

    ws1 = wb.worksheets[0]   # Sheet 1 — DANH MỤC HIỆU CHUẨN
    ws2 = wb.worksheets[1] if len(wb.worksheets) > 1 else wb.create_sheet("Sheet2")

    # ── Unpack data ─────────────────────────────────────────────────────────
    gcn_so          = extracted_data.get("gcn_so", "")
    ma_id           = extracted_data.get("ma_id", "")
    ten_uut         = extracted_data.get("ten_uut", "")
    khach_hang      = extracted_data.get("khach_hang", "")
    nguoi_thuc_hien = extracted_data.get("nguoi_thuc_hien", "")
    ngay_hc         = extracted_data.get("ngay_hc", "")
    ket_qua         = extracted_data.get("ket_qua", "OK")
    tem_hc          = extracted_data.get("tem_hc", "")
    ngay_ke_tiep    = extracted_data.get("ngay_ke_tiep", "")
    tb_chuan_1      = extracted_data.get("tb_chuan_1", "")
    don_vi          = extracted_data.get("don_vi", "")
    range_min       = extracted_data.get("range_min", None)
    range_max       = extracted_data.get("range_max", None)
    points          = extracted_data.get("points", [])

    # ────────────────────────────────────────────────────────────────────────
    # SHEET 1 — append a single row
    # ────────────────────────────────────────────────────────────────────────
    next_row_s1 = ws1.max_row + 1

    # Try to copy style from last data row (if it exists & is not header)
    if ws1.max_row >= 2:
        _copy_row_style(ws1, ws1.max_row, next_row_s1)

    row_s1 = [
        gcn_so,           # Col 1
        ma_id,            # Col 2
        ma_id,            # Col 3 — Mã số nhận dạng (same as Mã ID)
        ten_uut,          # Col 4
        khach_hang,       # Col 5
        "",               # Col 6 — Phiếu YCCV (empty)
        nguoi_thuc_hien,  # Col 7
        "DLVN76",         # Col 8 — P.pháp HC (hardcoded default)
        ngay_hc,          # Col 9
        ket_qua,          # Col 10
        tem_hc,           # Col 11
        ngay_ke_tiep,     # Col 12
        tb_chuan_1,       # Col 13
    ]

    for col_idx, value in enumerate(row_s1, start=1):
        ws1.cell(row=next_row_s1, column=col_idx, value=value)

    # ────────────────────────────────────────────────────────────────────────
    # SHEET 2 — append measurement point rows
    # ────────────────────────────────────────────────────────────────────────
    last_row_s2 = ws2.max_row

    # Add an empty separator row if the sheet already has data rows
    has_data_s2 = last_row_s2 >= 2
    if has_data_s2:
        last_row_s2 += 1  # blank separator row (leave empty)

    # Determine template style source row
    style_src = max(2, ws2.max_row - len(points)) if ws2.max_row > 1 else None

    for i, pt in enumerate(points):
        dest_row = last_row_s2 + 1 + i

        if style_src:
            _copy_row_style(ws2, style_src, dest_row)

        point_id = pt.get("point_id", f"D{i+1}")
        p_value  = _safe_float(pt.get("p_value"))
        p_tang   = _safe_float(pt.get("p_tang"))
        p_giam   = _safe_float(pt.get("p_giam"))

        # Col 5/6 — Min/Max: only on first row of each device block
        min_val = range_min if i == 0 else None
        max_val = range_max if i == 0 else None

        row_s2 = [
            f"{gcn_so}{point_id}",  # Col 1  — Mã Phụ
            gcn_so,                  # Col 2  — GCN Số
            ma_id,                   # Col 3  — Mã QL / Mã ID
            don_vi,                  # Col 4  — Đ.vị
            min_val,                 # Col 5  — Min
            max_val,                 # Col 6  — Max
            point_id,                # Col 7  — Điểm HC
            don_vi,                  # Col 8  — Đơn vị P
            p_value,                 # Col 9  — P
            don_vi,                  # Col 10 — Đơn vị Chuẩn P
            p_tang,                  # Col 11 — P c.tăng
            p_giam,                  # Col 12 — P c.giảm
        ]

        for col_idx, value in enumerate(row_s2, start=1):
            ws2.cell(row=dest_row, column=col_idx, value=value)

    # ── Save to in-memory buffer ─────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# HELPER: Build preview DataFrames from extracted data
# ─────────────────────────────────────────────────────────────────────────────

def _build_sheet1_df(data: dict) -> pd.DataFrame:
    row = {
        "GCN Số":            data.get("gcn_so", ""),
        "Mã ID":             data.get("ma_id", ""),
        "Mã số nhận dạng":   data.get("ma_id", ""),
        "Tên UUT":           data.get("ten_uut", ""),
        "Khách hàng":        data.get("khach_hang", ""),
        "Phiếu YCCV":        "",
        "Người thực hiện":   data.get("nguoi_thuc_hien", ""),
        "P.pháp HC":         "DLVN76",
        "Ngày hiệu chuẩn":  data.get("ngay_hc", ""),
        "Kết quả HC":        data.get("ket_qua", "OK"),
        "Tem hiệu chuẩn":   data.get("tem_hc", ""),
        "Ngày HC kế tiếp":  data.get("ngay_ke_tiep", ""),
        "TB Chuẩn 1":        data.get("tb_chuan_1", ""),
    }
    return pd.DataFrame([row])


def _build_sheet2_df(data: dict) -> pd.DataFrame:
    gcn_so    = data.get("gcn_so", "")
    ma_id     = data.get("ma_id", "")
    don_vi    = data.get("don_vi", "")
    range_min = data.get("range_min")
    range_max = data.get("range_max")
    points    = data.get("points", [])

    rows = []
    for i, pt in enumerate(points):
        point_id = pt.get("point_id", f"D{i+1}")
        rows.append({
            "Mã Phụ":         f"{gcn_so}{point_id}",
            "GCN Số":         gcn_so,
            "Mã QL / Mã ID":  ma_id,
            "Đ.vị":           don_vi,
            "Min":            range_min if i == 0 else "",
            "Max":            range_max if i == 0 else "",
            "Điểm HC":        point_id,
            "Đơn vị P":       don_vi,
            "P":              pt.get("p_value"),
            "Đơn vị Chuẩn P": don_vi,
            "P c.tăng":       pt.get("p_tang"),
            "P c.giảm":       pt.get("p_giam"),
        })

    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=SHEET2_HEADERS)


def _df_from_session_sheet1() -> pd.DataFrame:
    return _build_sheet1_df(st.session_state.get("extracted_data", {}))


def _df_from_session_sheet2() -> pd.DataFrame:
    return _build_sheet2_df(st.session_state.get("extracted_data", {}))


# ─────────────────────────────────────────────────────────────────────────────
# HELPER: Apply user edits back into the canonical extracted_data dict
# ─────────────────────────────────────────────────────────────────────────────

def _apply_sheet1_edits(df: pd.DataFrame, data: dict) -> dict:
    if df.empty:
        return data
    row = df.iloc[0]
    data["gcn_so"]          = str(row.get("GCN Số", ""))
    data["ma_id"]           = str(row.get("Mã ID", ""))
    data["ten_uut"]         = str(row.get("Tên UUT", ""))
    data["khach_hang"]      = str(row.get("Khách hàng", ""))
    data["nguoi_thuc_hien"] = str(row.get("Người thực hiện", ""))
    data["ngay_hc"]         = str(row.get("Ngày hiệu chuẩn", ""))
    data["ket_qua"]         = str(row.get("Kết quả HC", "OK"))
    data["tem_hc"]          = str(row.get("Tem hiệu chuẩn", ""))
    data["ngay_ke_tiep"]    = str(row.get("Ngày HC kế tiếp", ""))
    data["tb_chuan_1"]      = str(row.get("TB Chuẩn 1", ""))
    return data


def _apply_sheet2_edits(df: pd.DataFrame, data: dict) -> dict:
    if df.empty:
        return data
    if not df.empty and "Đ.vị" in df.columns and not df["Đ.vị"].empty:
        data["don_vi"] = str(df.iloc[0].get("Đ.vị", data.get("don_vi", "")))
    first_row = df[df["Min"].astype(str).str.strip() != ""]
    if not first_row.empty:
        data["range_min"] = _safe_float(first_row.iloc[0].get("Min"))
        data["range_max"] = _safe_float(first_row.iloc[0].get("Max"))

    points = []
    for _, row in df.iterrows():
        points.append({
            "point_id": str(row.get("Điểm HC", "")),
            "p_value":  _safe_float(row.get("P")),
            "p_tang":   _safe_float(row.get("P c.tăng")),
            "p_giam":   _safe_float(row.get("P c.giảm")),
        })
    data["points"] = points
    return data


# ─────────────────────────────────────────────────────────────────────────────
# MAIN STREAMLIT APPLICATION
# ─────────────────────────────────────────────────────────────────────────────

def main():
    # ── Session state initialisation ──────────────────────────────────────
    if "extracted_data" not in st.session_state:
        st.session_state.extracted_data = {}
    if "extraction_done" not in st.session_state:
        st.session_state.extraction_done = False
    if "excel_bytes" not in st.session_state:
        st.session_state.excel_bytes = None
    if "processing" not in st.session_state:
        st.session_state.processing = False
    if "available_models" not in st.session_state:
        st.session_state.available_models = GEMINI_MODELS
    if "selected_model" not in st.session_state:
        st.session_state.selected_model = GEMINI_MODELS[0]

    # ── SIDEBAR ───────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown(
            "<div style='text-align:center; padding: 10px 0 20px;'>"
            "<span style='font-size:2.4rem;'>🔬</span><br>"
            "<span style='font-size:1.05rem; font-weight:700; color:#63b3ed;'>"
            "CalibParser AI</span><br>"
            "<span style='font-size:.75rem; color:rgba(255,255,255,.4);'>"
            "Pressure Calibration Suite</span></div>",
            unsafe_allow_html=True,
        )
        st.divider()

        # ── API Key section ──────────────────────────────────────────────
        st.markdown(
            "<div class='section-header'>🔑 Google AI Configuration</div>",
            unsafe_allow_html=True,
        )

        env_key = os.environ.get("GOOGLE_API_KEY", "")
        use_env = st.checkbox(
            "Load key from environment variable",
            value=bool(env_key),
            help="Uses the `GOOGLE_API_KEY` environment variable if set.",
        )

        if use_env and env_key:
            api_key = env_key
            st.success("✅ API key loaded from environment", icon="🔑")
        else:
            api_key = st.text_input(
                "Google API Key",
                type="password",
                placeholder="AIzaSy...",
                value="" if use_env else "",
            )
            if api_key:
                st.success("API key entered", icon="✅")
            else:
                st.info("Enter your Google AI Studio API key above.", icon="ℹ️")

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Model selection ──────────────────────────────────────────────
        st.markdown(
            "<div class='section-header'>🤖 Model Settings</div>",
            unsafe_allow_html=True,
        )

        # Fetch available models button
        fetch_col, _ = st.columns([3, 1])
        with fetch_col:
            if st.button(
                "🔍 Lấy danh sách model",
                use_container_width=True,
                disabled=not api_key,
                help="Tự động lấy danh sách model đang hoạt động từ API key của bạn",
            ):
                with st.spinner("Đang kết nối Gemini API..."):
                    try:
                        fetched = fetch_available_models(api_key)
                        if fetched:
                            st.session_state.available_models = fetched
                            # keep current selection if still valid
                            if st.session_state.selected_model not in fetched:
                                st.session_state.selected_model = fetched[0]
                            st.success(f"✅ Tìm thấy {len(fetched)} model", icon="🤖")
                        else:
                            st.warning("Không tìm thấy model nào hỗ trợ generateContent.")
                    except Exception as e:
                        st.error(str(e))

        # Dynamic selectbox using fetched list
        current_models = st.session_state.available_models
        cur_idx = (
            current_models.index(st.session_state.selected_model)
            if st.session_state.selected_model in current_models
            else 0
        )
        selected = st.selectbox(
            "Chọn model",
            current_models,
            index=cur_idx,
            help="Nhấn 'Lấy danh sách model' để cập nhật danh sách theo API key của bạn.",
        )
        st.session_state.selected_model = selected

        # Manual override
        custom_model = st.text_input(
            "Hoặc nhập tên model thủ công",
            placeholder="vd: gemini-2.0-flash",
            help="Nhập chính xác tên model nếu không có trong danh sách trên.",
        )
        model_name = custom_model.strip() if custom_model.strip() else selected

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Excel template upload ────────────────────────────────────────
        st.markdown(
            "<div class='section-header'>📊 Excel Template</div>",
            unsafe_allow_html=True,
        )
        excel_file = st.file_uploader(
            "Upload Excel Template (.xlsx)",
            type=["xlsx"],
            help="Upload your existing Excel template (2 sheets). Data will be appended while preserving formatting.",
            key="excel_uploader",
        )

        if excel_file:
            st.success(f"📁 `{excel_file.name}` loaded", icon="✅")
            st.caption(f"Size: {excel_file.size / 1024:.1f} KB")

        st.divider()
        st.markdown(
            "<div style='text-align:center; font-size:.72rem; color:rgba(255,255,255,.3);'>"
            "Powered by Google Gemini AI<br>& openpyxl</div>",
            unsafe_allow_html=True,
        )

    # ── HERO BANNER ───────────────────────────────────────────────────────
    st.markdown(
        """
        <div class="hero-banner">
            <div class="hero-title">🔬 Pressure Calibration Parser</div>
            <div class="hero-sub">
                AI tự động đọc phiếu hiệu chuẩn áp suất viết tay → xuất dữ liệu vào Excel
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── HƯỚNG DẪN SỬ DỤNG ────────────────────────────────────────────────
    with st.expander("📖  Hướng dẫn sử dụng — click để mở", expanded=False):
        st.markdown(
            """
            <div style='padding:4px 0 8px;'>
            <div style='display:grid; grid-template-columns:1fr 1fr; gap:16px;'>

              <div style='background:rgba(99,179,237,0.07); border:1px solid rgba(99,179,237,0.2);
                          border-radius:12px; padding:16px;'>
                <div style='font-size:1rem; font-weight:700; color:#63b3ed; margin-bottom:10px;'>
                  🔧 Chuẩn bị trước
                </div>
                <p style='color:#c8d4ee; font-size:.88rem; line-height:1.8; margin:0;'>
                  <b style='color:#a78bfa;'>1. API Key:</b> Đăng ký miễn phí tại
                  <a href='https://aistudio.google.com' target='_blank'
                     style='color:#63b3ed;'>aistudio.google.com</a>
                  → vào <i>Get API Key</i> → sao chép key và dán vào ô <b>Google API Key</b> ở sidebar bên trái.<br><br>
                  <b style='color:#a78bfa;'>2. File Excel mẫu:</b> Upload file <code>.xlsx</code>
                  có sẵn 2 sheet (Sheet 1: danh mục, Sheet 2: số liệu đo).
                  Ứng dụng sẽ thêm dòng mới vào cuối mỗi sheet, không xóa dữ liệu cũ.
                </p>
              </div>

              <div style='background:rgba(167,139,250,0.07); border:1px solid rgba(167,139,250,0.2);
                          border-radius:12px; padding:16px;'>
                <div style='font-size:1rem; font-weight:700; color:#a78bfa; margin-bottom:10px;'>
                  ⚡ Các bước thực hiện
                </div>
                <p style='color:#c8d4ee; font-size:.88rem; line-height:1.9; margin:0;'>
                  <b style='color:#63b3ed;'>① Upload phiếu:</b>
                  Kéo thả hoặc click chọn file PDF / ảnh chụp phiếu hiệu chuẩn.<br>
                  <b style='color:#63b3ed;'>② Trích xuất AI:</b>
                  Nhấn nút <b>⚡ Extract Data with AI</b> — Gemini sẽ đọc và trả về dữ liệu có cấu trúc.<br>
                  <b style='color:#63b3ed;'>③ Kiểm tra & sửa:</b>
                  Xem lại 2 bảng dữ liệu, click vào ô bất kỳ để sửa tay nếu AI đọc sai.<br>
                  <b style='color:#63b3ed;'>④ Lưu Excel:</b>
                  Nhấn <b>💾 Save &amp; Append to Excel</b> → tải file đã cập nhật về máy.
                </p>
              </div>

            </div>
            <div style='margin-top:14px; background:rgba(246,135,179,0.07);
                        border:1px solid rgba(246,135,179,0.2); border-radius:10px; padding:14px;'>
              <span style='color:#f687b3; font-weight:600;'>💡 Lưu ý:</span>
              <span style='color:#c8d4ee; font-size:.87rem;'>
                Gemini AI hỗ trợ đọc chữ viết tay tiếng Việt.
                Nếu kết quả sai, hãy chọn model <b style='color:#a78bfa;'>gemini-1.5-pro</b> (chính xác hơn)
                hoặc chụp ảnh phiếu với ánh sáng tốt, độ phân giải cao.
              </span>
            </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # ── STEP INDICATORS ───────────────────────────────────────────────────
    col_s1, col_s2, col_s3 = st.columns(3)
    with col_s1:
        st.markdown(
            '<div class="metric-pill">'
            '<div class="metric-value">①</div>'
            '<div class="metric-label">Upload phiếu hiệu chuẩn</div>'
            "</div>",
            unsafe_allow_html=True,
        )
    with col_s2:
        st.markdown(
            '<div class="metric-pill">'
            '<div class="metric-value">②</div>'
            '<div class="metric-label">AI trích xuất dữ liệu</div>'
            "</div>",
            unsafe_allow_html=True,
        )
    with col_s3:
        st.markdown(
            '<div class="metric-pill">'
            '<div class="metric-value">③</div>'
            '<div class="metric-label">Kiểm tra &amp; lưu Excel</div>'
            "</div>",
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # ─────────────────────────────────────────────────────────────────────
    # SECTION 1 — Document Upload & Extraction
    # ─────────────────────────────────────────────────────────────────────
    st.markdown(
        "<div class='glass-card'>"
        "<div class='section-header'>📄 Bước 1 — Upload phiếu hiệu chuẩn</div>"
        "<p style='color:#94a3c0; font-size:.87rem; margin:-6px 0 14px;'>"
        "Kéo thả file vào ô bên dưới, hoặc click <b style='color:#c8d4ee;'>Browse files</b> để chọn từ máy tính. "
        "Hỗ trợ: <b style='color:#63b3ed;'>PDF</b> (nhiều trang), "
        "<b style='color:#63b3ed;'>PNG</b>, <b style='color:#63b3ed;'>JPG/JPEG</b>."
        "</p>",
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader(
        "Drop your calibration report here",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=False,
        help="Supports: PDF (multi-page), PNG, JPG/JPEG. Handwritten or printed.",
        key="doc_uploader",
        label_visibility="collapsed",
    )

    if uploaded_file:
        col_info1, col_info2, col_info3 = st.columns(3)
        with col_info1:
            st.info(f"📄 **File:** `{uploaded_file.name}`")
        with col_info2:
            st.info(f"📦 **Size:** {uploaded_file.size / 1024:.1f} KB")
        with col_info3:
            st.info(f"🗂️ **Type:** `{uploaded_file.type}`")

    st.markdown("</div>", unsafe_allow_html=True)

    # ─────────────────────────────────────────────────────────────────────
    # SECTION 2 — AI Extraction Trigger
    # ─────────────────────────────────────────────────────────────────────
    st.markdown("<div class='glass-card'>", unsafe_allow_html=True)
    st.markdown(
        "<div class='section-header'>🤖 Step 2 — Extract Data with AI</div>",
        unsafe_allow_html=True,
    )

    extract_col, spacer = st.columns([1, 3])
    with extract_col:
        extract_btn = st.button(
            "⚡ Extract Data with AI",
            use_container_width=True,
            type="primary",
            disabled=(uploaded_file is None or not api_key),
        )

    if not api_key:
        st.warning("⚠️ Please enter your Google API key in the sidebar first.", icon="⚠️")
    if uploaded_file is None:
        st.info("ℹ️ Upload a calibration document above to begin.", icon="ℹ️")

    if extract_btn and uploaded_file and api_key:
        file_bytes = uploaded_file.read()
        mime_type = uploaded_file.type or "application/octet-stream"

        with st.spinner("🔍 Analysing document with Gemini AI — please wait…"):
            try:
                result = extract_data_from_document(
                    file_bytes=file_bytes,
                    mime_type=mime_type,
                    api_key=api_key,
                    model_name=model_name,
                )
                st.session_state.extracted_data = result
                st.session_state.extraction_done = True
                st.session_state.excel_bytes = None  # reset previous output
                st.success(
                    f"✅ Extraction complete! Found **{len(result.get('points', []))}** "
                    "calibration points.",
                    icon="🎉",
                )
            except Exception as exc:
                st.error(f"❌ Extraction failed: {exc}", icon="🚨")
                with st.expander("📋 Full traceback"):
                    st.code(traceback.format_exc(), language="python")

    st.markdown("</div>", unsafe_allow_html=True)

    # ─────────────────────────────────────────────────────────────────────
    # SECTION 3 — Review & Edit Extracted Data
    # ─────────────────────────────────────────────────────────────────────
    if st.session_state.extraction_done and st.session_state.extracted_data:
        st.markdown("<div class='glass-card'>", unsafe_allow_html=True)
        st.markdown(
            "<div class='section-header'>✏️ Step 3 — Review & Edit Extracted Data</div>",
            unsafe_allow_html=True,
        )

        # ── Raw JSON toggle ──────────────────────────────────────────────
        with st.expander("🔍 View Raw AI JSON Response"):
            st.json(st.session_state.extracted_data)

        st.markdown("---")

        # ── SHEET 1 EDITOR ───────────────────────────────────────────────
        st.markdown("#### 📋 Sheet 1 — General / Header Information")
        st.caption(
            "Review and correct the general calibration report information below. "
            "Click any cell to edit."
        )

        df_s1 = _build_sheet1_df(st.session_state.extracted_data)
        edited_s1 = st.data_editor(
            df_s1,
            use_container_width=True,
            num_rows="fixed",
            hide_index=True,
            key="editor_sheet1",
            column_config={
                "GCN Số":           st.column_config.TextColumn("GCN Số", width="medium"),
                "Mã ID":            st.column_config.TextColumn("Mã ID", width="medium"),
                "Mã số nhận dạng":  st.column_config.TextColumn("Mã số nhận dạng", width="medium"),
                "Tên UUT":          st.column_config.TextColumn("Tên UUT", width="small"),
                "Khách hàng":       st.column_config.TextColumn("Khách hàng", width="medium"),
                "Phiếu YCCV":       st.column_config.TextColumn("Phiếu YCCV", width="small"),
                "Người thực hiện":  st.column_config.TextColumn("Người thực hiện", width="medium"),
                "P.pháp HC":        st.column_config.TextColumn("P.pháp HC", width="small"),
                "Ngày hiệu chuẩn": st.column_config.TextColumn("Ngày HC", width="medium"),
                "Kết quả HC":       st.column_config.SelectboxColumn(
                    "Kết quả HC",
                    options=["OK", "FAIL", "N/A"],
                    width="small",
                ),
                "Tem hiệu chuẩn":  st.column_config.TextColumn("Tem HC", width="medium"),
                "Ngày HC kế tiếp": st.column_config.TextColumn("Ngày kế tiếp", width="medium"),
                "TB Chuẩn 1":      st.column_config.TextColumn("TB Chuẩn 1", width="medium"),
            },
        )

        st.markdown("<br>", unsafe_allow_html=True)

        # ── SHEET 2 EDITOR ───────────────────────────────────────────────
        st.markdown("#### 📊 Sheet 2 — Measurement Points")
        st.caption(
            "Review the extracted calibration measurement table. "
            "You can add, remove, or correct individual readings."
        )

        df_s2 = _build_sheet2_df(st.session_state.extracted_data)
        edited_s2 = st.data_editor(
            df_s2,
            use_container_width=True,
            num_rows="dynamic",
            hide_index=True,
            key="editor_sheet2",
            column_config={
                "Mã Phụ":          st.column_config.TextColumn("Mã Phụ", width="large"),
                "GCN Số":          st.column_config.TextColumn("GCN Số", width="medium"),
                "Mã QL / Mã ID":   st.column_config.TextColumn("Mã QL", width="medium"),
                "Đ.vị":            st.column_config.TextColumn("Đ.vị", width="small"),
                "Min":             st.column_config.NumberColumn("Min", format="%.2f"),
                "Max":             st.column_config.NumberColumn("Max", format="%.2f"),
                "Điểm HC":         st.column_config.TextColumn("Điểm HC", width="small"),
                "Đơn vị P":        st.column_config.TextColumn("Đơn vị P", width="small"),
                "P":               st.column_config.NumberColumn("P", format="%.4f"),
                "Đơn vị Chuẩn P":  st.column_config.TextColumn("Đơn vị Chuẩn P", width="small"),
                "P c.tăng":        st.column_config.NumberColumn("P c.tăng", format="%.4f"),
                "P c.giảm":        st.column_config.NumberColumn("P c.giảm", format="%.4f"),
            },
        )

        st.markdown("</div>", unsafe_allow_html=True)

        # ── Summary metrics ───────────────────────────────────────────────
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.metric("📌 GCN Số", st.session_state.extracted_data.get("gcn_so", "—"))
        with m2:
            st.metric("🔖 Mã ID", st.session_state.extracted_data.get("ma_id", "—"))
        with m3:
            n_pts = len(st.session_state.extracted_data.get("points", []))
            st.metric("📐 Calibration Points", n_pts)
        with m4:
            st.metric("📏 Range", (
                f"{st.session_state.extracted_data.get('range_min', '?')} – "
                f"{st.session_state.extracted_data.get('range_max', '?')} "
                f"{st.session_state.extracted_data.get('don_vi', '')}"
            ))

        st.markdown("<br>", unsafe_allow_html=True)

        # ─────────────────────────────────────────────────────────────────
        # SECTION 4 — Save to Excel
        # ─────────────────────────────────────────────────────────────────
        st.markdown("<div class='glass-card'>", unsafe_allow_html=True)
        st.markdown(
            "<div class='section-header'>💾 Step 4 — Save & Append to Excel</div>",
            unsafe_allow_html=True,
        )

        if excel_file is None:
            st.warning(
                "⚠️ No Excel template uploaded. Please upload your `.xlsx` template in the sidebar.",
                icon="📊",
            )
        else:
            save_col, dl_col = st.columns([1, 1])
            with save_col:
                save_btn = st.button(
                    "💾 Save & Append to Excel",
                    use_container_width=True,
                    type="primary",
                )

            if save_btn:
                # Apply user edits back into extracted_data
                updated_data = dict(st.session_state.extracted_data)
                updated_data = _apply_sheet1_edits(edited_s1, updated_data)
                updated_data = _apply_sheet2_edits(edited_s2, updated_data)
                st.session_state.extracted_data = updated_data

                with st.spinner("⚙️ Writing data to Excel — please wait…"):
                    try:
                        # Save excel_file to a temp file for openpyxl
                        with tempfile.NamedTemporaryFile(
                            suffix=".xlsx", delete=False
                        ) as tmp:
                            tmp.write(excel_file.read())
                            tmp_path = tmp.name

                        excel_bytes = append_to_excel(tmp_path, updated_data)
                        st.session_state.excel_bytes = excel_bytes

                        # Clean up temp file
                        try:
                            os.unlink(tmp_path)
                        except OSError:
                            pass

                        st.success(
                            "✅ Data successfully appended to Excel template!",
                            icon="🎉",
                        )
                    except Exception as exc:
                        st.error(f"❌ Failed to write Excel: {exc}", icon="🚨")
                        with st.expander("📋 Full traceback"):
                            st.code(traceback.format_exc(), language="python")

            # ── Download button ─────────────────────────────────────────
            if st.session_state.excel_bytes:
                gcn = st.session_state.extracted_data.get("gcn_so", "output")
                safe_gcn = re.sub(r"[\\/*?:\"<>|]", "_", gcn)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"CalibReport_{safe_gcn}_{ts}.xlsx"

                with dl_col:
                    st.download_button(
                        label="⬇️ Download Updated Excel",
                        data=st.session_state.excel_bytes,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="secondary",
                    )

                st.info(
                    f"📥 Your file **`{filename}`** is ready for download. "
                    "All existing formatting, formulas, and other data in your template are preserved.",
                    icon="✅",
                )

        st.markdown("</div>", unsafe_allow_html=True)

    # ── Footer ────────────────────────────────────────────────────────────
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown(
        "<div style='text-align:center; color:rgba(255,255,255,.25); font-size:.75rem;'>"
        "CalibParser AI · Built with Streamlit + Google Gemini + openpyxl"
        "</div>",
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    main()
