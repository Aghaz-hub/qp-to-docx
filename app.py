"""
Question Paper → Word Document Converter
========================================
Upload scanned question paper images or PDFs (single / multi-page).
App runs OCR, structures the content, and generates a clean .docx file
ready for editing / reprinting.
"""

import streamlit as st
from PIL import Image, ImageEnhance, ImageFilter, ImageOps
import pytesseract
import io
import re
from datetime import datetime
from pdf2image import convert_from_bytes
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

st.set_page_config(
    page_title="QP → Word Converter",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

if "ocr_text" not in st.session_state:
    st.session_state.ocr_text = ""
if "docx_bytes" not in st.session_state:
    st.session_state.docx_bytes = None
if "processed" not in st.session_state:
    st.session_state.processed = False


# ------------------------------------------------------------------
# Image / PDF helpers
# ------------------------------------------------------------------
def preprocess(img: Image.Image) -> Image.Image:
    img = img.convert("L")
    img = ImageOps.autocontrast(img, cutoff=1)
    img = ImageEnhance.Contrast(img).enhance(1.7)
    img = img.filter(ImageFilter.SHARPEN)
    return img


def ocr_page(img: Image.Image) -> str:
    img = preprocess(img)
    # Try a couple of PSM modes and keep the longer / cleaner result
    t1 = pytesseract.image_to_string(img, config="--psm 6")
    t2 = pytesseract.image_to_string(img, config="--psm 4")
    return t1 if len(t1) >= len(t2) else t2


def expand_to_images(uploaded_files):
    """Return list of (PIL.Image, label)."""
    pages = []
    for f in uploaded_files:
        name = f.name
        data = f.read()
        lower = name.lower()
        if lower.endswith(".pdf"):
            try:
                pil_pages = convert_from_bytes(data, dpi=250, fmt="jpeg")
                for i, page in enumerate(pil_pages, 1):
                    pages.append((page, f"{name} – page {i}"))
            except Exception as e:
                st.error(f"Could not read PDF '{name}': {e}")
        else:
            try:
                img = Image.open(io.BytesIO(data)).convert("RGB")
                pages.append((img, name))
            except Exception as e:
                st.error(f"Could not open image '{name}': {e}")
    return pages


# ------------------------------------------------------------------
# Light structure extraction from OCR text
# ------------------------------------------------------------------
def extract_metadata(text: str) -> dict:
    meta = {
        "university": "MANAV RACHNA UNIVERSITY",
        "school": "SCHOOL OF ENGINEERING",
        "department": "DEPARTMENT OF COMPUTER SCIENCE & TECHNOLOGY",
        "exam_title": "",
        "semester": "",
        "date": "",
        "credits": "",
        "course_code": "",
        "session": "",
        "program": "",
        "course_name": "",
        "max_marks": "",
        "duration": "",
    }

    # Exam title
    m = re.search(r"[\"']?(Mid\s*Semester\s*Examination[^\"'\n]*)[\"']?", text, re.I)
    if m:
        meta["exam_title"] = m.group(1).strip().strip('"\'')
    else:
        m = re.search(r"(End\s*Term|Supplementary|Mid[\s-]*Sem)[^\n]{0,40}", text, re.I)
        if m:
            meta["exam_title"] = m.group(0).strip()

    # Common key-value patterns
    patterns = {
        "semester": r"SEMESTER\s*[:\|]?\s*([IVX\d]+)",
        "date": r"DATE\s*OF\s*EXAM\s*[:\|]?\s*([\d\.\-/]+)",
        "credits": r"CREDITS?\s*[:\|]?\s*(\d+)",
        "course_code": r"COURSE\s*CODE\s*[:\|]?\s*([A-Z0-9\-]+)",
        "session": r"SESSION\s*[:\|]?\s*([IVX\d]+)",
        "program": r"PROGRAM\s*[:\|]?\s*([A-Z0-9\s\+\(\)\/\-]+?)(?:\n|COURSE|MAX|$)",
        "course_name": r"COURSE\s*NAME\s*[:\|]?\s*([A-Za-z\s&]+?)(?:\n|MAX|TIME|$)",
        "max_marks": r"MAX\.?\s*MARKS\s*[:\|]?\s*(\d+)",
        "duration": r"(?:TIME\s*)?DURATION\s*[:\|]?\s*([\d\s]+MINUTES?)",
    }
    for key, pat in patterns.items():
        m = re.search(pat, text, re.I)
        if m:
            meta[key] = re.sub(r"\s+", " ", m.group(1)).strip()

    return meta


def clean_ocr_text(text: str) -> str:
    """Basic cleanup of common OCR noise."""
    text = text.replace("|", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


# ------------------------------------------------------------------
# Build Word document
# ------------------------------------------------------------------
def set_run_font(run, name="Times New Roman", size=11, bold=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold


def add_horizontal_line(paragraph):
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    pPr.append(pBdr)


def create_docx(full_text: str, meta: dict, source_name: str = "") -> bytes:
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)

    # ---- Header ----
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(meta.get("university") or "MANAV RACHNA UNIVERSITY")
    set_run_font(run, size=16, bold=True)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(meta.get("school") or "SCHOOL OF ENGINEERING")
    set_run_font(run, size=12, bold=True)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(meta.get("department") or "DEPARTMENT OF COMPUTER SCIENCE & TECHNOLOGY")
    set_run_font(run, size=11, bold=True)

    if meta.get("exam_title"):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f'"{meta["exam_title"]}"')
        set_run_font(run, size=12, bold=True)

    add_horizontal_line(doc.add_paragraph())

    # ---- Exam details table ----
    details = [
        ("SEMESTER", meta.get("semester", ""), "DATE OF EXAM", meta.get("date", "")),
        ("COURSE CODE", meta.get("course_code", ""), "SESSION", meta.get("session", "")),
        ("COURSE NAME", meta.get("course_name", ""), "MAX. MARKS", meta.get("max_marks", "")),
        ("PROGRAM", meta.get("program", ""), "TIME DURATION", meta.get("duration", "")),
        ("CREDITS", meta.get("credits", ""), "", ""),
    ]

    table = doc.add_table(rows=len(details), cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (k1, v1, k2, v2) in enumerate(details):
        row = table.rows[i]
        for j, val in enumerate([k1, v1, k2, v2]):
            cell = row.cells[j]
            cell.text = val
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    set_run_font(run, size=10, bold=(j % 2 == 0))

    doc.add_paragraph()

    # ---- Note ----
    p = doc.add_paragraph()
    run = p.add_run("Note: All questions are compulsory.")
    set_run_font(run, size=10, bold=True)

    doc.add_paragraph()

    # ---- Full OCR content (cleaned) under a heading ----
    p = doc.add_paragraph()
    run = p.add_run("QUESTION PAPER CONTENT (OCR Extracted)")
    set_run_font(run, size=12, bold=True)
    add_horizontal_line(doc.add_paragraph())

    # Split into paragraphs and add
    blocks = re.split(r"\n{2,}", full_text)
    for block in blocks:
        block = block.strip()
        if not block:
            continue
        # Detect PART headings
        if re.match(r"^PART\s*[-–]?\s*[A-Z]", block, re.I):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(block.split("\n")[0].upper())
            set_run_font(run, size=12, bold=True)
            rest = "\n".join(block.split("\n")[1:]).strip()
            if rest:
                p = doc.add_paragraph()
                run = p.add_run(rest)
                set_run_font(run, size=11)
        else:
            p = doc.add_paragraph()
            run = p.add_run(block)
            set_run_font(run, size=11)

    # ---- Footer note ----
    doc.add_paragraph()
    add_horizontal_line(doc.add_paragraph())
    p = doc.add_paragraph()
    run = p.add_run(
        f"Generated by QP → Word Converter on {datetime.now().strftime('%Y-%m-%d %H:%M')} "
        f"| Source: {source_name or 'uploaded file(s)'} | "
        "Please proof-read OCR text before final use."
    )
    set_run_font(run, size=8)
    run.italic = True

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ------------------------------------------------------------------
# Sidebar
# ------------------------------------------------------------------
with st.sidebar:
    st.title("📄 QP → Word")
    st.markdown("---")
    st.markdown(
        """
**How to use**

1. Upload scanned question paper  
   (JPG / PNG / PDF – single or multi-page)
2. Click **Convert to Word**
3. Download the `.docx` file
4. Open in MS Word / Google Docs and proof-read

**Tips for better OCR**
- Use clear, well-lit scans
- Prefer 200–300 DPI
- Avoid heavy shadows / skew
        """
    )
    st.markdown("---")
    st.caption("Works best with Manav Rachna / similar university question-paper formats.")


# ------------------------------------------------------------------
# Main UI
# ------------------------------------------------------------------
st.title("Question Paper → Word Document Converter")
st.caption(
    "Upload old scanned question papers (images or PDF). "
    "The app extracts text with OCR and builds a clean editable Word document."
)

uploaded = st.file_uploader(
    "Upload question paper – Images (JPG/PNG) or PDF (single / multi-page)",
    type=["jpg", "jpeg", "png", "webp", "pdf"],
    accept_multiple_files=True,
    help="You can upload multiple images or a multi-page PDF of the question paper.",
)

if uploaded:
    st.success(f"{len(uploaded)} file(s) selected")

    if st.button("🚀 Convert to Word Document", type="primary", use_container_width=True):
        with st.spinner("Converting pages and running OCR…"):
            pages = expand_to_images(uploaded)
            if not pages:
                st.error("No readable pages found.")
            else:
                all_text_parts = []
                progress = st.progress(0, text="Starting OCR…")
                for idx, (img, label) in enumerate(pages):
                    progress.progress(idx / len(pages), text=f"OCR page {idx+1}/{len(pages)}: {label}")
                    text = ocr_page(img)
                    all_text_parts.append(f"--- PAGE {idx+1}: {label} ---\n{text}")
                progress.progress(1.0, text="Building Word document…")

                full_text = clean_ocr_text("\n\n".join(all_text_parts))
                meta = extract_metadata(full_text)
                source_names = ", ".join(f.name for f in uploaded)
                docx_bytes = create_docx(full_text, meta, source_names)

                st.session_state.ocr_text = full_text
                st.session_state.docx_bytes = docx_bytes
                st.session_state.processed = True
                st.rerun()

if st.session_state.processed and st.session_state.docx_bytes:
    st.markdown("---")
    st.subheader("Conversion complete")

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="📥 Download Word Document (.docx)",
            data=st.session_state.docx_bytes,
            file_name=f"Question_Paper_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
    with col2:
        if st.button("🔄 Convert another paper", use_container_width=True):
            st.session_state.ocr_text = ""
            st.session_state.docx_bytes = None
            st.session_state.processed = False
            st.rerun()

    with st.expander("View extracted OCR text (for verification)"):
        st.text_area("OCR Output", st.session_state.ocr_text, height=400)

    st.info(
        "Always proof-read the Word file. OCR on old/faded scans can misread characters, "
        "especially mathematical symbols and small text."
    )

else:
    st.info("Upload one or more scanned question-paper pages (or a PDF), then click **Convert to Word Document**.")

st.markdown("---")
st.caption(
    "Supports images + multi-page PDF • OCR via Tesseract • Output is an editable .docx "
    "with university header, exam details table, and full question text."
)
