import streamlit as st
import io
import os
import zipfile

import openpyxl
from PIL import Image
from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib import colors

# ─── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Aquarius Document Generator",
    page_icon="📄",
    layout="centered",
)

# ─── Layout constants (pixel-measured from the 1316×924 background template) ─
PAGE_W, PAGE_H = landscape((420 * mm, 297 * mm))
LEFT_MM    = 17.87;  RIGHT_MM   = 399.57
CTOP_MM    = 279.32; CBOT_MM    = 42.11
FCENTER_MM = 30.05
AVAIL_W_MM = RIGHT_MM - LEFT_MM   # 381.70 mm
AVAIL_H_MM = CTOP_MM  - CBOT_MM   # 237.21 mm

# ─── Background template (bundled with the app) ───────────────────────────────
BG_PATH = os.path.join(os.path.dirname(__file__), "Background_pdf.pdf")

@st.cache_data
def load_background():
    with zipfile.ZipFile(BG_PATH, "r") as z:
        name = next(n for n in z.namelist() if n.lower().endswith((".jpeg", ".jpg", ".png")))
        return z.read(name)

# ─── Helpers ─────────────────────────────────────────────────────────────────

def extract_image_from_drawing(file_bytes: bytes) -> bytes:
    """Extract an image from either a ZIP-wrapped PDF or a plain PDF."""
    
    # 1. Try ZIP first (some CAD-export PDFs are ZIP containers)
    if zipfile.is_zipfile(io.BytesIO(file_bytes)):
        with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as z:
            name = next(
                (n for n in z.namelist() if n.lower().endswith((".jpeg", ".jpg", ".png"))),
                None,
            )
            if name:
                return z.read(name)

    # 2. Try to rasterize the first page of the PDF → PNG in memory
    try:
        import pdf2image
        images = pdf2image.convert_from_bytes(file_bytes, dpi=150, first_page=1, last_page=1)
        if images:
            buf = io.BytesIO()
            images[0].save(buf, format="PNG")
            return buf.getvalue()
    except Exception:
        pass

    # 3. Fallback: try extracting an embedded image via pypdf
    try:
        import pypdf
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        for page in reader.pages:
            for img_obj in page.images:
                return img_obj.data
    except Exception:
        pass

    raise ValueError("Could not extract an image from the uploaded drawing PDF.")


def read_excel_rows(file_bytes: bytes) -> list:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active
    rows = []
    for row in ws.iter_rows(values_only=True):
        if any(v is not None for v in row):
            rows.append([str(v) if v is not None else "" for v in row[:5]])
    return rows


def derive_doc_name(filename: str) -> str:
    base = os.path.splitext(filename)[0]
    return base.replace("__", "  ").replace("_", " ")

# ─── PDF generation ───────────────────────────────────────────────────────────

def draw_background(c, bg_bytes):
    c.drawImage(ImageReader(io.BytesIO(bg_bytes)), 0, 0,
                width=PAGE_W, height=PAGE_H, preserveAspectRatio=False)


def draw_footer_centered(c, doc_name):
    cx = (LEFT_MM + RIGHT_MM) / 2 * mm
    c.setFont("Helvetica-Bold", 10)
    c.setFillColorRGB(0, 0, 0)
    c.drawCentredString(cx, FCENTER_MM * mm, doc_name)

def draw_excel_table(c, rows):
    aw = AVAIL_W_MM * mm
    ah = AVAIL_H_MM * mm
    cw = [aw * 0.04, aw * 0.15, aw * 0.05, aw * 0.20, aw * 0.56] # Widened to prevent wrap
    
    # FIX: Increased leading to 20 so 15pt font doesn't overlap
    cs = ParagraphStyle("c", fontName="Helvetica",      fontSize=15, leading=20)
    hs = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=15, leading=20)
    
    data = [
        [Paragraph(str(cell), hs if i == 0 else cs) for cell in row]
        for i, row in enumerate(rows)
    ]
    
    t = Table(data, colWidths=cw)
    t.setStyle(TableStyle([
        ("TEXTCOLOR",     (0, 0), (-1, -1), colors.black),
        # REMOVED: Black line below header is gone
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING",   (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
    ]))
    
    _, th = t.wrapOn(c, aw, ah)
    t.drawOn(c, LEFT_MM * mm, CTOP_MM * mm - th)

def draw_content_image(c, content_bytes):
    aw = AVAIL_W_MM * mm
    ah = AVAIL_H_MM * mm
    pil = Image.open(io.BytesIO(content_bytes))
    asp = pil.width / pil.height
    dw, dh = (aw, aw / asp) if aw / asp <= ah else (ah * asp, ah)
    ix = LEFT_MM * mm + (aw - dw) / 2
    iy = CBOT_MM * mm + (ah - dh) / 2
    c.drawImage(ImageReader(io.BytesIO(content_bytes)), ix, iy,
                width=dw, height=dh, preserveAspectRatio=True)


def generate_pdf(xlsx_bytes, xlsx_name, content_bytes, doc_name) -> bytes:
    bg_bytes = load_background()
    rows     = read_excel_rows(xlsx_bytes)
    img_bytes = extract_image_from_drawing(content_bytes)

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))

    # Page 1 — Excel table
    draw_background(c, bg_bytes)
    draw_excel_table(c, rows)
    draw_footer_centered(c, doc_name)
    c.showPage()

    # Page 2 — Engineering drawing
    draw_background(c, bg_bytes)
    draw_content_image(c, img_bytes)
    draw_footer_centered(c, doc_name)
    c.showPage()

    c.save()
    return buf.getvalue()

# ─── UI ───────────────────────────────────────────────────────────────────────

# Header
col1, col2 = st.columns([1, 6])
with col1:
    st.markdown("## 📄")
with col2:
    st.markdown("## Aquarius Document Generator")
    st.caption("Upload your Excel parts list and engineering drawing PDF to generate a branded 2-page document.")

st.divider()

# Upload section
col_a, col_b = st.columns(2)

with col_a:
    st.markdown("#### 📊 Parts List")
    xlsx_file = st.file_uploader(
        "Upload Excel file",
        type=["xlsx", "xls"],
        key="xlsx",
        help="The parts/BOM spreadsheet (.xlsx)",
        label_visibility="collapsed",
    )
    if xlsx_file:
        st.success(f"✅ {xlsx_file.name}  ({xlsx_file.size // 1024} KB)")

with col_b:
    st.markdown("#### 🖼️ Engineering Drawing")
    pdf_file = st.file_uploader(
        "Upload Drawing PDF",
        type=["pdf"],
        key="pdf",
        help="The engineering drawing PDF",
        label_visibility="collapsed",
    )
    if pdf_file:
        st.success(f"✅ {pdf_file.name}  ({pdf_file.size // 1024} KB)")

st.markdown("---")

# Document name (auto-filled, editable)
if xlsx_file:
    default_name = derive_doc_name(xlsx_file.name)
else:
    default_name = ""

doc_name = st.text_input(
    "📝 Document name (shown in footer)",
    value=default_name,
    placeholder="e.g. SPEM02AO01  SB36 Z",
    help="This appears centred in the footer of both pages. Auto-filled from filename.",
)

st.markdown("")

# Preview table
if xlsx_file:
    with st.expander("👁️ Preview Excel data", expanded=False):
        try:
            rows = read_excel_rows(xlsx_file.getvalue())
            if rows:
                import pandas as pd
                df = pd.DataFrame(rows[1:], columns=rows[0] if rows else [])
                st.dataframe(df, use_container_width=True, hide_index=True)
        except Exception as e:
            st.warning(f"Could not preview: {e}")

# Generate button
ready = xlsx_file and pdf_file and doc_name.strip()

if st.button(
    "⚙️  Generate 2-page PDF",
    type="primary",
    use_container_width=True,
    disabled=not ready,
):
    with st.spinner("Generating your document…"):
        try:
            pdf_bytes = generate_pdf(
                xlsx_bytes    = xlsx_file.getvalue(),
                xlsx_name     = xlsx_file.name,
                content_bytes = pdf_file.getvalue(),
                doc_name      = doc_name.strip(),
            )
            out_filename = (doc_name.strip().replace("  ", "__").replace(" ", "_") or "output") + ".pdf"

            st.success("✅ PDF generated successfully!")
            st.balloons()

            st.download_button(
                label     = "⬇️  Download PDF",
                data      = pdf_bytes,
                file_name = out_filename,
                mime      = "application/pdf",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Generation failed: {e}")
            st.exception(e)

# Footer
st.divider()
st.caption("Aquarius Industries · Internal tool · Template is fixed — just upload new Excel + Drawing each time.")
