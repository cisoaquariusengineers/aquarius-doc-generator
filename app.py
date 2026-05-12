import streamlit as st
import io
import os
import zipfile

import openpyxl
from PIL import Image as PILImage
from reportlab.lib.pagesizes import landscape, A3
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

# ─── Page size: A3 landscape ─────────────────────────────────────────────────
PAGE_W, PAGE_H = landscape(A3)   # 420 × 297 mm in points

# Margins (mm) — these define the inner content area
MARGIN_L  = 15 * mm
MARGIN_R  = 15 * mm
MARGIN_T  = 15 * mm
MARGIN_B  = 25 * mm   # taller bottom for title block

CONTENT_X = MARGIN_L
CONTENT_Y = MARGIN_B
CONTENT_W = PAGE_W - MARGIN_L - MARGIN_R
CONTENT_H = PAGE_H - MARGIN_T - MARGIN_B

# Title block (footer strip)
TITLE_H   = 18 * mm
TITLE_Y   = MARGIN_B - TITLE_H   # sits below content area, above page edge

# ─── Helpers ─────────────────────────────────────────────────────────────────

def extract_image_from_drawing(file_bytes: bytes) -> bytes:
    """Return PNG bytes of the first page of the drawing PDF, or raise."""
    errors = []

    # 1. ZIP-wrapped PDF (some CAD exports)
    if zipfile.is_zipfile(io.BytesIO(file_bytes)):
        with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as z:
            name = next(
                (n for n in z.namelist()
                 if n.lower().endswith((".jpeg", ".jpg", ".png"))),
                None,
            )
            if name:
                return z.read(name)
            errors.append(f"ZIP ok but no image. Contents: {z.namelist()}")
    else:
        errors.append("Not a ZIP file.")

    # 2. PyMuPDF — most reliable on Streamlit Cloud (no poppler needed)
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        page = doc[0]
        mat  = fitz.Matrix(2.5, 2.5)   # ~180 dpi — sharp enough for A3
        pix  = page.get_pixmap(matrix=mat, alpha=False)
        return pix.tobytes("png")
    except Exception as e:
        errors.append(f"PyMuPDF failed: {e}")

    # 3. pdf2image / poppler fallback
    try:
        import pdf2image
        imgs = pdf2image.convert_from_bytes(file_bytes, dpi=150,
                                            first_page=1, last_page=1)
        if imgs:
            buf = io.BytesIO()
            imgs[0].save(buf, format="PNG")
            return buf.getvalue()
        errors.append("pdf2image returned no pages.")
    except Exception as e:
        errors.append(f"pdf2image failed: {e}")

    # 4. pypdf embedded image
    try:
        import pypdf
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        for page in reader.pages:
            imgs = list(page.images)
            if imgs:
                return imgs[0].data
        errors.append("pypdf: no embedded images found.")
    except Exception as e:
        errors.append(f"pypdf failed: {e}")

    raise ValueError("Could not extract image from PDF.\n" + "\n".join(errors))


def read_excel_rows(file_bytes: bytes) -> list[list[str]]:
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


# ─── Drawing helpers ──────────────────────────────────────────────────────────

def draw_outer_border(c: canvas.Canvas):
    """Thin outer rectangle matching the reference images."""
    c.setStrokeColorRGB(0, 0, 0)
    c.setLineWidth(1.0)
    bx = 8 * mm
    by = 8 * mm
    bw = PAGE_W - 16 * mm
    bh = PAGE_H - 16 * mm
    c.rect(bx, by, bw, bh)


def draw_inner_content_border(c: canvas.Canvas):
    """Inner border around the drawing / table area."""
    c.setStrokeColorRGB(0, 0, 0)
    c.setLineWidth(0.5)
    c.rect(CONTENT_X, CONTENT_Y, CONTENT_W, CONTENT_H)


def draw_title_block(c: canvas.Canvas, part_no: str, doc_name: str, logo_bytes: bytes | None):
    """
    Bottom title block — two rows:
      Row 1: doc title (large, bold)
      Row 2: part number (smaller)
    Logo in top-right corner of page.
    """
    # Title block background
    c.setFillColorRGB(1, 1, 1)
    c.rect(CONTENT_X, TITLE_Y, CONTENT_W, TITLE_H, fill=1, stroke=0)
    # Top line of title block
    c.setStrokeColorRGB(0, 0, 0)
    c.setLineWidth(0.8)
    c.line(CONTENT_X, CONTENT_Y, CONTENT_X + CONTENT_W, CONTENT_Y)

    cx = CONTENT_X + CONTENT_W / 2
    # Document name (bold)
    c.setFont("Helvetica-Bold", 11)
    c.setFillColorRGB(0, 0, 0)
    c.drawCentredString(cx, TITLE_Y + TITLE_H * 0.58, doc_name)
    # Part number
    c.setFont("Helvetica", 9)
    c.drawCentredString(cx, TITLE_Y + TITLE_H * 0.22, part_no)

    # Logo — top-right of the CONTENT area
    if logo_bytes:
        logo_w = 28 * mm
        logo_h = 12 * mm
        lx = CONTENT_X + CONTENT_W - logo_w - 2 * mm
        ly = CONTENT_Y + CONTENT_H - logo_h - 2 * mm
        c.drawImage(ImageReader(io.BytesIO(logo_bytes)),
                    lx, ly, width=logo_w, height=logo_h,
                    preserveAspectRatio=True, mask="auto")


def draw_watermark(c: canvas.Canvas, text: str = "AQUARIUS"):
    """Diagonal semi-transparent watermark text across the content area."""
    c.saveState()
    cx = CONTENT_X + CONTENT_W / 2
    cy = CONTENT_Y + CONTENT_H / 2
    c.translate(cx, cy)
    c.rotate(35)
    c.setFont("Helvetica-Bold", 72)
    c.setFillColorRGB(0.75, 0.75, 0.75, alpha=0.18)
    c.drawCentredString(0, 0, text)
    c.restoreState()


def draw_page_number(c: canvas.Canvas, current: int, total: int):
    """Small page number at bottom-right."""
    c.setFont("Helvetica", 8)
    c.setFillColorRGB(0, 0, 0)
    label = str(current)
    rx = CONTENT_X + CONTENT_W - 2 * mm
    ry = TITLE_Y + 3 * mm
    c.drawRightString(rx, ry, label)


# ─── Page renderers ───────────────────────────────────────────────────────────

def render_drawing_page(c: canvas.Canvas, img_bytes: bytes,
                        part_no: str, doc_name: str, logo_bytes: bytes | None,
                        page_num: int, total: int):
    draw_outer_border(c)
    draw_inner_content_border(c)
    draw_watermark(c)

    # Centre the drawing image inside the content area
    pil = PILImage.open(io.BytesIO(img_bytes))
    asp = pil.width / pil.height
    pad = 4 * mm
    avail_w = CONTENT_W - 2 * pad
    avail_h = CONTENT_H - 2 * pad
    if avail_w / asp <= avail_h:
        dw, dh = avail_w, avail_w / asp
    else:
        dw, dh = avail_h * asp, avail_h
    ix = CONTENT_X + pad + (avail_w - dw) / 2
    iy = CONTENT_Y + pad + (avail_h - dh) / 2
    c.drawImage(ImageReader(io.BytesIO(img_bytes)),
                ix, iy, width=dw, height=dh,
                preserveAspectRatio=True, mask="auto")

    draw_title_block(c, part_no, doc_name, logo_bytes)
    draw_page_number(c, page_num, total)
    c.showPage()


def render_table_page(c: canvas.Canvas, rows: list[list[str]],
                      part_no: str, doc_name: str, logo_bytes: bytes | None,
                      page_num: int, total: int):
    draw_outer_border(c)
    draw_inner_content_border(c)
    draw_watermark(c)

    # ── Build table ──────────────────────────────────────────────────────────
    # Column widths proportional to content: Sr | Part No | Qty | Tech Detail | Description
    cw_frac = [0.055, 0.130, 0.055, 0.185, 0.575]
    cw = [CONTENT_W * f for f in cw_frac]

    header_style = ParagraphStyle(
        "th",
        fontName="Helvetica-Bold",
        fontSize=8,
        leading=10,
        textColor=colors.white,
    )
    cell_style = ParagraphStyle(
        "td",
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        textColor=colors.black,
    )

    data = []
    for i, row in enumerate(rows):
        s = header_style if i == 0 else cell_style
        data.append([Paragraph(cell, s) for cell in row])

    t = Table(data, colWidths=cw, repeatRows=1)
    t.setStyle(TableStyle([
        # Header row
        ("BACKGROUND",     (0, 0), (-1, 0),  colors.HexColor("#1a3a5c")),
        ("TEXTCOLOR",      (0, 0), (-1, 0),  colors.white),
        ("LINEBELOW",      (0, 0), (-1, 0),  1.2, colors.HexColor("#1a3a5c")),
        # Alternating rows
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.white, colors.HexColor("#f0f4f8")]),
        # Grid
        ("GRID",           (0, 0), (-1, -1), 0.4, colors.HexColor("#b0c4d8")),
        # Alignment & padding
        ("VALIGN",         (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",     (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",  (0, 0), (-1, -1), 4),
        ("LEFTPADDING",    (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",   (0, 0), (-1, -1), 5),
    ]))

    # Wrap & draw — top-left of content area
    table_w, table_h = t.wrapOn(c, CONTENT_W, CONTENT_H)
    tx = CONTENT_X
    ty = CONTENT_Y + CONTENT_H - table_h   # pin to top of content area
    t.drawOn(c, tx, ty)

    draw_title_block(c, part_no, doc_name, logo_bytes)
    draw_page_number(c, page_num, total)
    c.showPage()


# ─── Main PDF generator ──────────────────────────────────────────────────────

def generate_pdf(xlsx_bytes: bytes, xlsx_name: str,
                 drawing_bytes: bytes, doc_name: str,
                 logo_bytes: bytes | None) -> bytes:
    rows     = read_excel_rows(xlsx_bytes)
    img_bytes = extract_image_from_drawing(drawing_bytes)

    # Derive part number from first data row (col index 1) or filename
    part_no = rows[1][1] if len(rows) > 1 and len(rows[1]) > 1 else ""

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))

    # Page 1 — Engineering drawing
    render_drawing_page(c, img_bytes, part_no, doc_name, logo_bytes, 1, 2)

    # Page 2 — Parts list table
    render_table_page(c, rows, part_no, doc_name, logo_bytes, 2, 2)

    c.save()
    return buf.getvalue()


# ─── UI ──────────────────────────────────────────────────────────────────────

col1, col2 = st.columns([1, 6])
with col1:
    st.markdown("## 📄")
with col2:
    st.markdown("## Aquarius Document Generator")
    st.caption("Upload your Excel parts list and engineering drawing PDF "
               "to generate a branded 2-page A3 document.")

st.divider()

# ── Uploads ──────────────────────────────────────────────────────────────────
col_a, col_b, col_c = st.columns(3)

with col_a:
    st.markdown("#### 📊 Parts List")
    xlsx_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"],
                                 key="xlsx", label_visibility="collapsed")
    if xlsx_file:
        st.success(f"✅ {xlsx_file.name}  ({xlsx_file.size // 1024} KB)")

with col_b:
    st.markdown("#### 🖼️ Engineering Drawing")
    pdf_file = st.file_uploader("Upload Drawing PDF", type=["pdf"],
                                key="pdf", label_visibility="collapsed")
    if pdf_file:
        st.success(f"✅ {pdf_file.name}  ({pdf_file.size // 1024} KB)")

with col_c:
    st.markdown("#### 🏷️ Logo (optional)")
    logo_file = st.file_uploader("Upload Logo PNG/JPG", type=["png", "jpg", "jpeg"],
                                 key="logo", label_visibility="collapsed")
    if logo_file:
        st.success(f"✅ {logo_file.name}")

st.markdown("---")

# ── Document name ─────────────────────────────────────────────────────────────
default_name = derive_doc_name(xlsx_file.name) if xlsx_file else ""
doc_name = st.text_input(
    "📝 Document name (shown in title block footer)",
    value=default_name,
    placeholder="e.g. SB36 Z PLACER BOOM & ACCESSORIES",
)

st.markdown("")

# ── Preview ───────────────────────────────────────────────────────────────────
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

# ── Generate ──────────────────────────────────────────────────────────────────
ready = xlsx_file and pdf_file and doc_name.strip()

if st.button("⚙️  Generate 2-page PDF", type="primary",
             use_container_width=True, disabled=not ready):
    with st.spinner("Generating your document…"):
        try:
            logo_bytes = logo_file.getvalue() if logo_file else None
            pdf_bytes  = generate_pdf(
                xlsx_bytes    = xlsx_file.getvalue(),
                xlsx_name     = xlsx_file.name,
                drawing_bytes = pdf_file.getvalue(),
                doc_name      = doc_name.strip(),
                logo_bytes    = logo_bytes,
            )
            out_filename = (
                doc_name.strip().replace("  ", "__").replace(" ", "_") or "output"
            ) + ".pdf"

            st.success("✅ PDF generated successfully!")
            st.balloons()
            st.download_button(
                label            = "⬇️  Download PDF",
                data             = pdf_bytes,
                file_name        = out_filename,
                mime             = "application/pdf",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Generation failed: {e}")
            st.exception(e)

st.divider()
st.caption("Aquarius Industries · Internal tool · "
           "Upload new Excel + Drawing each time.")
