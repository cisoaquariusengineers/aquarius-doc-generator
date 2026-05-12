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
    layout="wide",
)

# ─── Layout constants ─────────────────────────────────────────────────────────
PAGE_W, PAGE_H = landscape((420 * mm, 297 * mm))
LEFT_MM    = 17.87;  RIGHT_MM   = 399.57
CTOP_MM    = 279.32; CBOT_MM    = 42.11
FCENTER_MM = 30.05
AVAIL_W_MM = RIGHT_MM - LEFT_MM
AVAIL_H_MM = CTOP_MM  - CBOT_MM

# ─── Background template ──────────────────────────────────────────────────────
BG_PATH = os.path.join(os.path.dirname(__file__), "Background_pdf.pdf")

@st.cache_data
def load_background():
    with zipfile.ZipFile(BG_PATH, "r") as z:
        name = next(n for n in z.namelist() if n.lower().endswith((".jpeg", ".jpg", ".png")))
        return z.read(name)

# ─── Helpers ──────────────────────────────────────────────────────────────────

def extract_image_from_drawing(file_bytes: bytes) -> bytes:
    errors = []

    if zipfile.is_zipfile(io.BytesIO(file_bytes)):
        with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as z:
            name = next(
                (n for n in z.namelist() if n.lower().endswith((".jpeg", ".jpg", ".png"))),
                None,
            )
            if name:
                return z.read(name)
            errors.append(f"ZIP ok but no image. Contents: {z.namelist()}")
    else:
        errors.append("Not a ZIP file.")

    try:
        import fitz
        doc  = fitz.open(stream=file_bytes, filetype="pdf")
        page = doc[0]
        mat  = fitz.Matrix(2.5, 2.5)
        pix  = page.get_pixmap(matrix=mat, alpha=False)
        return pix.tobytes("png")
    except Exception as e:
        errors.append(f"PyMuPDF failed: {e}")

    try:
        import pdf2image
        imgs = pdf2image.convert_from_bytes(file_bytes, dpi=150, first_page=1, last_page=1)
        if imgs:
            buf = io.BytesIO()
            imgs[0].save(buf, format="PNG")
            return buf.getvalue()
        errors.append("pdf2image returned no pages.")
    except Exception as e:
        errors.append(f"pdf2image failed: {e}")

    try:
        import pypdf
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        for pg in reader.pages:
            imgs = list(pg.images)
            if imgs:
                return imgs[0].data
        errors.append("pypdf: no embedded images.")
    except Exception as e:
        errors.append(f"pypdf failed: {e}")

    raise ValueError("Could not extract image from PDF.\n" + "\n".join(errors))


def read_excel_rows(file_bytes: bytes) -> list:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active
    rows = []
    for row in ws.iter_rows(values_only=True):
        if any(v is not None for v in row):
            rows.append([str(v) if v is not None else "" for v in row[:5]])
    
    # If more than 30 rows, keep only first 25 (header + 24 data rows)
    if len(rows) > 30:
        rows = rows[:25]
    
    return rows


def derive_doc_name(filename: str) -> str:
    base = os.path.splitext(filename)[0]
    return base.replace("__", "  ").replace("_", " ")


# ─── PDF drawing helpers ──────────────────────────────────────────────────────

def draw_background(c, bg_bytes):
    c.drawImage(ImageReader(io.BytesIO(bg_bytes)), 0, 0,
                width=PAGE_W, height=PAGE_H, preserveAspectRatio=False)

def draw_footer_centered(c, doc_name, part_no=""):
    cx = (LEFT_MM + RIGHT_MM) / 2 * mm   # true centre
    c.setFont("Helvetica-Bold", 16)
    c.setFillColorRGB(0, 0, 0)
    c.drawCentredString(cx, (FCENTER_MM + 6) * mm, doc_name)
    if part_no:
        c.setFont("Helvetica", 12)
        c.drawCentredString(cx, (FCENTER_MM - 2) * mm, part_no)   # +5mm gap between lines

def draw_badge(c, group_no: str, ref_no: str):
    rx = (RIGHT_MM - 2) * mm
    c.setFont("Helvetica-Bold", 16)
    c.setFillColorRGB(0, 0, 0)
    c.drawRightString(rx, (FCENTER_MM + 6) * mm, group_no)   # same Y as doc_name
    if ref_no:
        c.setFont("Helvetica", 13)
        c.drawRightString(rx, (FCENTER_MM - 2) * mm, ref_no)  # same Y as part_no, +5mm gap

def draw_page_number(c, page_num: int):
    """Page number centred at the very bottom of the page."""
    c.setFont("Helvetica", 13)
    c.setFillColorRGB(0, 0, 0)
    cx = (LEFT_MM + RIGHT_MM) / 2 * mm
    c.drawCentredString(cx, 6 * mm, str(page_num))


def draw_excel_table(c, rows):
    aw = AVAIL_W_MM * mm
    ah = AVAIL_H_MM * mm
    cw = [aw * 0.04, aw * 0.15, aw * 0.05, aw * 0.20, aw * 0.56]

    cs = ParagraphStyle("c", fontName="Times-Roman", fontSize=13, leading=16)
    hs = ParagraphStyle("h", fontName="Times-Bold",  fontSize=15, leading=17)

    data = [
        [Paragraph(str(cell), hs if i == 0 else cs) for cell in row]
        for i, row in enumerate(rows)
    ]

    t = Table(data, colWidths=cw)
    t.setStyle(TableStyle([
        ("TEXTCOLOR",     (0, 0), (-1, -1), colors.black),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 3.6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.3),
        ("LEFTPADDING",   (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 3),
    ]))

    _, th = t.wrapOn(c, aw, ah)
    t.drawOn(c, LEFT_MM * mm, (CTOP_MM * mm - th) - 0.8 * mm)


def draw_content_image(c, content_bytes):
    aw = AVAIL_W_MM * mm
    ah = AVAIL_H_MM * mm
    scale_factor = 0.89
    pil = Image.open(io.BytesIO(content_bytes))
    asp = pil.width / pil.height
    dw, dh = (aw, aw / asp) if aw / asp <= ah else (ah * asp, ah)
    dw *= scale_factor
    dh *= scale_factor
    ix = LEFT_MM * mm + (aw - dw) / 2
    iy = CBOT_MM * mm + (ah - dh) / 2
    c.drawImage(ImageReader(io.BytesIO(content_bytes)), ix, iy,
                width=dw, height=dh, preserveAspectRatio=True)


# ─── Main PDF generator ───────────────────────────────────────────────────────

def generate_pdf(slots: list, start_page: int) -> bytes:
    """
    slots: list of dicts — xlsx_bytes, pdf_bytes, group_no, ref_no
    start_page: page number for first page
    Each slot produces 2 pages: drawing page then table page.
    """
    bg_bytes = load_background()
    buf = io.BytesIO()
    c   = canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))

    current_page = start_page

    for slot in slots:
        group_no = slot["group_no"].strip()
        ref_no   = slot["ref_no"].strip()
        part_desc, part_no = "", ""

        if slot["xlsx_bytes"]:
            rows = read_excel_rows(slot["xlsx_bytes"])
            part_desc = rows[1][4] if len(rows) > 1 and len(rows[1]) > 4 else ""
            part_no   = rows[1][1] if len(rows) > 1 and len(rows[1]) > 1 else ""

        # Drawing page — only if PDF uploaded
        if slot["pdf_bytes"]:
            img_bytes = extract_image_from_drawing(slot["pdf_bytes"])
            draw_background(c, bg_bytes)
            draw_content_image(c, img_bytes)
            draw_footer_centered(c, part_desc, part_no)
            draw_badge(c, group_no, ref_no)
            draw_page_number(c, current_page)
            c.showPage()
            current_page += 1

        # Table page — only if Excel uploaded
        if slot["xlsx_bytes"]:
            draw_background(c, bg_bytes)
            draw_excel_table(c, rows)
            draw_footer_centered(c, part_desc, part_no)
            draw_badge(c, group_no, ref_no)
            draw_page_number(c, current_page)
            c.showPage()
            current_page += 1

    c.save()
    return buf.getvalue()


# ─── UI ───────────────────────────────────────────────────────────────────────

st.markdown("## 📄 Aquarius Document Generator")
st.caption("Add slots below — each slot is one drawing PDF + one Excel parts list. "
           "All slots compile into a single paginated PDF.")
st.divider()

# ── Session state ─────────────────────────────────────────────────────────────
if "num_slots" not in st.session_state:
    st.session_state.num_slots = 1

# ── Starting page number ──────────────────────────────────────────────────────
col_pg, _ = st.columns([1, 3])
with col_pg:
    start_page = st.number_input(
        "📄 Starting page number",
        min_value=1, value=1, step=1,
        help="The first drawing page will carry this number; subsequent pages increment automatically.",
    )

st.markdown("---")

# ── Slot cards ────────────────────────────────────────────────────────────────
slot_data  = []   # accumulate ready slots in order
page_cursor = int(start_page)

for i in range(st.session_state.num_slots):
    with st.container(border=True):
        hcol, _ = st.columns([6, 1])
        with hcol:
            st.markdown(
                f"**Slot {i + 1}** &nbsp;·&nbsp; "
                f"<span style='color:gray;font-size:13px'>pages {page_cursor} & {page_cursor + 1}</span>",
                unsafe_allow_html=True,
            )

        c1, c2, c3, c4 = st.columns([2.5, 2.5, 1, 1.2])

        with c1:
            pdf_f = st.file_uploader(
                "🖼️ Engineering Drawing (PDF)",
                type=["pdf"],
                key=f"pdf_{i}",
            )
        with c2:
            xlsx_f = st.file_uploader(
                "📊 Parts List (Excel)",
                type=["xlsx", "xls"],
                key=f"xlsx_{i}",
            )
        with c3:
            group_no = st.text_input(
                "Group No.",
                key=f"group_{i}",
                placeholder="e.g. 3.1.0",
            )
        with c4:
            ref_no = st.text_input(
                "Reference No.",
                key=f"ref_{i}",
                placeholder="e.g. 3100-020415",
            )

        has_pdf  = pdf_f  is not None
        has_xlsx = xlsx_f is not None

        if has_pdf or has_xlsx:
            st.success(f"✅ Ready — pages {page_cursor} {'(drawing)' if has_pdf else ''} {'& ' + str(page_cursor+1) + ' (table)' if has_pdf and has_xlsx else '(table)' if has_xlsx else ''}")
            slot_data.append({
                "xlsx_bytes": xlsx_f.getvalue() if has_xlsx else None,
                "pdf_bytes":  pdf_f.getvalue()  if has_pdf  else None,
                "group_no":   group_no,
                "ref_no":     ref_no,
            })
        else:
            st.warning("⚠️ Upload at least a Drawing PDF or Excel file.")
            
    pages_this_slot = (1 if pdf_f else 0) + (1 if xlsx_f else 0)
    page_cursor += max(pages_this_slot, 1)  # at least 1 even if empty, so numbering stays predictable

st.markdown("")

# ── Add / Remove buttons ──────────────────────────────────────────────────────
ba, br, *_ = st.columns([1, 1, 4])
with ba:
    if st.button("➕  Add slot", use_container_width=True):
        st.session_state.num_slots += 1
        st.rerun()
with br:
    if st.session_state.num_slots > 1:
        if st.button("➖  Remove last", use_container_width=True):
            st.session_state.num_slots -= 1
            st.rerun()

st.divider()

# ── Generate ──────────────────────────────────────────────────────────────────
total_slots = st.session_state.num_slots
ready_slots = len(slot_data)

if ready_slots < total_slots:
    st.info(f"{ready_slots} of {total_slots} slots ready. "
            f"Fill all slots before generating, or remove empty ones.")

if st.button(
    f"⚙️  Generate PDF  ({ready_slots * 2} pages)",
    type="primary",
    use_container_width=True,
    disabled=(ready_slots == 0),
):
    with st.spinner("Generating…"):
        try:
            pdf_bytes = generate_pdf(slot_data, start_page=int(start_page))
            st.success(
                f"✅ Done — {ready_slots * 2} pages, "
                f"numbered {start_page}–{int(start_page) + ready_slots * 2 - 1}"
            )
            st.download_button(
                label            = "⬇️  Download PDF",
                data             = pdf_bytes,
                file_name        = "aquarius_document.pdf",
                mime             = "application/pdf",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Generation failed: {e}")
            st.exception(e)

st.divider()
st.caption("Aquarius Industries · Internal tool · All slots compile into one PDF.")
