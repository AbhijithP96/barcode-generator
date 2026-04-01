import io
import base64
import tempfile
import openpyxl
import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import code128

# Page config
st.set_page_config(
    page_title="Barcode Generator",
    page_icon="▌▌▌ ▌▌ ▌",
    layout="wide",
)

# Styles
st.markdown("""
<style>
    /* ── Global background + text ── */
    html, body, [data-testid="stAppViewContainer"],
    [data-testid="stMain"], .main, section.main {
        background-color: #000000 ;
        color: #ffffff ;
    }
    [data-testid="stSidebar"] { background-color: #0d0d0d ; }
    .block-container { padding-top: 2rem; background: #000000 ; }

    /* ── All text ── */
    h1, h2, h3, h4, h5, h6, p, span, label, div,
    .stMarkdown, .stCaption, [data-testid="stText"] {
        color: #ffffff ;
    }
    h1 { font-size: 1.8rem ; }

    /* ── Inputs & widgets ── */
    .stSelectbox > div > div,
    .stMultiSelect > div > div {
        background-color: #1a1a1a ;
        color: #ffffff ;
        border-color: #444 ;
    }
    [data-baseweb="select"] * { color: #ffffff ; background: #1a1a1a ; }
    [data-baseweb="menu"]       { background: #1a1a1a ; }
    [data-baseweb="option"]:hover { background: #333 ; }
    [data-testid="stFileUploader"] > div {
        background: #111 ;
        border-color: #444 ;
        color: #ffffff ;
    }
    hr { border-color: #333; }

    /* ── Buttons ── */
    .stButton > button {
        width: 100%;
        background: #ffffff ;
        color: #000000 ;
        border: none ;
        padding: 0.6rem 1.2rem;
        border-radius: 8px;
        font-size: 1rem;
        font-weight: 700;
        margin-top: 0.5rem;
    }
    .stButton > button *          { color: #000000 ; }
    .stButton > button p          { color: #000000 ; }
    .stButton > button:hover      { background: #e0e0e0 ; color: #000000 ; }
    .stButton > button:hover *    { color: #000000 ; }
    .stButton > button:disabled   { background: #333 ; color: #777 ; }
    .stButton > button:disabled * { color: #777 ; }

    [data-testid="stDownloadButton"] > button   { background: #ffffff ; color: #000000 ; font-weight: 700 ; border-radius: 8px; border: none ; width: 100%; }
    [data-testid="stDownloadButton"] > button * { color: #000000 ; }
    [data-testid="stDownloadButton"] > button p { color: #000000 ; }
    [data-testid="stDownloadButton"] > button:hover { background: #e0e0e0 ; }

    /* ── Info box ── */
    .info-box {
        background: #111;
        border-left: 4px solid #ffffff;
        padding: 0.8rem 1rem;
        border-radius: 4px;
        margin: 0.5rem 0 1rem 0;
        font-size: 0.9rem;
        color: #ffffff;
    }

    /* ── Stat cards ── */
    .stat-card {
        background: #111;
        border: 1px solid #333;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
    }
    .stat-num   { font-size: 2rem; font-weight: 700; color: #ffffff; }
    .stat-label { font-size: 0.75rem; color: #aaaaaa; text-transform: uppercase; }

    /* ── Misc ── */
    [data-testid="stSpinner"] * { color: #ffffff ; }
    [data-testid="stAlert"]     { background: #1a1a1a ; color: #fff ; }
</style>
""", unsafe_allow_html=True)

# Layout constants
FONT_NAME = "Helvetica"

# Core functions

def read_excel(file, has_header: bool) -> dict:
    """Return {sheet_name: DataFrame} for all sheets.
    If no header, columns are named Column 1, Column 2, and all rows are data.
    """
    if has_header:
        dfs = pd.read_excel(file, sheet_name=None, header=0, dtype=str)
    else:
        dfs = pd.read_excel(file, sheet_name=None, header=None, dtype=str)
        dfs = {
            name: df.rename(columns={i: f"Column {i+1}" for i in df.columns})
            for name, df in dfs.items()
        }
    return dfs


def generate_pdf(
    sheets_data: dict,        # {sheet_name: [serial, ...]}
    cols: int,
    cell_h_mm: float,
    bar_h_mm: float,
    bar_color_hex: str,
    serial_size: int,
    show_sheet_label: bool,
) -> bytes:
    PAGE_W, PAGE_H = A4
    MARGIN   = 10 * mm
    CELL_H   = cell_h_mm * mm
    BAR_H    = bar_h_mm  * mm
    CELL_W   = (PAGE_W - 2 * MARGIN) / cols
    BAR_W    = CELL_W - 6 * mm

    all_items = [
        (sheet, serial)
        for sheet, serials in sheets_data.items()
        for serial in serials
    ]

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle("Serial Number Barcodes")

    usable_h      = PAGE_H - 2 * MARGIN - 12 * mm
    rows_per_page = max(1, int(usable_h // CELL_H))
    per_page      = cols * rows_per_page
    total_pages   = (len(all_items) + per_page - 1) // per_page

    bar_fill = colors.HexColor(bar_color_hex)

    for page_idx in range(total_pages):
        chunk = all_items[page_idx * per_page: (page_idx + 1) * per_page]

        # Page header
        c.setFont(FONT_NAME + "-Bold", 10)
        c.setFillColor(colors.HexColor("#222222"))
        c.drawString(MARGIN, PAGE_H - MARGIN + 2, "Serial Number Barcodes")
        c.setFont(FONT_NAME, 8)
        c.setFillColor(colors.HexColor("#666666"))
        c.drawRightString(PAGE_W - MARGIN, PAGE_H - MARGIN + 2,
                          f"Page {page_idx + 1} of {total_pages}")
        c.setStrokeColor(colors.HexColor("#333333"))
        c.setLineWidth(0.8)
        c.line(MARGIN, PAGE_H - MARGIN - 2, PAGE_W - MARGIN, PAGE_H - MARGIN - 2)

        for i, (sheet_name, serial) in enumerate(chunk):
            col_i = i % cols
            row_i = i // cols
            x = MARGIN + col_i * CELL_W
            y = PAGE_H - MARGIN - 10 * mm - (row_i + 1) * CELL_H

            # Cell border
            c.setStrokeColor(colors.HexColor("#CCCCCC"))
            c.setLineWidth(0.5)
            c.rect(x + 2, y + 2, CELL_W - 4, CELL_H - 4)

            # Sheet label
            if show_sheet_label:
                c.setFont(FONT_NAME + "-Oblique", 5.5)
                c.setFillColor(colors.HexColor("#999999"))
                c.drawString(x + 5, y + CELL_H - 8, sheet_name)

            # Positions (anchored bottom-up)
            serial_y = y + 3
            bar_y    = serial_y + 4 * mm
            bar_x    = x + (CELL_W - BAR_W) / 2

            # Barcode
            try:
                bc = code128.Code128(
                    serial,
                    barWidth=0.7,
                    barHeight=BAR_H,
                    humanReadable=False,
                )
                bc.barFillColor = bar_fill
                bc.drawOn(c, bar_x, bar_y)
            except Exception:
                c.setFont(FONT_NAME, 6)
                c.setFillColor(colors.red)
                c.drawCentredString(x + CELL_W / 2, bar_y + BAR_H / 2, "[error]")

            # Serial label
            c.setFont(FONT_NAME + "-Bold", serial_size)
            c.setFillColor(colors.black)
            c.drawCentredString(x + CELL_W / 2, serial_y, serial)

        c.showPage()

    c.save()
    buf.seek(0)
    return buf.read()


def pdf_preview_html(pdf_bytes: bytes) -> str:
    b64 = base64.b64encode(pdf_bytes).decode()
    return f"""
    <iframe
        src="data:application/pdf;base64,{b64}"
        width="100%" height="700px"
        style="border:none; border-radius:8px; box-shadow:0 2px 12px rgba(0,0,0,0.15);"
    ></iframe>
    """


# UI

st.title("▌▌▌ ▌▌ ▌  Barcode PDF Generator")
st.caption("Upload an Excel file → pick sheets & column → generate A4 barcode PDF")

left, right = st.columns([1, 2], gap="large")

with left:
    st.subheader("⚙️ Settings")

    uploaded = st.file_uploader(
        "Upload Excel file (.xlsx)",
        type=["xlsx"],
        help="File can have multiple sheets",
    )

    has_header = st.checkbox(
        "First row is a header",
        value=True,
        help="Uncheck if your data starts from row 1 with no column names",
    )

    sheets_data_raw = {}
    all_sheet_names = []
    all_columns     = []

    if uploaded:
        try:
            raw = read_excel(uploaded, has_header)
            all_sheet_names = list(raw.keys())
            # Collect all columns across sheets
            all_columns = sorted({
                col for df in raw.values() for col in df.columns
            })
            st.markdown(
                f'<div class="info-box">📄 <b>{len(all_sheet_names)} sheet(s)</b> found: '
                f'{", ".join(all_sheet_names)}</div>',
                unsafe_allow_html=True,
            )
        except Exception as e:
            st.error(f"Could not read file: {e}")

    selected_sheets = st.multiselect(
        "Sheets to include",
        options=all_sheet_names,
        default=all_sheet_names,
        disabled=not uploaded,
    )

    selected_col = st.selectbox(
        "Serial number column",
        options=all_columns if all_columns else ["—"],
        disabled=not uploaded,
        help="Column that contains the serial numbers",
    )

    st.divider()
    st.markdown("**Layout**")

    cols_count = st.slider("Columns per row", 2, 6, 4)
    cell_h     = st.slider("Cell height (mm)", 15, 60, 25)
    bar_h      = st.slider("Barcode height (mm)", 6, 30, 12)

    st.markdown("**Style**")
    bar_color      = st.color_picker("Barcode colour", "#0a0a0a")
    serial_size    = st.slider("Serial number font size", 5, 14, 7)
    show_label     = st.checkbox("Show sheet name label", value=True)

    st.divider()
    generate_btn = st.button("🖨️ Generate PDF", disabled=not (uploaded and selected_sheets and selected_col != "—"))


# Generation

if generate_btn and uploaded and selected_sheets:
    raw = read_excel(uploaded, has_header)

    for sheet in selected_sheets:
        df = raw.get(sheet)
        if df is not None and selected_col in df.columns:
            serials = df[selected_col].dropna().astype(str).str.strip().tolist()
            serials = [s for s in serials if s]
            sheets_data_raw[sheet] = serials

    total_barcodes = sum(len(v) for v in sheets_data_raw.values())

    if total_barcodes == 0:
        with right:
            st.warning("No serial numbers found in selected sheets/column.")
    else:
        with st.spinner("Generating PDF…"):
            pdf_bytes = generate_pdf(
                sheets_data_raw,
                cols=cols_count,
                cell_h_mm=cell_h,
                bar_h_mm=bar_h,
                bar_color_hex=bar_color,
                serial_size=serial_size,
                show_sheet_label=show_label,
            )

        # Stats
        per_page   = cols_count * max(1, int(((841.89 - 20*mm - 12*mm) // (cell_h*mm))))
        est_pages  = max(1, (total_barcodes + per_page - 1) // per_page)

        with right:
            st.subheader("📊 Summary")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f'<div class="stat-card"><div class="stat-num">{total_barcodes}</div><div class="stat-label">Barcodes</div></div>', unsafe_allow_html=True)
            with c2:
                st.markdown(f'<div class="stat-card"><div class="stat-num">{est_pages}</div><div class="stat-label">Pages</div></div>', unsafe_allow_html=True)
            with c3:
                sz = f"{len(pdf_bytes)/1024:.1f} KB"
                st.markdown(f'<div class="stat-card"><div class="stat-num">{sz}</div><div class="stat-label">File size</div></div>', unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # Download + Print buttons
            dl1, dl2 = st.columns(2)
            with dl1:
                st.download_button(
                    "\u2193 Download PDF",
                    data=pdf_bytes,
                    file_name="barcodes.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            with dl2:
                b64 = base64.b64encode(pdf_bytes).decode()
                st.markdown(f"""
                <a href="data:application/pdf;base64,{b64}" target="_blank">
                    <button style="
                        width:100%; background:#ffffff; color:#000000;
                        border:none; padding:0.45rem 1rem; border-radius:8px;
                        font-size:0.95rem; font-weight:700; cursor:pointer;
                    ">🖨️ Open to Print</button>
                </a>
                """, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.subheader("👁️ Preview")
            st.markdown(pdf_preview_html(pdf_bytes), unsafe_allow_html=True)

elif not uploaded:
    with right:
        st.markdown("""
        <div style="
            display:flex; flex-direction:column; align-items:center;
            justify-content:center; height:400px;
            background:#111; border-radius:12px;
            border: 2px dashed #444; color:#666;
        ">
            <div style="font-size:3rem">▌▌▌ ▌▌ ▌</div>
            <div style="font-size:1.1rem; margin-top:1rem; color:#888;">Upload an Excel file to get started</div>
        </div>
        """, unsafe_allow_html=True)