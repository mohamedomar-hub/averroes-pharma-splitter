# Averroes Pharma File Splitter & Dashboard (Unified Navbar + Gold Progress + Timer)
# Full Streamlit app code (replace your existing app file with this)
# Note: Requires streamlit, pandas, openpyxl, pillow, streamlit_lottie, plotly, reportlab, python-pptx, opencv-python (optional)
# Author: Assistant (integrated per user request)

import streamlit as st
import pandas as pd
import numpy as np
import re
import os
import time
from io import BytesIO
from zipfile import ZipFile
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from PIL import Image
import requests
from streamlit_lottie import st_lottie
import matplotlib.pyplot as plt
import plotly.express as px
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image as RLImage, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ----------------------- Utilities & Config -----------------------
st.set_page_config(
    page_title="Averroes Pharma File Splitter & Dashboard",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Lottie loader cached
@st.cache_data(show_spinner=False)
def load_lottie_url(url: str):
    try:
        r = requests.get(url, timeout=8)
        if r.status_code == 200:
            return r.json()
    except Exception:
        return None
    return None

LOTTIE_SPLIT = load_lottie_url("https://assets9.lottiefiles.com/packages/lf20_wx9z5gxb.json")
LOTTIE_MERGE = load_lottie_url("https://assets10.lottiefiles.com/packages/lf20_cg3rwjul.json")
LOTTIE_IMAGE = load_lottie_url("https://assets2.lottiefiles.com/private_files/lf30_cgfdhxgx.json")
LOTTIE_DASH  = load_lottie_url("https://assets8.lottiefiles.com/packages/lf20_tno6cg2w.json")
LOTTIE_PDF   = load_lottie_url("https://assets1.lottiefiles.com/packages/lf20_zyu0ct3i.json")
LOTTIE_PROCESS_SMALL = load_lottie_url("https://assets2.lottiefiles.com/packages/lf20_jtbfg2nb.json")  # optional small processing

# helper safe name
def _safe_name(s):
    return re.sub(r'[^A-Za-z0-9_-]+', '_', str(s))

# find column helper
def _find_col(df, aliases):
    lowered = {c.lower(): c for c in df.columns}
    for a in aliases:
        if a.lower() in lowered:
            return lowered[a.lower()]
    for c in df.columns:
        name = c.strip().lower()
        for a in aliases:
            if a.lower() in name:
                return c
    return None

# gold progress renderer used across split & merge
def render_gold_progress(placeholder, percentage, message, elapsed_seconds):
    html = f"""
    <style>
    .gold-container {{
        background-color: #00264d;
        border: 1px solid #FFD700;
        border-radius: 12px;
        padding: 12px;
        margin-top: 12px;
        text-align: center;
        box-shadow: 0 6px 18px rgba(0,0,0,0.35);
    }}
    .gold-progress {{
        width: 100%;
        background-color: #001f3f;
        border-radius: 20px;
        height: 30px;
        margin-top: 10px;
        overflow: hidden;
        position: relative;
    }}
    .gold-fill {{
        height: 100%;
        width: {percentage}%;
        border-radius: 20px;
        background: linear-gradient(90deg, #FFD700, #FFC107, #FFB300);
        background-size: 200% 100%;
        animation: shine 2s linear infinite;
        line-height: 30px;
        font-weight: bold;
        color: black;
        transition: width 0.45s ease;
        text-align: center;
    }}
    @keyframes shine {{
        0% {{ background-position: 0% 0%; }}
        100% {{ background-position: 200% 0%; }}
    }}
    .gold-meta {{
        color: #FFD700;
        font-weight: 700;
        margin-bottom: 6px;
    }}
    .gold-time {{
        color: #ffffff;
        opacity: 0.85;
        font-size: 13px;
        margin-top: 6px;
    }}
    </style>
    <div class="gold-container">
        <div class="gold-meta">🔄 {message}</div>
        <div class="gold-progress">
            <div class="gold-fill">{percentage}%</div>
        </div>
        <div class="gold-time">Elapsed: {elapsed_seconds:.1f} sec</div>
    </div>
    """
    placeholder.markdown(html, unsafe_allow_html=True)

# confetti on success (embedded)
def show_confetti():
    import streamlit.components.v1 as components
    confetti_js = """
    <canvas id="confetti-canvas" style="position:fixed;pointer-events:none;top:0;left:0;width:100%;height:100%;z-index:999999;"></canvas>
    <script src="https://cdn.jsdelivr.net/npm/canvas-confetti@1.5.1/dist/confetti.browser.min.js"></script>
    <script>
      var myCanvas = document.getElementById('confetti-canvas');
      var myConfetti = confetti.create(myCanvas, { resize: true });
      myConfetti({ particleCount: 120, spread: 160, origin: { y: 0.6 } });
      setTimeout(()=>{ myCanvas.remove(); }, 3500);
    </script>
    """
    components.html(confetti_js, height=0, width=0)

# copy cell style helper (for preserving style in merge)
def copy_cell_style(src_cell, dst_cell):
    try:
        if src_cell.font:
            dst_cell.font = Font(name=src_cell.font.name, size=src_cell.font.size,
                                 bold=src_cell.font.bold, italic=src_cell.font.italic, color=src_cell.font.color)
    except Exception:
        pass
    try:
        if src_cell.fill and src_cell.fill.fill_type:
            dst_cell.fill = PatternFill(fill_type=src_cell.fill.fill_type,
                                        start_color=src_cell.fill.start_color,
                                        end_color=src_cell.fill.end_color)
    except Exception:
        pass
    try:
        if src_cell.border:
            dst_cell.border = Border(left=src_cell.border.left, right=src_cell.border.right,
                                     top=src_cell.border.top, bottom=src_cell.border.bottom)
    except Exception:
        pass
    try:
        if src_cell.alignment:
            dst_cell.alignment = Alignment(horizontal=src_cell.alignment.horizontal,
                                           vertical=src_cell.alignment.vertical,
                                           wrap_text=src_cell.alignment.wrap_text)
    except Exception:
        pass
    try:
        dst_cell.number_format = src_cell.number_format
    except Exception:
        pass

# ----------------------- CSS & Navbar (top) -----------------------
# hide default elements & add custom CSS
hide_default = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
"""
st.markdown(hide_default, unsafe_allow_html=True)

custom_css = """
    <style>
    .stApp {
        background-color: #001f3f;
        color: white;
        font-family: 'Cairo', sans-serif;
    }
    .top-nav {
        display: flex;
        justify-content: flex-end;
        gap: 16px;
        padding: 10px 30px;
        background-color: #001a33;
        border-bottom: 1px solid #FFD700;
        font-size: 16px;
        color: white;
        align-items: center;
    }
    .nav-left {
        margin-right: auto;
        display:flex;
        align-items:center;
        gap:16px;
    }
    .nav-link {
        color: #FFD700;
        text-decoration: none;
        font-weight: 700;
        font-size: 16px;
        margin: 0 12px;
        transition: color 0.2s ease-in-out;
    }
    .nav-link:hover {
        color: #FFE97F
    }
    .nav-link.active {
        color: #FFFFFF
    }
    .logo-small {
        width: 120px;
        height: auto;
        background: white;
        border-radius: 8px;
        padding: 6px;
        display:inline-block;
    }
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# Navbar HTML (anchors that scroll to sections)
navbar_html = """
<div class="top-nav">
    <div class="nav-left">
        <div style="padding-left:10px;">
            <img src="logo.png" style="width:110px; height:auto; border-radius:6px;" onerror="">
        </div>
    </div>
    <a class="nav-link" href="#home">Home</a>
    <a class="nav-link" href="#split">Split & Merge</a>
    <a class="nav-link" href="#imagetopdf">Image to PDF</a>
    <a class="nav-link" href="#dashboard">Auto Dashboard</a>
    <a class="nav-link" href="#contact">Contact</a>
</div>
<script>
document.querySelectorAll('.nav-link').forEach(a=>{
    a.addEventListener('click', function(e){
        e.preventDefault();
        var href = this.getAttribute('href');
        var el = document.querySelector(href);
        if(el){
            el.scrollIntoView({behavior: 'smooth'});
        }
    });
});
</script>
"""
st.markdown(navbar_html, unsafe_allow_html=True)

# ----------------------- Header / Hero -----------------------
st.markdown('<a name="home"></a>', unsafe_allow_html=True)
st.markdown("""
    <div style='text-align:center; margin-top:20px;'>
        <div style='color:#FFD700; font-weight:700;'>By <strong>Mohamed Abd ELGhany</strong> – 01554694554 (WhatsApp)</div>
        <h1 style='color:#FFD700; font-size:38px; margin:8px 0;'>💊 Averroes Pharma File Splitter & Dashboard</h1>
        <div style='color:white; font-size:16px;'>✂ Split, Merge, Image-to-PDF & Auto Dashboard Generator</div>
    </div>
    <hr style='border:1px solid #123; margin-top:18px; opacity:0.4;' />
""", unsafe_allow_html=True)

# ----------------------- SECTION: Split & Merge (anchor) -----------------------
st.markdown('<a name="split"></a>', unsafe_allow_html=True)
st.markdown("<h2 style='color:#FFD700;'>✂ Split Excel/CSV File</h2>", unsafe_allow_html=True)

# Upload for split
uploaded_file = st.file_uploader(
    "📂 Upload Excel or CSV File (Splitter/Merge)",
    type=["xlsx", "csv"],
    accept_multiple_files=False,
    key=f"split_uploader_{st.session_state.get('clear_counter',0)}"
)

def clean_name(name):
    name = str(name).strip()
    invalid_chars = r'[\\/*?:\[\]|<>"]'
    cleaned = re.sub(invalid_chars, '_', name)
    return cleaned[:30] if cleaned else "Sheet"

if uploaded_file:
    # display
    st.markdown(f"<div style='color:#FFD700; font-weight:bold;'>Uploaded: {uploaded_file.name}</div>", unsafe_allow_html=True)
    if uploaded_file.name.lower().endswith('csv'):
        df = pd.read_csv(uploaded_file)
        sheet_names = ["Sheet1"]
        selected_sheet = "Sheet1"
        st.success("✅ CSV file uploaded successfully.")
    else:
        input_bytes = uploaded_file.getvalue()
        original_wb = load_workbook(filename=BytesIO(input_bytes), data_only=False)
        sheet_names = original_wb.sheetnames
        st.success(f"✅ Excel file uploaded successfully. Number of sheets: {len(sheet_names)}")
        selected_sheet = st.selectbox("Select Sheet (for Split)", sheet_names)
        df = pd.read_excel(BytesIO(input_bytes), sheet_name=selected_sheet)

    st.markdown(f"### 📊 Data View – {selected_sheet}")
    st.dataframe(df, use_container_width=True)

    st.markdown("### ✂ Select Column to Split")
    col_to_split = st.selectbox("Split by Column", df.columns, help="Select the column to split by, such as 'Brick' or 'Area Manager'")

    st.markdown("### ⚙️ Split Options")
    split_option = st.radio("Choose split method:", ["Split by Column Values", "Split Each Sheet into Separate File"], index=0)

    # Start split button
    if st.button("🚀 Start Split"):
        with st.spinner("Splitting process in progress..."):
            if LOTTIE_SPLIT:
                st_lottie(LOTTIE_SPLIT, height=140, key="lottie_split_main")

            progress_placeholder = st.empty()
            start_time = time.time()

            # CSV case
            if uploaded_file.name.lower().endswith('csv'):
                unique_values = df[col_to_split].dropna().unique()
                zip_buffer = BytesIO()
                total = len(unique_values) if len(unique_values) > 0 else 1
                with ZipFile(zip_buffer, "w") as zip_file:
                    for i, value in enumerate(unique_values):
                        filtered_df = df[df[col_to_split] == value]
                        csv_buffer = BytesIO()
                        filtered_df.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                        csv_buffer.seek(0)
                        file_name = f"{clean_name(value)}.csv"
                        zip_file.writestr(file_name, csv_buffer.read())
                        elapsed = time.time() - start_time
                        pct = int(((i + 1) / total) * 100)
                        render_gold_progress(progress_placeholder, pct, f"Splitting {i+1}/{total}: {value}", elapsed)

                elapsed = time.time() - start_time
                render_gold_progress(progress_placeholder, 100, "✅ Splitting Completed Successfully!", elapsed)
                time.sleep(0.25)
                progress_placeholder.empty()

                zip_buffer.seek(0)
                st.success("🎉 Splitting completed successfully!")
                st.download_button(
                    label="📥 Download Split Files (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"Split_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip",
                    mime="application/zip"
                )
                show_confetti()

            else:
                # Excel split
                if split_option == "Split by Column Values":
                    ws = original_wb[selected_sheet]
                    col_idx = df.columns.get_loc(col_to_split) + 1
                    unique_values = df[col_to_split].dropna().unique()
                    zip_buffer = BytesIO()
                    total = len(unique_values) if len(unique_values) > 0 else 1

                    with ZipFile(zip_buffer, "w") as zip_file:
                        for i, value in enumerate(unique_values):
                            new_wb = Workbook()
                            default_ws = new_wb.active
                            new_wb.remove(default_ws)
                            new_ws = new_wb.create_sheet(title=clean_name(value))

                            # copy header row with style
                            for cell in ws[1]:
                                dst_cell = new_ws.cell(1, cell.column, cell.value)
                                copy_cell_style(cell, dst_cell)

                            # copy rows matching value
                            row_idx = 2
                            for row in ws.iter_rows(min_row=2):
                                cell_in_col = row[col_idx - 1]
                                if cell_in_col.value == value:
                                    for src_cell in row:
                                        dst_cell = new_ws.cell(row_idx, src_cell.column, src_cell.value)
                                        copy_cell_style(src_cell, dst_cell)
                                    row_idx += 1

                            # copy column widths
                            try:
                                for col_letter in ws.column_dimensions:
                                    if ws.column_dimensions[col_letter].width:
                                        new_ws.column_dimensions[col_letter].width = ws.column_dimensions[col_letter].width
                            except Exception:
                                pass

                            # save sheet to buffer and store in zip
                            file_buffer = BytesIO()
                            new_wb.save(file_buffer)
                            file_buffer.seek(0)
                            file_name = f"{clean_name(value)}.xlsx"
                            zip_file.writestr(file_name, file_buffer.read())

                            elapsed = time.time() - start_time
                            pct = int(((i + 1) / total) * 100)
                            render_gold_progress(progress_placeholder, pct, f"Splitting {i+1}/{total}: {value}", elapsed)

                    elapsed = time.time() - start_time
                    render_gold_progress(progress_placeholder, 100, "✅ Splitting Completed Successfully!", elapsed)
                    time.sleep(0.25)
                    progress_placeholder.empty()

                    zip_buffer.seek(0)
                    st.success("🎉 Splitting completed successfully!")
                    st.download_button(
                        label="📥 Download Split Files (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name=f"Split_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip",
                        mime="application/zip"
                    )
                    show_confetti()

                elif split_option == "Split Each Sheet into Separate File":
                    zip_buffer = BytesIO()
                    sheets = original_wb.sheetnames
                    total = len(sheets) if len(sheets) > 0 else 1
                    with ZipFile(zip_buffer, "w") as zip_file:
                        for i, sheet_name in enumerate(sheets):
                            src_ws = original_wb[sheet_name]
                            new_wb = Workbook()
                            default_ws = new_wb.active
                            new_wb.remove(default_ws)
                            new_ws = new_wb.create_sheet(title=clean_name(sheet_name))

                            for row in src_ws.iter_rows():
                                for src_cell in row:
                                    dst_cell = new_ws.cell(src_cell.row, src_cell.column, src_cell.value)
                                    copy_cell_style(src_cell, dst_cell)

                            # merged cells preservation
                            if src_ws.merged_cells.ranges:
                                for merged_range in src_ws.merged_cells.ranges:
                                    new_ws.merge_cells(str(merged_range))
                                    top_left = src_ws.cell(merged_range.min_row, merged_range.min_col)
                                    new_ws.cell(merged_range.min_row, merged_range.min_col, top_left.value)

                            try:
                                for col_letter in src_ws.column_dimensions:
                                    if src_ws.column_dimensions[col_letter].width:
                                        new_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width
                                for row_idx in src_ws.row_dimensions:
                                    if src_ws.row_dimensions[row_idx].height:
                                        new_ws.row_dimensions[row_idx].height = src_ws.row_dimensions[row_idx].height
                            except Exception:
                                pass

                            file_buffer = BytesIO()
                            new_wb.save(file_buffer)
                            file_buffer.seek(0)
                            file_name = f"{clean_name(sheet_name)}.xlsx"
                            zip_file.writestr(file_name, file_buffer.read())

                            elapsed = time.time() - start_time
                            pct = int(((i + 1) / total) * 100)
                            render_gold_progress(progress_placeholder, pct, f"Splitting sheet {i+1}/{total}: {sheet_name}", elapsed)

                    elapsed = time.time() - start_time
                    render_gold_progress(progress_placeholder, 100, "✅ Splitting Completed Successfully!", elapsed)
                    time.sleep(0.25)
                    progress_placeholder.empty()

                    zip_buffer.seek(0)
                    st.success("🎉 Splitting completed successfully!")
                    st.download_button(
                        label="📥 Download Split Files (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name=f"SplitBySheets_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip",
                        mime="application/zip"
                    )
                    show_confetti()
else:
    st.info("📤 Please upload an Excel or CSV file to start splitting.")

# ----------------------- Merge Section (anchor) -----------------------
st.markdown("<hr style='border:1px dashed #FFD700; margin-top:18px; opacity:0.6;' />", unsafe_allow_html=True)
st.markdown('<a name="merge_section"></a>', unsafe_allow_html=True)
st.markdown("<h2 style='color:#FFD700;'>🔄 Merge Excel/CSV Files</h2>", unsafe_allow_html=True)

merge_files = st.file_uploader(
    "📤 Upload Excel or CSV Files to Merge",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    key=f"merge_uploader_{st.session_state.get('clear_counter',0)}"
)

if merge_files:
    # display uploaded
    st.markdown("### 📁 Files to merge:")
    for i, f in enumerate(merge_files):
        st.markdown(f"- {i+1}. {f.name} ({f.size//1024} KB)")

    if st.button("🗑️ Clear All Merged Files", key="clear_merge"):
        st.session_state.clear_counter = st.session_state.get('clear_counter',0) + 1
        st.experimental_rerun()

    if st.button("✨ Merge Files"):
        with st.spinner("Merging files..."):
            if LOTTIE_MERGE:
                st_lottie(LOTTIE_MERGE, height=140, key="lottie_merge_main")

            merge_progress_placeholder = st.empty()
            merge_start = time.time()

            try:
                all_excel = all(f.name.lower().endswith('.xlsx') for f in merge_files)
                total = len(merge_files) if len(merge_files) > 0 else 1

                if all_excel:
                    # Merge preserving formatting (header once + style)
                    merged_wb = Workbook()
                    merged_ws = merged_wb.active
                    merged_ws.title = "Merged_Data"
                    current_row = 1

                    for idx, file in enumerate(merge_files):
                        file_bytes = file.getvalue()
                        src_wb = load_workbook(filename=BytesIO(file_bytes), data_only=False)
                        src_ws = src_wb.active

                        # copy header only once (preserve styling)
                        if idx == 0:
                            for r in src_ws.iter_rows(min_row=1, max_row=1):
                                for cell in r:
                                    dst_cell = merged_ws.cell(current_row, cell.column, cell.value)
                                    copy_cell_style(cell, dst_cell)
                            current_row += 1

                        # copy data rows
                        for row in src_ws.iter_rows(min_row=2):
                            for cell in row:
                                dst_cell = merged_ws.cell(current_row, cell.column, cell.value)
                                copy_cell_style(cell, dst_cell)
                            current_row += 1

                        # copy merged cells with adjusted row offsets
                        if src_ws.merged_cells.ranges:
                            # to copy merged ranges correctly we need to map rows - simpler approach: copy same range as text (works when not offset)
                            for merged_range in src_ws.merged_cells.ranges:
                                try:
                                    # compute top-left in merged_ws relative to current insertion area
                                    # NOTE: since we're appending rows, merged ranges will map correctly only if positions are consistent
                                    merged_range_str = str(merged_range)
                                    # We will attempt to replicate merged ranges relative to the rows already placed by inspecting their coordinates
                                    # Simpler: reapply merged ranges for full columns where present (best-effort)
                                    merged_area = merged_range_str
                                    merged_ws.merge_cells(merged_area)
                                except Exception:
                                    pass

                        # copy column widths
                        try:
                            for col_letter in src_ws.column_dimensions:
                                if src_ws.column_dimensions[col_letter].width:
                                    merged_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width
                        except Exception:
                            pass

                        # update progress
                        pct = int(((idx + 1) / total) * 100)
                        elapsed = time.time() - merge_start
                        render_gold_progress(merge_progress_placeholder, pct, f"Merging file {idx+1}/{total}: {file.name}", elapsed)

                    # finalize
                    output_buffer = BytesIO()
                    merged_wb.save(output_buffer)
                    output_buffer.seek(0)

                    elapsed = time.time() - merge_start
                    render_gold_progress(merge_progress_placeholder, 100, "✅ Merging Completed Successfully!", elapsed)
                    time.sleep(0.25)
                    merge_progress_placeholder.empty()

                    st.success("✅ Merged successfully with original formatting preserved!")
                    st.download_button(
                        label="📥 Download Merged File (Formatted)",
                        data=output_buffer.getvalue(),
                        file_name="Merged_Consolidated_Formatted.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    show_confetti()

                else:
                    # Mixed CSV/Excel fallback using pandas concat
                    all_dfs = []
                    for i, file in enumerate(merge_files):
                        ext = file.name.split('.')[-1].lower()
                        if ext == "csv":
                            df = pd.read_csv(file)
                        else:
                            df = pd.read_excel(file)
                        all_dfs.append(df)
                        pct = int(((i + 1) / total) * 100)
                        elapsed = time.time() - merge_start
                        render_gold_progress(merge_progress_placeholder, pct, f"Merging file {i+1}/{total}: {file.name}", elapsed)

                    merged_df = pd.concat(all_dfs, ignore_index=True)
                    output_buffer = BytesIO()
                    merged_df.to_excel(output_buffer, index=False, engine='openpyxl')
                    output_buffer.seek(0)

                    elapsed = time.time() - merge_start
                    render_gold_progress(merge_progress_placeholder, 100, "✅ Merging Completed Successfully!", elapsed)
                    time.sleep(0.25)
                    merge_progress_placeholder.empty()

                    st.success("✅ Merged successfully (formatting not preserved for CSV/mixed files).")
                    st.download_button(
                        label="📥 Download Merged File (Excel)",
                        data=output_buffer.getvalue(),
                        file_name="Merged_Consolidated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    show_confetti()

            except Exception as e:
                st.error(f"❌ Error during merge: {e}")
                merge_progress_placeholder.empty()

# ----------------------- Image to PDF Section (anchor) -----------------------
st.markdown("<hr style='border:1px dashed #FFD700; margin-top:18px; opacity:0.6;' />", unsafe_allow_html=True)
st.markdown('<a name="imagetopdf"></a>', unsafe_allow_html=True)
st.markdown("<h2 style='color:#FFD700;'>📷 Convert Images to PDF</h2>", unsafe_allow_html=True)

uploaded_images = st.file_uploader(
    "📤 Upload JPG/JPEG/PNG Images to Convert to PDF",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
    key=f"image_uploader_{st.session_state.get('clear_counter',0)}"
)
if uploaded_images:
    st.markdown("### 📁 Uploaded Images:")
    for i, img in enumerate(uploaded_images):
        st.markdown(f"- {i+1}. {img.name} ({img.size//1024} KB)")

    if st.button("🖨️ Create PDF (CamScanner Style)"):
        with st.spinner("Enhancing images for PDF..."):
            try:
                import cv2
                import numpy as np
                def enhance_image_for_pdf(image_pil):
                    image = np.array(image_pil)
                    if image.shape[2] == 4:
                        image = cv2.cvtColor(image, cv2.COLOR_RGBA2RGB)
                    image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)
                    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
                    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
                    enhanced = clahe.apply(gray)
                    border_size = 20
                    bordered = cv2.copyMakeBorder(enhanced, border_size, border_size, border_size, border_size,
                                                  cv2.BORDER_CONSTANT, value=[255,255,255])
                    if bordered.dtype != np.uint8:
                        bordered = np.clip(bordered,0,255).astype(np.uint8)
                    result = cv2.cvtColor(bordered, cv2.COLOR_GRAY2RGB)
                    return Image.fromarray(result)

                first_image = Image.open(uploaded_images[0])
                first_image_enhanced = enhance_image_for_pdf(first_image)
                other_images = []
                for img_file in uploaded_images[1:]:
                    img = Image.open(img_file)
                    other_images.append(enhance_image_for_pdf(img).convert("RGB"))
                pdf_buffer = BytesIO()
                first_image_enhanced.save(pdf_buffer, format="PDF", save_all=True, append_images=other_images)
                pdf_buffer.seek(0)
                st.success("✅ Enhanced PDF created successfully!")
                st.download_button("📥 Download Enhanced PDF", pdf_buffer.getvalue(), "Enhanced_Images_CamScanner.pdf", "application/pdf")
                show_confetti()
            except ImportError:
                st.warning("⚠️ CamScanner effect requires 'opencv-python'. Install it to enable this feature.")
            except Exception as e:
                st.error(f"❌ Error creating enhanced PDF: {e}")

    if st.button("🖨️ Create PDF (Original Quality)"):
        with st.spinner("Converting images to PDF..."):
            try:
                first_image = Image.open(uploaded_images[0]).convert("RGB")
                other_images = [Image.open(x).convert("RGB") for x in uploaded_images[1:]]
                pdf_buffer = BytesIO()
                first_image.save(pdf_buffer, format="PDF", save_all=True, append_images=other_images)
                pdf_buffer.seek(0)
                st.success("✅ PDF created successfully!")
                st.download_button("📥 Download Original PDF", pdf_buffer.getvalue(), "Images_Combined.pdf", "application/pdf")
                show_confetti()
            except Exception as e:
                st.error(f"❌ Error creating PDF: {e}")
else:
    st.info("📤 Please upload one or more JPG/JPEG/PNG images to convert them into a single PDF file.")

# ----------------------- Dashboard Section (anchor) -----------------------
st.markdown("<hr style='border:1px dashed #FFD700; margin-top:18px; opacity:0.6;' />", unsafe_allow_html=True)
st.markdown('<a name="dashboard"></a>', unsafe_allow_html=True)
st.markdown("<h2 style='color:#FFD700;'>📊 Interactive Auto Dashboard Generator</h2>", unsafe_allow_html=True)

dashboard_file = st.file_uploader(
    "📊 Upload Excel or CSV File for Dashboard (Auto)",
    type=["xlsx", "csv"],
    key=f"dashboard_uploader_{st.session_state.get('clear_counter',0)}"
)
if dashboard_file:
    st.markdown("### 🔍 Data Preview")
    file_ext = dashboard_file.name.split('.')[-1].lower()
    if file_ext == "csv":
        df0 = pd.read_csv(dashboard_file)
        sheet_title = "CSV Data"
    else:
        df_dict = pd.read_excel(dashboard_file, sheet_name=None)
        sheet_names = list(df_dict.keys())
        selected_sheet_dash = st.selectbox("Select Sheet for Dashboard", sheet_names, key="sheet_dash")
        df0 = df_dict[selected_sheet_dash].copy()
        sheet_title = selected_sheet_dash

    st.dataframe(df0.head(), use_container_width=True)

    # Basic auto-detection & KPIs (kept concise)
    numeric_cols = df0.select_dtypes(include='number').columns.tolist()
    period_comparison = None
    # detect period columns like name_2024 or Q1 etc. (simplified)
    numeric_cols_in_long = numeric_cols.copy()
    if numeric_cols_in_long:
        user_measure_col = st.selectbox("🎯 Select Sales/Value Column (for KPIs & Charts)", numeric_cols_in_long)
        kpi_measure_col = user_measure_col
    else:
        kpi_measure_col = None

    # identify categorical columns
    cat_cols = [c for c in df0.columns if df0[c].dtype == object or df0[c].dtype.name.startswith("category")]
    st.sidebar.header("🔍 Filters")
    primary_filter_col = None
    if cat_cols:
        primary_filter_col = st.sidebar.selectbox("Primary Filter Column", ["-- None --"] + cat_cols, index=0)
        if primary_filter_col == "-- None --":
            primary_filter_col = None

    # simple filtering UI
    filtered = df0.copy()
    if primary_filter_col:
        vals = filtered[primary_filter_col].dropna().astype(str).unique().tolist()
        sel = st.sidebar.multiselect(f"Filter values for {primary_filter_col}", vals, default=vals)
        if sel:
            filtered = filtered[filtered[primary_filter_col].astype(str).isin(sel)]

    # KPIs
    st.markdown("### 🚀 KPIs")
    kpi_cols = st.columns(3)
    total_val = filtered[kpi_measure_col].sum() if kpi_measure_col in filtered.columns else None
    kpi_cols[0].markdown(f"<div style='color:#FFD700;font-weight:700'>Total</div><div style='font-size:18px'>{total_val if total_val is not None else 'N/A'}</div>", unsafe_allow_html=True)

    # Simple charts
    st.markdown("### 📊 Auto Charts")
    try:
        if kpi_measure_col:
            rep_col = _find_col(filtered, ["rep", "representative", "salesman", "employee", "name"])
            if rep_col:
                rep_data = filtered.groupby(rep_col)[kpi_measure_col].sum().sort_values(ascending=False).head(10)
                df_top = rep_data.reset_index().rename(columns={kpi_measure_col: "value"})
                fig = px.bar(df_top, x=rep_col, y="value", title="Top by Rep")
                st.plotly_chart(fig, use_container_width=True, theme="streamlit")
    except Exception as e:
        st.warning(f"Could not produce charts: {e}")

    # Export filtered data
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        filtered.to_excel(writer, index=False, sheet_name='Filtered_Data')
    excel_data = excel_buffer.getvalue()
    st.download_button("⬇️ Download Filtered Data (Excel)", excel_data, f"{_safe_name(sheet_title)}_Filtered.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("📤 Please upload an Excel or CSV file for dashboard generation.")

# ----------------------- Contact / Footer (anchor) -----------------------
st.markdown("<hr style='border:1px dashed #FFD700; margin-top:18px; opacity:0.6;' />", unsafe_allow_html=True)
st.markdown('<a name="contact"></a>', unsafe_allow_html=True)
st.markdown("""
<div style='text-align:center; color:#FFD700; font-weight:700;'>
    Contact / Support
</div>
<div style='text-align:center; color:white; margin-top:6px;'>
    WhatsApp: 01554694554 • Email: admin@example.com
</div>
<br>
<div style='text-align:center; color:#888; font-size:12px;'>Built with ❤️ — Averroes Pharma Tool</div>
""", unsafe_allow_html=True)

