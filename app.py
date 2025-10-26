# Modified Averroes Pharma File Splitter & Dashboard
# Applied user's requested UI & functionality changes (Navbar style, WhatsApp contact link, Clear All, KPI improvements, logo fixes, separators, Excel icon, etc.)
# NOTE: Replace or provide 'logo.png' in the same folder if you want a custom logo. If not present a fallback SVG will appear.

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

# ----------------------- Page Config -----------------------
st.set_page_config(
    page_title="Averroes Pharma — File Splitter & Dashboard",
    page_icon="🟩",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ----------------------- Helpers -----------------------
@st.cache_data(show_spinner=False)
def load_lottie_url(url: str):
    try:
        r = requests.get(url, timeout=8)
        if r.status_code == 200:
            return r.json()
    except Exception:
        return None
    return None

# small lotties (kept same keys as original but optional)
LOTTIE_SPLIT = load_lottie_url("https://assets9.lottiefiles.com/packages/lf20_wx9z5gxb.json")
LOTTIE_MERGE = load_lottie_url("https://assets10.lottiefiles.com/packages/lf20_cg3rwjul.json")
LOTTIE_IMAGE = load_lottie_url("https://assets2.lottiefiles.com/private_files/lf30_cgfdhxgx.json")
LOTTIE_DASH  = load_lottie_url("https://assets8.lottiefiles.com/packages/lf20_tno6cg2w.json")

# safe name
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

# gold progress renderer
def render_gold_progress(placeholder, percentage, message, elapsed_seconds):
    html = f"""
    <style>
    .gold-container {{ background-color: #001f3f; border: 1px solid #FFD700; border-radius: 12px; padding: 12px; margin-top: 12px; text-align: center; }}
    .gold-progress {{ width:100%; background-color:#001a2e; border-radius:20px; height:28px; overflow:hidden; }}
    .gold-fill {{ height:100%; width:{percentage}%; border-radius:20px; background: linear-gradient(90deg,#FFD700,#FFC107); line-height:28px; font-weight:700; color:#000; text-align:center; }}
    .gold-meta {{ color:#FFD700; font-weight:800; margin-bottom:6px; }}
    .gold-time {{ color:#fff; opacity:0.85; font-size:12px; margin-top:6px; }}
    </style>
    <div class="gold-container">
      <div class="gold-meta">{message}</div>
      <div class="gold-progress"><div class="gold-fill">{percentage}%</div></div>
      <div class="gold-time">Elapsed: {elapsed_seconds:.1f}s</div>
    </div>
    """
    placeholder.markdown(html, unsafe_allow_html=True)

# confetti
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
    components.html(confetti_js, height=0)

# ----------------------- CSS (improved Navbar, separators, logo) -----------------------
st.markdown("""
<style>
/* hide streamlit default */
#MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}

/* app background */
.stApp { background-color: #001529; color: #fff; font-family: 'Cairo', sans-serif; }

/* Top filler area: use to move navbar up and fill empty space */
.top-filler { background: linear-gradient(180deg, rgba(0,21,40,0.95), rgba(0,21,40,0.9)); padding:10px 18px; border-bottom:1px solid rgba(255,215,0,0.08); }

/* navbar */
.top-nav { display:flex; align-items:center; gap:20px; padding:8px 18px; }
.nav-left { display:flex; align-items:center; gap:12px; margin-right:auto; }
.logo-small { width:132px; height:auto; border-radius:8px; padding:6px; background:#012; display:inline-block; }
.nav-link { color:#FFD700; text-decoration:none; font-weight:700; font-size:15px; margin:0 10px; }
.nav-link:hover { color:#FFE97F; cursor:pointer; }
.nav-icon { display:inline-flex; align-items:center; gap:8px; }

/* golden separators between sections */
.gold-sep { border:0; height:2px; background: linear-gradient(90deg, rgba(0,0,0,0), #FFD700, rgba(0,0,0,0)); margin:20px 0; border-radius:3px; }

/* page title icon (excel-like) */
.page-icon { display:inline-block; width:26px; height:26px; border-radius:4px; background:#107C10; color:white; font-weight:900; text-align:center; line-height:26px; margin-right:8px; }

/* Contact button style inside navbar (whatsapp style) */
.contact-btn { background: linear-gradient(90deg,#f9f9f9,#ffeaa7); padding:6px 10px; border-radius:8px; border:1px solid #FFD700; font-weight:800; }

/* KPI cards */
.kpi-card { background:#001a2a; padding:12px; border-radius:10px; border:1px solid rgba(255,215,0,0.08); }

/* small responsive tweaks */
@media (max-width: 600px) {
  .logo-small { width:90px; }
  .nav-link { font-size:13px; }
}
</style>
""", unsafe_allow_html=True)

# ----------------------- Navbar HTML (with interactive Contact opening WhatsApp) -----------------------
# replace phone with user's WhatsApp number
WHATSAPP_NUMBER = "201554694554"  # international format without + and leading zeros; adjust if needed
# WhatsApp link: https://wa.me/{number}

navbar_html = f"""
<div class='top-filler'>
  <div class='top-nav'>
    <div class='nav-left'>
      <img src='logo.png' class='logo-small' onerror="this.src='data:image/svg+xml;utf8,<svg xmlns=\'http://www.w3.org/2000/svg\' width=\'300\' height=\'80\'><rect width=\'100%\' height=\'100%\' fill=\'%2300112b\' rx=\'8\' ry=\'8\'/><text x=\'20\' y=\'50\' font-size=\'28\' fill=\'%23ffd700\' font-family=\'Arial\' >Averroes</text></svg>'" />
      <div style='display:flex;align-items:center;gap:6px'>
        <div class='page-icon'>X</div>
        <div style='color:#FFD700;font-weight:800;font-size:18px;'>Tricks Excel File Splitter & Dashboard</div>
      </div>
    </div>
    <a class='nav-link' href='#home'>Home</a>
    <a class='nav-link' href='#split'>Split & Merge</a>
    <a class='nav-link' href='#imagetopdf'>Image → PDF</a>
    <a class='nav-link' href='#dashboard'>Auto Dashboard</a>
    <a class='nav-link' id='contact-link' href='#contact'>Contact</a>
    <div style='margin-left:12px;'>
      <button class='contact-btn' onclick="window.open('https://wa.me/{WHATSAPP_NUMBER}','_blank')">🟢 WhatsApp</button>
    </div>
  </div>
</div>
<script>
// smooth scroll behavior for anchors
document.querySelectorAll('.nav-link').forEach(a=>{{
  a.addEventListener('click', function(e){{
    e.preventDefault();
    var href = this.getAttribute('href');
    var el = document.querySelector(href);
    if(el){{ el.scrollIntoView({{behavior:'smooth'}}); }}
  }});
}});
</script>
"""

st.markdown(navbar_html, unsafe_allow_html=True)

# ----------------------- Header (removed manual author display per user request) -----------------------
st.markdown('<a name="home"></a>', unsafe_allow_html=True)
st.markdown("""
<div style='text-align:center; padding:10px 0'>
  <h1 style='color:#FFD700; margin:6px 0; font-size:34px;'>💼 Tricks Excel — File Splitter & Dashboard</h1>
  <div style='color:#ddd; font-size:14px;'>Split, Merge, Images → PDF & Auto KPI Dashboard</div>
</div>
<hr class='gold-sep' />
""", unsafe_allow_html=True)

# ----------------------- Split Section -----------------------
st.markdown('<a name="split"></a>', unsafe_allow_html=True)
st.markdown("""
<h2 style='color:#FFD700;'>✂ Split Excel / CSV File</h2>
""", unsafe_allow_html=True)

# Clear counter state (used to force re-render of uploaders)
if 'clear_counter' not in st.session_state:
    st.session_state['clear_counter'] = 0

uploaded_file = st.file_uploader(
    "📂 Upload Excel or CSV File (Splitter/Merge)",
    type=["xlsx", "csv"],
    accept_multiple_files=False,
    key=f"split_uploader_{st.session_state.get('clear_counter',0)}"
)

# Add Clear All Files button for split
col_clear = st.columns([1,9])
if col_clear[0].button("🗑️ Clear All files", key='clear_split'):
    st.session_state['clear_counter'] = st.session_state.get('clear_counter',0) + 1
    st.rerun()
def clean_name(name):
    name = str(name).strip()
    invalid_chars = r'[\\/*?:\[\]|<>"]'
    cleaned = re.sub(invalid_chars, '_', name)
    return cleaned[:30] if cleaned else "Sheet"

if uploaded_file:
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
                st_lottie(LOTTIE_SPLIT, height=120, key="lottie_split_main")

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
                                try:
                                    if cell.font: dst_cell.font = Font(name=cell.font.name, size=cell.font.size, bold=cell.font.bold)
                                except Exception:
                                    pass

                            # copy rows matching value
                            row_idx = 2
                            for row in ws.iter_rows(min_row=2):
                                cell_in_col = row[col_idx - 1]
                                if cell_in_col.value == value:
                                    for src_cell in row:
                                        dst_cell = new_ws.cell(row_idx, src_cell.column, src_cell.value)
                                    row_idx += 1

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

st.markdown("<hr class='gold-sep' />", unsafe_allow_html=True)

# ----------------------- Merge Section -----------------------
st.markdown('<a name="merge_section"></a>', unsafe_allow_html=True)
st.markdown("""
<h2 style='color:#FFD700;'>🔄 Merge Excel / CSV Files</h2>
""", unsafe_allow_html=True)

merge_files = st.file_uploader(
    "📤 Upload Excel or CSV Files to Merge",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    key=f"merge_uploader_{st.session_state.get('clear_counter',0)}"
)

if merge_files:
    st.markdown("### 📁 Files to merge:")
    for i, f in enumerate(merge_files):
        st.markdown(f"- {i+1}. {f.name} ({f.size//1024} KB)")

    if st.button("🗑️ Clear All Merged Files", key="clear_merge"):
        st.session_state.clear_counter = st.session_state.get('clear_counter',0) + 1
        st.rerun()
    if st.button("✨ Merge Files"):
        with st.spinner("Merging files..."):
            if LOTTIE_MERGE:
                st_lottie(LOTTIE_MERGE, height=120, key="lottie_merge_main")

            merge_progress_placeholder = st.empty()
            merge_start = time.time()

            try:
                all_excel = all(f.name.lower().endswith('.xlsx') for f in merge_files)
                total = len(merge_files) if len(merge_files) > 0 else 1

                if all_excel:
                    merged_wb = Workbook()
                    merged_ws = merged_wb.active
                    merged_ws.title = "Merged_Data"
                    current_row = 1

                    for idx, file in enumerate(merge_files):
                        file_bytes = file.getvalue()
                        src_wb = load_workbook(filename=BytesIO(file_bytes), data_only=False)
                        src_ws = src_wb.active

                        # copy header only once
                        if idx == 0:
                            for r in src_ws.iter_rows(min_row=1, max_row=1):
                                for cell in r:
                                    dst_cell = merged_ws.cell(current_row, cell.column, cell.value)
                            current_row += 1

                        for row in src_ws.iter_rows(min_row=2):
                            for cell in row:
                                dst_cell = merged_ws.cell(current_row, cell.column, cell.value)
                            current_row += 1

                        pct = int(((idx + 1) / total) * 100)
                        elapsed = time.time() - merge_start
                        render_gold_progress(merge_progress_placeholder, pct, f"Merging file {idx+1}/{total}: {file.name}", elapsed)

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
else:
    st.info("📤 Please upload Excel or CSV files to merge.")

st.markdown("<hr class='gold-sep' />", unsafe_allow_html=True)

# ----------------------- Image to PDF Section -----------------------
st.markdown('<a name="imagetopdf"></a>', unsafe_allow_html=True)
st.markdown("""
<h2 style='color:#FFD700;'>📷 Convert Images to PDF</h2>
""", unsafe_allow_html=True)

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

st.markdown("<hr class='gold-sep' />", unsafe_allow_html=True)

# ----------------------- Dashboard Section (KPIs & Charts) -----------------------
st.markdown('<a name="dashboard"></a>', unsafe_allow_html=True)
st.markdown("""
<h2 style='color:#FFD700;'>📊 Interactive Auto Dashboard Generator</h2>
""", unsafe_allow_html=True)

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

    # detect numeric columns and allow user to choose
    numeric_cols = df0.select_dtypes(include='number').columns.tolist()
    if numeric_cols:
        kpi_measure_col = st.selectbox("🎯 Select Measure Column (for KPIs & Charts)", numeric_cols)
    else:
        kpi_measure_col = None

    # categorical columns detection for grouping/filtering
    cat_cols = [c for c in df0.columns if df0[c].dtype == object or df0[c].dtype.name.startswith("category")]
    primary_filter_col = None
    if cat_cols:
        primary_filter_col = st.sidebar.selectbox("Primary Filter Column", ["-- None --"] + cat_cols, index=0)
        if primary_filter_col == "-- None --": primary_filter_col = None

    filtered = df0.copy()
    if primary_filter_col:
        vals = filtered[primary_filter_col].dropna().astype(str).unique().tolist()
        sel = st.sidebar.multiselect(f"Filter values for {primary_filter_col}", vals, default=vals)
        if sel:
            filtered = filtered[filtered[primary_filter_col].astype(str).isin(sel)]

    # KPI cards row
    st.markdown("### 🚀 KPIs")
    kpi_cols = st.columns(3)
    if kpi_measure_col:
        total_val = filtered[kpi_measure_col].sum()
        avg_val = filtered[kpi_measure_col].mean()
        max_val = filtered[kpi_measure_col].max()
        kpi_cols[0].markdown(f"<div class='kpi-card'><div style='color:#FFD700;font-weight:800'>Total</div><div style='font-size:20px'>{total_val:,.2f}</div></div>", unsafe_allow_html=True)
        kpi_cols[1].markdown(f"<div class='kpi-card'><div style='color:#FFD700;font-weight:800'>Average</div><div style='font-size:20px'>{avg_val:,.2f}</div></div>", unsafe_allow_html=True)
        kpi_cols[2].markdown(f"<div class='kpi-card'><div style='color:#FFD700;font-weight:800'>Max</div><div style='font-size:20px'>{max_val:,.2f}</div></div>", unsafe_allow_html=True)
    else:
        st.info("No numeric columns detected — KPIs unavailable.")

    st.markdown("### 📊 Auto Charts")
    # Bar chart: top by rep or top categorical
    try:
        if kpi_measure_col:
            group_col = _find_col(filtered, ["rep", "representative", "salesman", "employee", "name"]) or (cat_cols[0] if cat_cols else None)
            if group_col:
                rep_data = filtered.groupby(group_col)[kpi_measure_col].sum().sort_values(ascending=False).head(10)
                df_top = rep_data.reset_index().rename(columns={kpi_measure_col: "value"})
                fig_bar = px.bar(df_top, x=group_col, y="value", title=f"Top by {group_col}")
                st.plotly_chart(fig_bar, use_container_width=True, theme="streamlit")

                # pie chart for top categories
                fig_pie = px.pie(df_top, names=group_col, values="value", title=f"Share by {group_col}")
                st.plotly_chart(fig_pie, use_container_width=True, theme="streamlit")

                # optional time series if date-like col exists
                date_col = _find_col(filtered, ["date", "month", "year", "day"]) or None
                if date_col and pd.api.types.is_datetime64_any_dtype(filtered[date_col]):
                    ts = filtered.copy()
                    ts[date_col] = pd.to_datetime(ts[date_col])
                    ts_grouped = ts.groupby(pd.Grouper(key=date_col, freq='M'))[kpi_measure_col].sum().reset_index()
                    if not ts_grouped.empty:
                        fig_ts = px.line(ts_grouped, x=date_col, y=kpi_measure_col, title='Trend over time')
                        st.plotly_chart(fig_ts, use_container_width=True, theme="streamlit")
    except Exception as e:
        st.warning(f"Could not produce charts: {e}")

    # export filtered
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        filtered.to_excel(writer, index=False, sheet_name='Filtered_Data')
    excel_data = excel_buffer.getvalue()
    st.download_button("⬇️ Download Filtered Data (Excel)", excel_data, f"{_safe_name(sheet_title)}_Filtered.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("📤 Please upload an Excel or CSV file for dashboard generation.")

st.markdown("<hr class='gold-sep' />", unsafe_allow_html=True)

# ----------------------- Contact / Footer -----------------------
st.markdown('<a name="contact"></a>', unsafe_allow_html=True)
st.markdown("""
<div style='text-align:center; color:#FFD700; font-weight:800;'>Contact / Support</div>
<div style='text-align:center; color:white; margin-top:6px;'>
  <button style='padding:8px 12px;border-radius:8px;border:1px solid #FFD700;background:#002233;color:#fff;font-weight:700' onclick="window.open('https://wa.me/201554694554','_blank')">🟢 Message on WhatsApp</button>
  <div style='margin-top:8px;color:#ddd;font-size:14px;'>Email: lmohamedomar825@Gmail.com</div>
</div>
<br>
<div style='text-align:center; color:#888; font-size:12px;'>Built with ❤️ — Averroes Pharma Tool</div>
""", unsafe_allow_html=True)

# End of file


