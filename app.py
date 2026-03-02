# -*- coding: utf-8 -*-
"""
Streamlit App — Light Theme Refresh (UI Polished, English UI)
- Default: Light, clean palette (subtle gray background)
- Optional: Dark mode toggle in sidebar
- Tools: Split / Merge / Dynamic Excel Processor / Images → PDF
"""
import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
import re
import os
import base64
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from PIL import Image

# Optional animations
try:
    from streamlit_lottie import st_lottie  # type: ignore
    import requests  # type: ignore
except Exception:
    st_lottie = None
    requests = None

def load_lottie_url(url: str):
    if not requests:
        return None
    try:
        r = requests.get(url, timeout=8)
        if r.status_code == 200:
            return r.json()
    except Exception:
        return None
    return None

# ------------------ Page Setup ------------------
st.set_page_config(
    page_title="Tricks For Excel — Split/Merge & PDF Tools",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

LOTTIE_SPLIT = load_lottie_url("https://assets9.lottiefiles.com/packages/lf20_wx9z5gxb.json")
LOTTIE_MERGE = load_lottie_url("https://assets10.lottiefiles.com/packages/lf20_cg3rwjul.json")

# =================== THEME TOGGLE (optional) ===================
if 'ui_theme' not in st.session_state:
    st.session_state.ui_theme = 'Light'

with st.sidebar:
    st.markdown("### 🎨 Theme")
    st.session_state.ui_theme = st.radio("Choose theme", ["Light", "Dark"], index=0)

is_dark = st.session_state.ui_theme == 'Dark'

# ------------------ Custom CSS (Light-first) ------------------
colors_light = {
    'primary':  '#2563eb',  # primary blue
    'primary2': '#60a5fa',  # lighter blue gradient
    'accent':   '#22c55e',  # accent green
    'bg':       '#f5f6fa',  # very light gray app background
    'card':     '#ffffff',  # cards
    'card2':    '#eef2ff',  # subtle header tint
    'text':     '#0f172a',  # dark text
    'muted':    '#475569',  # secondary text
    'border':   '#e5e7eb',  # light borders
}
colors_dark = {
    'primary':  '#60a5fa',
    'primary2': '#93c5fd',
    'accent':   '#34d399',
    'bg':       '#0b1220',
    'card':     '#111827',
    'card2':    '#0f172a',
    'text':     '#e5e7eb',
    'muted':    '#94a3b8',
    'border':   '#1f2937',
}

C = colors_dark if is_dark else colors_light

custom_css = f"""
<style>
:root {{
--primary: {C['primary']};
--primary-2: {C['primary2']};
--accent: {C['accent']};
--app-bg: {C['bg']};
--bg-elev: {C['card']};
--bg-elev-2: {C['card2']};
--text: {C['text']};
--muted: {C['muted']};
--border: {C['border']};
}}
html, body {{ background: var(--app-bg) !important; }}
section.main > div {{ padding-top: 10px; }}
html, body, [class^="css"]  {{ font-family: 'Segoe UI', system-ui, -apple-system, Cairo, Tahoma, sans-serif; color: var(--text); }}
/* Header */
.app-header {{
display:flex; align-items:center; justify-content:center; gap:16px; padding:18px 20px;
background: linear-gradient(135deg, var(--bg-elev-2), #ffffff10);
border: 1px solid var(--border); border-radius: 16px;
box-shadow: 0 6px 18px rgba(17,24,39,.06);
}}
.app-titlewrap {{ text-align:center; }}
.app-title {{ margin:0; font-weight:800; letter-spacing:.2px; color: var(--text); }}
.app-sub {{ margin:6px 0 0; color: var(--muted); font-size: 14.5px; }}
.app-logo {{
width: 170px; height: auto; border-radius: 12px;
box-shadow: 0 8px 22px rgba(0,0,0,.12);
}}
/* Section card */
.card {{
border:1px solid var(--border); border-radius:16px; padding:18px 18px 10px; background: var(--bg-elev);
margin: 10px 0 22px; box-shadow: 0 2px 12px rgba(17,24,39,.07);
}}
.card h3, .card h2, .card h4 {{ margin-top:0; display:flex; align-items:center; gap:8px; color: var(--text); }}
/* Hints - clear & highlighted */
.hint {{
display:block;
background: var(--bg-elev-2);
border-left: 4px solid var(--primary);
color: var(--muted);
font-size: 14.5px;
padding: 10px 12px;
border-radius: 10px;
margin: -4px 0 10px 0;
}}
/* Buttons */
.stButton > button {{
border-radius:12px; padding:10px 14px; font-weight:600; border:1px solid transparent;
background: linear-gradient(135deg, var(--primary), var(--primary-2)); color:#fff; box-shadow: 0 4px 14px rgba(45,114,217,.25);
}}
.stButton > button:hover {{ filter:brightness(1.06); transform: translateY(-1px); }}
.stButton > button:active {{ transform: translateY(0); }}
/* File uploader */
[data-testid="stFileUploader"] {{
background: var(--bg-elev-2);
padding:12px; border-radius: 12px; border:1px dashed var(--border);
}}
/* Download buttons */
.stDownloadButton > button {{
border-radius:10px; border:1px solid var(--border); background: var(--bg-elev-2); color: var(--text);
}}
/* Dataframe wrapper */
.css-1m1b9qw, .stDataFrame {{ border-radius: 10px; overflow:hidden; border:1px solid var(--border); }}
/* Divider */
hr {{ border: none; height: 1px; background: var(--border); margin: 14px 0; }}
/* Force app background containers */
[data-testid="stAppViewContainer"] > .main {{ background: var(--app-bg); }}
[data-testid="stHeader"] {{ background: var(--app-bg); border-bottom: 0; }}
</style>
"""

st.markdown(custom_css, unsafe_allow_html=True)

# ------------------ Helpers ------------------
def display_uploaded_files(file_list, file_type="Files"):
    if file_list:
        st.markdown("**Uploaded files:**")
        for i, f in enumerate(file_list):
            st.caption(f"{i+1}. {f.name} — {f.size//1024} KB")

def _safe_name(s):
    return re.sub(r"[^A-Za-z0-9_-]+", "_", str(s))

def get_image_as_base64(image_path):
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except Exception:
        return None

# ------------------ Header ------------------
logo_b64 = get_image_as_base64("logo.png")
header_html = f"""
<div class="app-header">
{('<img class="app-logo" src="data:image/png;base64,' + logo_b64 + '" alt="Logo" />') if logo_b64 else ''}
<div class="app-titlewrap">
<h2 class="app-title">Tricks For Excel</h2>
<p class="app-sub">Quick tools for Excel & Images • Split • Merge • Processor • PDF</p>
</div>
</div>
"""
st.markdown(header_html, unsafe_allow_html=True)

if 'clear_counter' not in st.session_state:
    st.session_state.clear_counter = 0

# ===================== Split Card =====================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### ✂️ Split Excel/CSV File")
    st.markdown('<span class="hint">Upload an Excel or CSV file, then select the column to split by. A ZIP will be generated with one file per value.</span>', unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader(
        "📂 Upload Excel or CSV",
        type=["xlsx", "csv"],
        accept_multiple_files=False,
        key=f"split_uploader_{st.session_state.clear_counter}",
    )
    
    if uploaded_file:
        display_uploaded_files([uploaded_file])
        c1, c2 = st.columns([1,1])
        with c1:
            if st.button("🧹 Clear file", key="clear_split"):
                st.session_state.clear_counter += 1
                st.rerun()
        
        try:
            file_ext = uploaded_file.name.split(".")[-1].lower()
            df = None
            original_wb = None
            selected_sheet = None

            if file_ext == "csv":
                df = pd.read_csv(uploaded_file)
                selected_sheet = "Sheet1"
                st.success("✅ CSV file uploaded successfully")
            else:
                input_bytes = uploaded_file.getvalue()
                original_wb = load_workbook(filename=BytesIO(input_bytes), data_only=False)
                sheet_names = original_wb.sheetnames
                selected_sheet = st.selectbox("Select sheet to split", sheet_names)
                df = pd.read_excel(BytesIO(input_bytes), sheet_name=selected_sheet)
            
            st.dataframe(df.head(200), use_container_width=True)
            
            col_to_split = st.selectbox("Select column to split by", df.columns)
            
            split_option = "Split by Column Values"
            if file_ext == "xlsx":
                split_option = st.radio(
                    "Split method:",
                    ["Split by Column Values", "Split Each Sheet into Separate File"],
                    horizontal=True,
                )

            if st.button("🚀 Start Split"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                with st.spinner("Processing..."):
                    if st_lottie and LOTTIE_SPLIT:
                        st_lottie(LOTTIE_SPLIT, height=110, key="lottie_split")

                    def clean_name(name: str) -> str:
                        name = str(name).strip()
                        invalid_chars = r'[\\/*?:\[\]<>:"\']'
                        cleaned = re.sub(invalid_chars, "_", name)
                        return cleaned[:30] if cleaned else "Sheet"

                    total_steps = 100
                    current_step = 0

                    if file_ext == "csv":
                        unique_values = df[col_to_split].dropna().unique()
                        zip_buffer = BytesIO()
                        total_items = len(unique_values)
                        
                        with ZipFile(zip_buffer, "w") as zip_file:
                            for i, value in enumerate(unique_values):
                                status_text.text(f"Processing value: {value} ({i+1}/{total_items})")
                                filtered_df = df[df[col_to_split] == value]
                                csv_buffer = BytesIO()
                                filtered_df.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                                csv_buffer.seek(0)
                                zip_file.writestr(f"{clean_name(value)}.csv", csv_buffer.read())
                                current_step = int((i + 1) / total_items * 100)
                                progress_bar.progress(current_step)
                        
                        zip_buffer.seek(0)
                        st.success("🎉 Split completed! ZIP is ready.")
                        st.download_button(
                            "⬇️ Download (ZIP)",
                            zip_buffer.getvalue(),
                            file_name=f"Split_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip",
                            mime="application/zip"
                        )
                    else:
                        ws = original_wb[selected_sheet]
                        if split_option == "Split by Column Values":
                            col_idx = df.columns.get_loc(col_to_split) + 1
                            unique_values = df[col_to_split].dropna().unique()
                            zip_buffer = BytesIO()
                            total_items = len(unique_values)

                            with ZipFile(zip_buffer, "w") as zip_file:
                                for i, value in enumerate(unique_values):
                                    status_text.text(f"Processing value: {value} ({i+1}/{total_items})")
                                    new_wb = Workbook()
                                    default_ws = new_wb.active
                                    new_wb.remove(default_ws)
                                    new_ws = new_wb.create_sheet(title=clean_name(value))
                                    
                                    # Header
                                    for cell in ws[1]:
                                        dst = new_ws.cell(1, cell.column, cell.value)
                                        if cell.has_style:
                                            try:
                                                if cell.font: dst.font = Font(name=cell.font.name, size=cell.font.size, bold=cell.font.bold, italic=cell.font.italic, color=cell.font.color)
                                                if cell.fill and cell.fill.fill_type: dst.fill = PatternFill(fill_type=cell.fill.fill_type, start_color=cell.fill.start_color, end_color=cell.fill.end_color)
                                                if cell.alignment: dst.alignment = Alignment(horizontal=cell.alignment.horizontal, vertical=cell.alignment.vertical, wrap_text=cell.alignment.wrap_text)
                                                dst.number_format = cell.number_format
                                            except Exception: pass
                                    
                                    # Rows
                                    row_out = 2
                                    for row in ws.iter_rows(min_row=2):
                                        if row[col_idx - 1].value == value:
                                            for src in row:
                                                dst = new_ws.cell(row_out, src.column, src.value)
                                                if src.has_style:
                                                    try:
                                                        if src.font: dst.font = Font(name=src.font.name, size=src.font.size, bold=src.font.bold, italic=src.font.italic, color=src.font.color)
                                                        if src.fill and src.fill.fill_type: dst.fill = PatternFill(fill_type=src.fill.fill_type, start_color=src.fill.start_color, end_color=src.fill.end_color)
                                                        if src.alignment: dst.alignment = Alignment(horizontal=src.alignment.horizontal, vertical=src.alignment.vertical, wrap_text=src.alignment.wrap_text)
                                                        dst.number_format = src.number_format
                                                    except Exception: pass
                                            row_out += 1
                                    
                                    # Column widths
                                    try:
                                        for col_letter in ws.column_dimensions:
                                            width = ws.column_dimensions[col_letter].width
                                            if width: new_ws.column_dimensions[col_letter].width = width
                                    except Exception: pass
                                    
                                    fb = BytesIO()
                                    new_wb.save(fb)
                                    fb.seek(0)
                                    zip_file.writestr(f"{clean_name(value)}.xlsx", fb.read())
                                    
                                    current_step = int((i + 1) / total_items * 100)
                                    progress_bar.progress(current_step)

                            zip_buffer.seek(0)
                            st.success("🎉 Split completed! ZIP is ready.")
                            st.download_button(
                                "⬇️ Download (ZIP)",
                                zip_buffer.getvalue(),
                                file_name=f"Split_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip",
                                mime="application/zip"
                            )
                        else:
                            zip_buffer = BytesIO()
                            total_sheets = len(original_wb.sheetnames)
                            with ZipFile(zip_buffer, "w") as zip_file:
                                for i, sheet_name in enumerate(original_wb.sheetnames):
                                    status_text.text(f"Processing sheet: {sheet_name} ({i+1}/{total_sheets})")
                                    new_wb = Workbook()
                                    default_ws = new_wb.active
                                    new_wb.remove(default_ws)
                                    new_ws = new_wb.create_sheet(title=sheet_name)
                                    src_ws = original_wb[sheet_name]
                                    
                                    for row in src_ws.iter_rows():
                                        for src_cell in row:
                                            dst = new_ws.cell(src_cell.row, src_cell.column, src_cell.value)
                                            if src_cell.has_style:
                                                try:
                                                    if src_cell.font: dst.font = src_cell.font
                                                    if src_cell.fill and src_cell.fill.fill_type: dst.fill = src_cell.fill
                                                    if src_cell.alignment: dst.alignment = src_cell.alignment
                                                    dst.number_format = src_cell.number_format
                                                except Exception: pass
                                    
                                    try:
                                        for col_letter in src_ws.column_dimensions:
                                            if src_ws.column_dimensions[col_letter].width:
                                                new_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width
                                    except Exception: pass
                                    
                                    fb = BytesIO()
                                    new_wb.save(fb)
                                    fb.seek(0)
                                    zip_file.writestr(f"{_safe_name(sheet_name)}.xlsx", fb.read())
                                    
                                    current_step = int((i + 1) / total_sheets * 100)
                                    progress_bar.progress(current_step)
                            
                            zip_buffer.seek(0)
                            st.success("🎉 Split by sheets completed! ZIP is ready.")
                            st.download_button(
                                "⬇️ Download (ZIP)",
                                zip_buffer.getvalue(),
                                file_name=f"SplitBySheets_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip",
                                mime="application/zip"
                            )
                    
                    progress_bar.progress(100)
                    status_text.text("Done!")

        except Exception as e:
            st.error(f"❌ Error while splitting: {str(e)}")
    st.markdown('</div>', unsafe_allow_html=True)

# ===================== Merge Card =====================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 🔁 Merge Excel/CSV Files")
    st.markdown('<span class="hint">Upload multiple files and they will be merged into one file (Excel formatting is preserved where possible).</span>', unsafe_allow_html=True)
    
    merge_files = st.file_uploader(
        "📂 Upload Excel/CSV files to merge",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        key=f"merge_uploader_{st.session_state.clear_counter}",
    )
    
    if merge_files:
        display_uploaded_files(merge_files)
        c1, c2 = st.columns([1,1])
        with c1:
            if st.button("🧹 Clear files", key="clear_merge"):
                st.session_state.clear_counter += 1
                st.rerun()
        
        with c2:
            if st.button("✨ Merge files"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                with st.spinner("Merging..."):
                    if st_lottie and LOTTIE_MERGE:
                        st_lottie(LOTTIE_MERGE, height=100, key="lottie_merge")
                    try:
                        all_excel = all(f.name.lower().endswith('.xlsx') for f in merge_files)
                        total_files = len(merge_files)
                        
                        if all_excel:
                            merged_wb = Workbook()
                            merged_ws = merged_wb.active
                            merged_ws.title = "Merged_Data"
                            current_row = 1
                            
                            for idx, file in enumerate(merge_files):
                                status_text.text(f"Merging file: {file.name} ({idx+1}/{total_files})")
                                file_bytes = file.getvalue()
                                src_wb = load_workbook(filename=BytesIO(file_bytes), data_only=False)
                                src_ws = src_wb.active
                                
                                # Header from first file
                                if idx == 0:
                                    for row in src_ws.iter_rows(min_row=1, max_row=1):
                                        for cell in row:
                                            dst = merged_ws.cell(current_row, cell.column, cell.value)
                                            if cell.has_style:
                                                try:
                                                    if cell.font: dst.font = cell.font
                                                    if cell.fill and cell.fill.fill_type: dst.fill = cell.fill
                                                    if cell.alignment: dst.alignment = cell.alignment
                                                    dst.number_format = cell.number_format
                                                except Exception: pass
                                    current_row += 1
                                
                                # Data rows
                                for row in src_ws.iter_rows(min_row=2):
                                    for cell in row:
                                        dst = merged_ws.cell(current_row, cell.column, cell.value)
                                        if cell.has_style:
                                            try:
                                                if cell.font: dst.font = cell.font
                                                if cell.fill and cell.fill.fill_type: dst.fill = cell.fill
                                                if cell.alignment: dst.alignment = cell.alignment
                                                dst.number_format = cell.number_format
                                            except Exception: pass
                                    current_row += 1
                                
                                # Column widths
                                try:
                                    for col_letter in src_ws.column_dimensions:
                                        width = src_ws.column_dimensions[col_letter].width
                                        if width: merged_ws.column_dimensions[col_letter].width = width
                                except Exception: pass
                                
                                progress_bar.progress(int((idx + 1) / total_files * 100))

                            out = BytesIO()
                            merged_wb.save(out)
                            out.seek(0)
                            st.success("✅ Merge completed")
                            st.download_button(
                                "⬇️ Download file",
                                out.getvalue(),
                                file_name="Merged_Consolidated_Formatted.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            # Merge CSVs / mix as DataFrames
                            all_dfs = []
                            for idx, file in enumerate(merge_files):
                                status_text.text(f"Reading file: {file.name} ({idx+1}/{total_files})")
                                ext = file.name.split(".")[-1].lower()
                                df = pd.read_csv(file) if ext == "csv" else pd.read_excel(file)
                                all_dfs.append(df)
                                progress_bar.progress(int((idx + 1) / total_files * 50)) # Reading phase
                            
                            status_text.text("Consolidating data...")
                            merged_df = pd.concat(all_dfs, ignore_index=True)
                            progress_bar.progress(80)
                            
                            out = BytesIO()
                            merged_df.to_excel(out, index=False, engine='openpyxl')
                            out.seek(0)
                            progress_bar.progress(100)
                            
                            st.success("✅ Merge completed")
                            st.download_button(
                                "⬇️ Download file",
                                out.getvalue(),
                                file_name="Merged_Consolidated.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    except Exception as e:
                        st.error(f"❌ Error while merging: {str(e)}")
    st.markdown('</div>', unsafe_allow_html=True)

# ===================== Dynamic Excel Processor Card =====================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 🧰 Dynamic Excel Processor")
    st.markdown('<span class="hint">Configure columns dynamically: Rename, Delete, and Reorder without changing code.</span>', unsafe_allow_html=True)
    
    proc_file = st.file_uploader(
        "📂 Upload Excel file to process (xlsx/xlsm)",
        type=["xlsx", "xlsm"],
        accept_multiple_files=False,
        key=f"processor_uploader_{st.session_state.clear_counter}",
    )
    
    if proc_file:
        st.write("**File:**", proc_file.name)
        
        # Load file to get headers first
        try:
            temp_wb = load_workbook(proc_file, data_only=True)
            temp_ws = temp_wb.active
            headers = [temp_ws.cell(1, col).value for col in range(1, temp_ws.max_column + 1)]
            headers = [h for h in headers if h is not None] # Filter None
            temp_wb.close()
            
            if not headers:
                st.error("❌ No headers found in the first row.")
            else:
                st.info(f"Detected {len(headers)} columns: {', '.join(headers)}")
                
                # --- Configuration Section ---
                st.markdown("#### ⚙️ Configuration")
                
                # 1. Rename Columns
                st.markdown("**1. Rename Columns** (Old Name -> New Name)")
                rename_container = st.container()
                with rename_container:
                    num_renames = st.number_input("Number of columns to rename", min_value=0, max_value=len(headers), value=0, step=1)
                    rename_map = {}
                    for i in range(num_renames):
                        c1, c2 = st.columns(2)
                        with c1:
                            old_name = st.selectbox(f"Old Name #{i+1}", headers, key=f"rename_old_{i}")
                        with c2:
                            new_name = st.text_input(f"New Name #{i+1}", key=f"rename_new_{i}")
                        if old_name and new_name:
                            rename_map[old_name] = new_name
                
                # 2. Delete Columns
                st.markdown("**2. Delete Columns**")
                cols_to_delete = st.multiselect(
                    "Select columns to remove",
                    options=headers,
                    default=[],
                    help="Hold Ctrl/Cmd to select multiple"
                )
                
                # 3. Reorder Columns
                st.markdown("**3. Final Column Order**")
                st.caption("Select columns in the desired order. Unselected columns will be dropped.")
                final_order = st.multiselect(
                    "Select and order columns",
                    options=headers,
                    default=headers, # Default to all
                    help="The order you select here will be the final order in the output file."
                )
                
                # --- Process Button ---
                if st.button("⚙️ Start Processing"):
                    if not final_order:
                        st.error("❌ Please select at least one column for the final order.")
                    else:
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        try:
                            status_text.text("Loading workbook...")
                            wb = load_workbook(proc_file, data_only=False)
                            ws = wb.active
                            progress_bar.progress(10)
                            
                            header_row = 1
                            current_headers = [ws.cell(header_row, col).value for col in range(1, ws.max_column + 1)]
                            header_to_idx = {h: i+1 for i, h in enumerate(current_headers) if h is not None}
                            
                            # Step 1: Rename
                            status_text.text("Renaming columns...")
                            renamed_count = 0
                            for old_name, new_name in rename_map.items():
                                if old_name in header_to_idx:
                                    cidx = header_to_idx[old_name]
                                    ws.cell(header_row, cidx).value = new_name
                                    # Update mapping
                                    del header_to_idx[old_name]
                                    header_to_idx[new_name] = cidx
                                    renamed_count += 1
                            progress_bar.progress(30)
                            
                            # Refresh headers after rename
                            current_headers = [ws.cell(header_row, col).value for col in range(1, ws.max_column + 1)]
                            header_to_idx = {h: i+1 for i, h in enumerate(current_headers) if h is not None}
                            
                            # Step 2: Delete
                            status_text.text("Deleting selected columns...")
                            # Identify indices to delete based on CURRENT headers (after rename)
                            # We need to map the user's selection (which might be old names if not renamed, or new names)
                            # Since multiselect uses the initial 'headers' list, we need to be careful.
                            # Strategy: The user selected names from the ORIGINAL list.
                            # If a name was renamed, the user selected the OLD name. We need to find its NEW position.
                            # If a name wasn't renamed, it stays same.
                            
                            indices_to_delete = []
                            for user_sel in cols_to_delete:
                                # Check if this user_sel was renamed
                                target_name = rename_map.get(user_sel, user_sel)
                                if target_name in header_to_idx:
                                    indices_to_delete.append(header_to_idx[target_name])
                            
                            indices_to_delete.sort(reverse=True)
                            deleted_count = 0
                            for cidx in indices_to_delete:
                                ws.delete_cols(cidx)
                                deleted_count += 1
                            
                            progress_bar.progress(50)
                            
                            # Refresh headers and mapping again after deletion
                            current_headers = [ws.cell(header_row, col).value for col in range(1, ws.max_column + 1)]
                            header_to_idx = {h: i+1 for i, h in enumerate(current_headers) if h is not None}
                            
                            # Step 3: Reorder & Create New Workbook
                            status_text.text("Reordering columns and creating new file...")
                            
                            # Map final_order names to current indices
                            # Note: final_order contains names from the ORIGINAL list.
                            # We need to resolve them to current names/indices.
                            col_mapping_indices = []
                            valid_final_columns = []
                            
                            for name in final_order:
                                # Resolve name
                                resolved_name = rename_map.get(name, name)
                                # Check if this column still exists (wasn't deleted)
                                # A column is deleted if it was in cols_to_delete
                                if name in cols_to_delete:
                                    continue # Skip if deleted
                                
                                if resolved_name in header_to_idx:
                                    col_mapping_indices.append(header_to_idx[resolved_name])
                                    valid_final_columns.append(resolved_name)
                                else:
                                    st.warning(f"Column '{name}' (resolved to '{resolved_name}') not found or already deleted.")
                            
                            if not col_mapping_indices:
                                st.error("No valid columns remaining to process.")
                                progress_bar.progress(0)
                            else:
                                new_wb = Workbook()
                                new_ws = new_wb.active
                                
                                total_rows = ws.max_row
                                for row_idx in range(1, total_rows + 1):
                                    if row_idx % 100 == 0:
                                        prog = 50 + int((row_idx / total_rows) * 40)
                                        progress_bar.progress(prog)
                                        status_text.text(f"Processing row {row_idx}/{total_rows}...")
                                        
                                    for new_col_idx, old_col_idx in enumerate(col_mapping_indices, start=1):
                                        old_cell = ws.cell(row_idx, old_col_idx)
                                        new_cell = new_ws.cell(row_idx, new_col_idx)
                                        new_cell.value = old_cell.value
                                        
                                        if old_cell.has_style:
                                            try:
                                                if old_cell.font: new_cell.font = old_cell.font
                                                if old_cell.fill and old_cell.fill.fill_type: new_cell.fill = old_cell.fill
                                                if old_cell.alignment: new_cell.alignment = old_cell.alignment
                                                new_cell.number_format = old_cell.number_format
                                            except Exception: pass
                                
                                # Set column widths
                                for c in range(1, len(valid_final_columns) + 1):
                                    new_ws.column_dimensions[get_column_letter(c)].width = 15
                                    # Optionally set header value explicitly if not copied correctly (though loop covers row 1)
                                    if row_idx >= 1: # Ensure header is set
                                         new_ws.cell(1, c).value = valid_final_columns[c-1]

                                progress_bar.progress(95)
                                status_text.text("Saving file...")
                                
                                out_buf = BytesIO()
                                new_wb.save(out_buf)
                                out_buf.seek(0)
                                
                                progress_bar.progress(100)
                                status_text.text("Done!")
                                
                                st.success("✅ Processing completed successfully!")
                                base = os.path.splitext(proc_file.name)[0]
                                st.download_button(
                                    "⬇️ Download Processed File",
                                    out_buf.getvalue(),
                                    file_name=f"{_safe_name(base)}_processed.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        except Exception as e:
                            st.error(f"❌ Error while processing: {str(e)}")
                            import traceback
                            st.code(traceback.format_exc())
        except Exception as e:
            st.error(f"❌ Error reading file structure: {str(e)}")
            
    st.markdown('</div>', unsafe_allow_html=True)

# ===================== Images → PDF Card =====================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 🖼️ Convert Images to PDF")
    st.markdown('<span class="hint">Upload one or more images and they will be combined into a single PDF file while preserving original quality.</span>', unsafe_allow_html=True)
    
    uploaded_images = st.file_uploader(
        "📂 Upload JPG/PNG images",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key=f"image_uploader_{st.session_state.clear_counter}",
    )
    
    if uploaded_images:
        display_uploaded_files(uploaded_images, "Images")
        c1, c2 = st.columns([1,1])
        with c1:
            if st.button("🧹 Clear images", key="clear_images"):
                st.session_state.clear_counter += 1
                st.rerun()
        
        with c2:
            if st.button("🖨️ Create PDF"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                with st.spinner("Creating PDF..."):
                    try:
                        total_imgs = len(uploaded_images)
                        status_text.text(f"Processing image 1/{total_imgs}...")
                        first = Image.open(uploaded_images[0]).convert("RGB")
                        progress_bar.progress(30)
                        
                        others = []
                        for i, x in enumerate(uploaded_images[1:], start=2):
                            status_text.text(f"Processing image {i}/{total_imgs}...")
                            others.append(Image.open(x).convert("RGB"))
                            progress_bar.progress(30 + int((i-1)/total_imgs * 60))
                        
                        status_text.text("Generating PDF...")
                        pdf_buffer = BytesIO()
                        first.save(pdf_buffer, format="PDF", save_all=True, append_images=others)
                        pdf_buffer.seek(0)
                        
                        progress_bar.progress(100)
                        st.success("✅ PDF created successfully")
                        st.download_button(
                            "⬇️ Download PDF",
                            pdf_buffer.getvalue(),
                            file_name="Images_Combined.pdf",
                            mime="application/pdf"
                        )
                    except Exception as e:
                        st.error(f"❌ Error while creating PDF: {str(e)}")
    st.markdown('</div>', unsafe_allow_html=True)

# Footer
st.markdown("<hr>", unsafe_allow_html=True)
st.caption("© Tricks For Excel — Contact: WhatsApp 01554694554")
