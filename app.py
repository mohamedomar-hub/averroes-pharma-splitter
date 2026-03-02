# -*- coding: utf-8 -*-
"""
Streamlit App — Light Theme Refresh (UI Polished, English UI)
- Default: Light, clean palette (subtle gray background)
- Optional: Dark mode toggle in sidebar
- Tools: Split / Merge / Excel Processor / Images → PDF
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
import re
import os
import base64
import requests
from datetime import datetime

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle
from PIL import Image

# Optional animations
try:
    from streamlit_lottie import st_lottie  # type: ignore
except Exception:
    st_lottie = None


def load_lottie_url(url: str):
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

custom_css = """
<style>
:root {
  --primary: %(primary)s;
  --primary-2: %(primary2)s;
  --accent: %(accent)s;
  --app-bg: %(bg)s;
  --bg-elev: %(card)s;
  --bg-elev-2: %(card2)s;
  --text: %(text)s;
  --muted: %(muted)s;
  --border: %(border)s;
}

html, body { background: var(--app-bg) !important; }
section.main > div { padding-top: 10px; }
html, body, [class^="css"]  { font-family: 'Segoe UI', system-ui, -apple-system, Cairo, Tahoma, sans-serif; color: var(--text); }

/* Header */
.app-header {
  display:flex; align-items:center; justify-content:center; gap:16px; padding:18px 20px;
  background: linear-gradient(135deg, var(--bg-elev-2), #ffffff10);
  border: 1px solid var(--border); border-radius: 16px;
  box-shadow: 0 6px 18px rgba(17,24,39,.06);
}
.app-titlewrap { text-align:center; }
.app-title { margin:0; font-weight:800; letter-spacing:.2px; color: var(--text); }
.app-sub { margin:6px 0 0; color: var(--muted); font-size: 14.5px; }
.app-logo {
  width: 170px; height: auto; border-radius: 12px;
  box-shadow: 0 8px 22px rgba(0,0,0,.12);
}

/* Section card */
.card {
  border:1px solid var(--border); border-radius:16px; padding:18px 18px 10px; background: var(--bg-elev);
  margin: 10px 0 22px; box-shadow: 0 2px 12px rgba(17,24,39,.07);
}
.card h3, .card h2, .card h4 { margin-top:0; display:flex; align-items:center; gap:8px; color: var(--text); }

/* Hints - clear & highlighted */
.hint {
  display:block;
  background: var(--bg-elev-2);
  border-left: 4px solid var(--primary);
  color: var(--muted);
  font-size: 14.5px;
  padding: 10px 12px;
  border-radius: 10px;
  margin: -4px 0 10px 0;
}

/* Buttons */
.stButton > button {
  border-radius:12px; padding:10px 14px; font-weight:600; border:1px solid transparent;
  background: linear-gradient(135deg, var(--primary), var(--primary-2)); color:#fff; box-shadow: 0 4px 14px rgba(45,114,217,.25);
}
.stButton > button:hover { filter:brightness(1.06); transform: translateY(-1px); }
.stButton > button:active { transform: translateY(0); }

/* File uploader */
[data-testid="stFileUploader"] {
  background: var(--bg-elev-2);
  padding:12px; border-radius: 12px; border:1px dashed var(--border);
}

/* Download buttons */
.stDownloadButton > button {
  border-radius:10px; border:1px solid var(--border); background: var(--bg-elev-2); color: var(--text);
}

/* Dataframe wrapper */
.css-1m1b9qw, .stDataFrame { border-radius: 10px; overflow:hidden; border:1px solid var(--border); }

/* Divider */
hr { border: none; height: 1px; background: var(--border); margin: 14px 0; }

/* Force app background containers */
[data-testid="stAppViewContainer"] > .main { background: var(--app-bg); }
[data-testid="stHeader"] { background: var(--app-bg); border-bottom: 0; }
</style>
""" % C
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

def copy_cell_style(src_cell, dst_cell):
    """نسخ كل التنسيقات من خلية لأخرى"""
    if src_cell.has_style:
        try:
            if src_cell.font:
                dst_cell.font = Font(
                    name=src_cell.font.name,
                    size=src_cell.font.size,
                    bold=src_cell.font.bold,
                    italic=src_cell.font.italic,
                    vertAlign=src_cell.font.vertAlign,
                    underline=src_cell.font.underline,
                    strike=src_cell.font.strike,
                    color=src_cell.font.color
                )
            if src_cell.fill and src_cell.fill.fill_type:
                dst_cell.fill = PatternFill(
                    fill_type=src_cell.fill.fill_type,
                    start_color=src_cell.fill.start_color,
                    end_color=src_cell.fill.end_color
                )
            if src_cell.alignment:
                dst_cell.alignment = Alignment(
                    horizontal=src_cell.alignment.horizontal,
                    vertical=src_cell.alignment.vertical,
                    text_rotation=src_cell.alignment.text_rotation,
                    wrap_text=src_cell.alignment.wrap_text,
                    shrink_to_fit=src_cell.alignment.shrink_to_fit,
                    indent=src_cell.alignment.indent
                )
            if src_cell.border:
                dst_cell.border = Border(
                    left=src_cell.border.left,
                    right=src_cell.border.right,
                    top=src_cell.border.top,
                    bottom=src_cell.border.bottom,
                    diagonal=src_cell.border.diagonal,
                    diagonal_direction=src_cell.border.diagonal_direction,
                    outline=src_cell.border.outline,
                    vertical=src_cell.border.vertical,
                    horizontal=src_cell.border.horizontal
                )
            dst_cell.number_format = src_cell.number_format
        except Exception:
            pass

def copy_column_widths(src_ws, dst_ws):
    """نسخ عرض الأعمدة"""
    try:
        for col_letter in src_ws.column_dimensions:
            col_width = src_ws.column_dimensions[col_letter].width
            if col_width:
                dst_ws.column_dimensions[col_letter].width = col_width
    except Exception:
        pass

def load_bum_mapping():
    """تحميل ملف BUM من Google Sheets"""
    try:
        url = "https://docs.google.com/spreadsheets/d/1XQnQNDFHDKrWYn23ROAeFS2cELNbKurC/export?format=xlsx"
        response = requests.get(url)
        if response.status_code == 200:
            wb = load_workbook(filename=BytesIO(response.content))
            ws = wb.active
            
            # قراءة البيانات
            data = []
            headers = [cell.value for cell in ws[1]]
            mr_idx = None
            bum_idx = None
            
            for i, header in enumerate(headers):
                if header and "MR" in str(header):
                    mr_idx = i + 1
                elif header and "BUM" in str(header):
                    bum_idx = i + 1
            
            if mr_idx and bum_idx:
                for row in range(2, ws.max_row + 1):
                    mr_value = ws.cell(row, mr_idx).value
                    bum_value = ws.cell(row, bum_idx).value
                    if mr_value and bum_value:
                        data.append({
                            'MR': str(mr_value).strip(),
                            'BUM': str(bum_value).strip()
                        })
            
            return pd.DataFrame(data)
    except Exception as e:
        st.warning(f"⚠️ Could not load BUM mapping: {e}")
        return pd.DataFrame()


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
            split_option = st.radio(
                "Split method:",
                ["Split by Column Values", "Split Each Sheet into Separate File"],
                horizontal=True,
            )

            if st.button("🚀 Start"):
                with st.spinner("Processing..."):
                    if st_lottie and LOTTIE_SPLIT:
                        st_lottie(LOTTIE_SPLIT, height=110, key="lottie_split")

                    def clean_name(name: str) -> str:
                        name = str(name).strip()
                        invalid_chars = r'[\\/*?:\[\]\n<>:"\']'
                        cleaned = re.sub(invalid_chars, "_", name)
                        return cleaned[:30] if cleaned else "Sheet"

                    if file_ext == "csv":
                        unique_values = df[col_to_split].dropna().unique()
                        zip_buffer = BytesIO()
                        with ZipFile(zip_buffer, "w") as zip_file:
                            for value in unique_values:
                                filtered_df = df[df[col_to_split] == value]
                                csv_buffer = BytesIO()
                                filtered_df.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                                csv_buffer.seek(0)
                                zip_file.writestr(f"{clean_name(value)}.csv", csv_buffer.read())
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
                            
                            # Progress bar
                            progress_bar = st.progress(0)
                            status_text = st.empty()
                            
                            zip_buffer = BytesIO()
                            with ZipFile(zip_buffer, "w") as zip_file:
                                for i, value in enumerate(unique_values):
                                    status_text.text(f"Processing: {clean_name(value)}")
                                    progress_bar.progress((i + 1) / len(unique_values))
                                    
                                    new_wb = Workbook()
                                    default_ws = new_wb.active
                                    new_wb.remove(default_ws)
                                    new_ws = new_wb.create_sheet(title=clean_name(value))
                                    
                                    # Copy header with style
                                    for cell in ws[1]:
                                        dst = new_ws.cell(1, cell.column, cell.value)
                                        copy_cell_style(cell, dst)
                                    
                                    # Copy matching rows
                                    row_out = 2
                                    for row in ws.iter_rows(min_row=2):
                                        if row[col_idx - 1].value == value:
                                            for src in row:
                                                dst = new_ws.cell(row_out, src.column, src.value)
                                                copy_cell_style(src, dst)
                                            row_out += 1
                                    
                                    # Copy column widths
                                    copy_column_widths(ws, new_ws)
                                    
                                    fb = BytesIO()
                                    new_wb.save(fb)
                                    fb.seek(0)
                                    zip_file.writestr(f"{clean_name(value)}.xlsx", fb.read())
                            
                            status_text.empty()
                            progress_bar.empty()
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
                            with ZipFile(zip_buffer, "w") as zip_file:
                                for sheet_name in original_wb.sheetnames:
                                    new_wb = Workbook()
                                    default_ws = new_wb.active
                                    new_wb.remove(default_ws)
                                    new_ws = new_wb.create_sheet(title=sheet_name)
                                    src_ws = original_wb[sheet_name]
                                    
                                    # Copy all rows with styles
                                    for row in src_ws.iter_rows():
                                        for src_cell in row:
                                            dst = new_ws.cell(src_cell.row, src_cell.column, src_cell.value)
                                            copy_cell_style(src_cell, dst)
                                    
                                    copy_column_widths(src_ws, new_ws)
                                    
                                    fb = BytesIO()
                                    new_wb.save(fb)
                                    fb.seek(0)
                                    zip_file.writestr(f"{_safe_name(sheet_name)}.xlsx", fb.read())
                            zip_buffer.seek(0)
                            st.success("🎉 Split by sheets completed! ZIP is ready.")
                            st.download_button(
                                "⬇️ Download (ZIP)",
                                zip_buffer.getvalue(),
                                file_name=f"SplitBySheets_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip",
                                mime="application/zip"
                            )
        except Exception as e:
            st.error(f"❌ Error while splitting: {e}")
    st.markdown('</div>', unsafe_allow_html=True)


# ===================== Merge Card =====================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 🔁 Merge Excel/CSV Files")
    st.markdown('<span class="hint">Upload multiple files and they will be merged into one file with preserved formatting.</span>', unsafe_allow_html=True)

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
                with st.spinner("Merging..."):
                    if st_lottie and LOTTIE_MERGE:
                        st_lottie(LOTTIE_MERGE, height=100, key="lottie_merge")
                    try:
                        all_excel = all(f.name.lower().endswith('.xlsx') for f in merge_files)
                        
                        if all_excel:
                            # Create new workbook
                            merged_wb = Workbook()
                            merged_ws = merged_wb.active
                            merged_ws.title = "Merged_Data"
                            
                            current_row = 1
                            headers_copied = False
                            
                            # Progress bar
                            progress_bar = st.progress(0)
                            status_text = st.empty()
                            
                            for idx, file in enumerate(merge_files):
                                status_text.text(f"Processing: {file.name}")
                                progress_bar.progress((idx + 1) / len(merge_files))
                                
                                file_bytes = file.getvalue()
                                src_wb = load_workbook(filename=BytesIO(file_bytes), data_only=False)
                                src_ws = src_wb.active
                                
                                # Copy headers from first file only
                                if not headers_copied:
                                    for col, cell in enumerate(src_ws[1], start=1):
                                        dst_cell = merged_ws.cell(current_row, col, cell.value)
                                        copy_cell_style(cell, dst_cell)
                                    current_row += 1
                                    headers_copied = True
                                
                                # Copy data rows with their styles
                                for row in src_ws.iter_rows(min_row=2):
                                    for col, cell in enumerate(row, start=1):
                                        if cell.value is not None:  # Only copy non-empty cells
                                            dst_cell = merged_ws.cell(current_row, col, cell.value)
                                            copy_cell_style(cell, dst_cell)
                                    current_row += 1
                            
                            # Copy column widths from the first file
                            if merge_files:
                                first_file_bytes = merge_files[0].getvalue()
                                first_wb = load_workbook(filename=BytesIO(first_file_bytes), data_only=False)
                                first_ws = first_wb.active
                                copy_column_widths(first_ws, merged_ws)
                            
                            status_text.empty()
                            progress_bar.empty()
                            
                            # Save merged workbook
                            out = BytesIO()
                            merged_wb.save(out)
                            out.seek(0)
                            
                            st.success("✅ Merge completed with preserved formatting")
                            st.download_button(
                                "⬇️ Download merged file",
                                out.getvalue(),
                                file_name="Merged_Consolidated_Formatted.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            # Handle CSV files (simple merge without formatting)
                            all_dfs = []
                            for file in merge_files:
                                ext = file.name.split(".")[-1].lower()
                                df = pd.read_csv(file) if ext == "csv" else pd.read_excel(file)
                                all_dfs.append(df)
                            merged_df = pd.concat(all_dfs, ignore_index=True)
                            
                            out = BytesIO()
                            merged_df.to_excel(out, index=False, engine='openpyxl')
                            out.seek(0)
                            
                            st.success("✅ Merge completed")
                            st.download_button(
                                "⬇️ Download file",
                                out.getvalue(),
                                file_name="Merged_Consolidated.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    except Exception as e:
                        st.error(f"❌ Error while merging: {e}")
    st.markdown('</div>', unsafe_allow_html=True)


# ===================== Excel Processor Card =====================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 🧰 Excel Processor with BUM Mapping")
    st.markdown('<span class="hint">Process Excel file: Update BUM column (L4 Emp Name) based on MR name, and move CRM Interval Date to the beginning.</span>', unsafe_allow_html=True)

    proc_file = st.file_uploader(
        "📂 Upload Excel file to process (xlsx/xlsm)",
        type=["xlsx", "xlsm"],
        accept_multiple_files=False,
        key=f"processor_uploader_{st.session_state.clear_counter}",
    )

    # Load BUM mapping
    bum_df = load_bum_mapping()
    if not bum_df.empty:
        bum_dict = dict(zip(bum_df['MR'], bum_df['BUM']))
    else:
        bum_dict = {}
        # No warning message - just continue silently

    COLUMNS_TO_DELETE = [
        "Status","Status Date","Assigned To","Employees Count","Attendees Count",
        "Not Listed Invitees Count","Cost Per Person","Governorate","Accomodation Type",
        "CSR Enabled","Early Bird Due Date","Reservation Type","Meal Type","Meal Title",
        "Delivery Type","Start Date","End Date","Budget Date","Delivery Date","Invoice Date",
        "Actual Cost","Deduct","Net Amount","Other Professionals","Customers","Reps",
        "Items\\Brands","Created At","Request Professional Classifications",
        "Request Professional Ids","Segments","Accounts","Account Professionals","Category",
        "Type","Request Serial Number","Shared","Updated","L5 Emp Name",
        "Sponsored Company Name","Item Type","Item Brand","Promotional Code","Venue",
        "Restaurant","Business Type","Highlighted","Link Details",
    ]

    COLUMN_RENAME_MAP = { 
        "L1 Emp Name": "MR", 
        "L2 Emp Name": "DM", 
        "L3 Emp Name": "AM",
        "L4 Emp Name": "BUM"  # تم تغيير L4 Emp Name إلى BUM
    }

    FINAL_COLUMN_ORDER = [
        "CRM Interval Date",  # هننقلها للبداية
        "Tracking Number",
        "MR",
        "DM",
        "AM",
        "BUM",  # دي هي L4 Emp Name بعد التعديل
        "Line",
        "Activity",
        "Description",
        "Account Number",
        "Vendor",
        "Bank",
        "Cost",
        "Bricks",
        "Professionl Accounts",
        "Request Professionals",
        "Specialities",
        "Request Date",
    ]

    if proc_file:
        st.write("**File:**", proc_file.name)
        
        if st.button("⚙️ Start processing"):
            try:
                # Load the workbook
                wb = load_workbook(proc_file, data_only=False)
                ws = wb.active
                
                # Get headers
                header_row = 1
                headers = [ws.cell(header_row, col).value for col in range(1, ws.max_column + 1)]
                header_to_idx = {h: i+1 for i, h in enumerate(headers) if h is not None}
                
                # Create a new workbook for the result
                new_wb = Workbook()
                new_ws = new_wb.active
                new_ws.title = "Processed_Data"
                
                # Find important column indices
                mr_col_idx = None
                bum_col_idx = None
                crm_interval_idx = None
                
                for old_name, new_name in COLUMN_RENAME_MAP.items():
                    if old_name in header_to_idx:
                        if new_name == "MR":
                            mr_col_idx = header_to_idx[old_name]
                        elif new_name == "BUM":
                            bum_col_idx = header_to_idx[old_name]
                
                # Find CRM Interval Date column
                for col_name in headers:
                    if col_name and "CRM Interval Date" in str(col_name):
                        crm_interval_idx = header_to_idx[col_name]
                        break
                
                # Prepare final columns list
                final_cols_info = []
                
                # 1. CRM Interval Date (move to beginning)
                if crm_interval_idx:
                    final_cols_info.append({
                        'name': 'CRM Interval Date',
                        'type': 'existing',
                        'source_col': crm_interval_idx,
                        'original_name': headers[crm_interval_idx - 1]
                    })
                else:
                    # If not found, create empty column
                    final_cols_info.append({
                        'name': 'CRM Interval Date',
                        'type': 'new',
                        'value': ''
                    })
                
                # 2. Add remaining columns in order
                for col_name in FINAL_COLUMN_ORDER[1:]:  # Skip CRM Interval Date as we already added it
                    if col_name == "BUM" and bum_col_idx:
                        # BUM column - will be updated based on MR
                        final_cols_info.append({
                            'name': 'BUM',
                            'type': 'bum',
                            'source_col': bum_col_idx,
                            'mr_col': mr_col_idx
                        })
                    else:
                        # Find original column name considering renaming
                        found = False
                        for old_name, new_name in COLUMN_RENAME_MAP.items():
                            if col_name == new_name and old_name in header_to_idx:
                                final_cols_info.append({
                                    'name': col_name,
                                    'type': 'existing',
                                    'source_col': header_to_idx[old_name],
                                    'original_name': old_name
                                })
                                found = True
                                break
                        
                        if not found:
                            # Check if column exists with the same name
                            if col_name in header_to_idx:
                                final_cols_info.append({
                                    'name': col_name,
                                    'type': 'existing',
                                    'source_col': header_to_idx[col_name],
                                    'original_name': col_name
                                })
                
                # Copy headers with styles
                for col_idx, col_info in enumerate(final_cols_info, start=1):
                    # Write header name
                    new_ws.cell(1, col_idx, col_info['name'])
                    
                    # Copy header style if available
                    if col_info.get('source_col'):
                        src_cell = ws.cell(1, col_info['source_col'])
                        dst_cell = new_ws.cell(1, col_idx)
                        copy_cell_style(src_cell, dst_cell)
                
                # Process data rows
                for row_idx in range(2, ws.max_row + 1):
                    for col_idx, col_info in enumerate(final_cols_info, start=1):
                        if col_info['type'] == 'new':
                            # New empty column
                            new_ws.cell(row_idx, col_idx, '')
                        
                        elif col_info['type'] == 'bum':
                            # BUM column - update based on MR if mapping exists
                            src_cell = ws.cell(row_idx, col_info['source_col'])
                            
                            # Get MR value to find BUM from mapping
                            if col_info.get('mr_col'):
                                mr_value = ws.cell(row_idx, col_info['mr_col']).value
                                if mr_value and str(mr_value).strip() in bum_dict:
                                    # Update BUM value from mapping
                                    new_value = bum_dict[str(mr_value).strip()]
                                    dst_cell = new_ws.cell(row_idx, col_idx, new_value)
                                else:
                                    # Keep original value if no mapping found
                                    dst_cell = new_ws.cell(row_idx, col_idx, src_cell.value)
                            else:
                                dst_cell = new_ws.cell(row_idx, col_idx, src_cell.value)
                            
                            # Copy style from original cell
                            copy_cell_style(src_cell, dst_cell)
                        
                        else:
                            # Existing column - copy value and style
                            src_col = col_info['source_col']
                            src_cell = ws.cell(row_idx, src_col)
                            dst_cell = new_ws.cell(row_idx, col_idx, src_cell.value)
                            copy_cell_style(src_cell, dst_cell)
                
                # Copy column widths
                for col_idx, col_info in enumerate(final_cols_info, start=1):
                    if col_info.get('source_col'):
                        src_col_letter = get_column_letter(col_info['source_col'])
                        if src_col_letter in ws.column_dimensions:
                            width = ws.column_dimensions[src_col_letter].width
                            if width:
                                new_ws.column_dimensions[get_column_letter(col_idx)].width = width
                    else:
                        # Set default width for new columns
                        new_ws.column_dimensions[get_column_letter(col_idx)].width = 15
                
                # Save the processed workbook
                out_buf = BytesIO()
                new_wb.save(out_buf)
                out_buf.seek(0)
                
                st.success("✅ Processing completed: BUM column updated and CRM Interval Date moved to beginning")
                base = os.path.splitext(proc_file.name)[0]
                st.download_button(
                    "⬇️ Download processed file",
                    out_buf.getvalue(),
                    file_name=f"{_safe_name(base)}_processed_with_BUM.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"❌ Error while processing: {e}")
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
                with st.spinner("Creating PDF..."):
                    try:
                        # Progress bar for images
                        progress_bar = st.progress(0)
                        
                        # Open and convert images
                        images = []
                        for i, img_file in enumerate(uploaded_images):
                            img = Image.open(img_file)
                            if img.mode != 'RGB':
                                img = img.convert('RGB')
                            images.append(img)
                            progress_bar.progress((i + 1) / len(uploaded_images))
                        
                        # Create PDF
                        pdf_buffer = BytesIO()
                        images[0].save(
                            pdf_buffer,
                            format="PDF",
                            save_all=True,
                            append_images=images[1:],
                            quality=95
                        )
                        pdf_buffer.seek(0)
                        
                        progress_bar.empty()
                        st.success("✅ PDF created successfully")
                        st.download_button(
                            "⬇️ Download PDF",
                            pdf_buffer.getvalue(),
                            file_name="Images_Combined.pdf",
                            mime="application/pdf"
                        )
                    except Exception as e:
                        st.error(f"❌ Error while creating PDF: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

# Footer
st.markdown("<hr>", unsafe_allow_html=True)
st.caption("© Tricks For Excel — Contact: WhatsApp 01554694554")

