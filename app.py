# -*- coding: utf-8 -*-
"""
Streamlit App — Light Theme Refresh (UI Polished, English UI)
- Default: Light, clean palette (subtle gray background)
- Optional: Dark mode toggle in sidebar
- Tools: Split / Merge / Excel Processor / AI Dashboard & Chat
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
import json
import plotly.express as px
import plotly.graph_objects as go

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle

# ===================== Google Sheet ID loader =====================
@st.cache_data(ttl=3600)
def load_online_doctor_ids():
    from openpyxl import load_workbook
    GDRIVE_SHEET_URL = ("https://docs.google.com/spreadsheets/d/1-u3cegWgrsoXvJYWVwQQRJbyYbdYtjIMDIifnalwHqo/export?format=xlsx")
    try:
        r = requests.get(GDRIVE_SHEET_URL)
        if r.status_code != 200:
            return {}, "⚠️ Cannot access Google Sheet."
        wb = load_workbook(BytesIO(r.content))
        ws = wb.active
        headers = []
        for col in range(1, ws.max_column + 1):
            val = ws.cell(1, col).value
            headers.append(str(val).strip().lower() if val else '')
        dcol = icol = None
        for i, h in enumerate(headers):
            if any(x in h for x in ['doctor', 'اسم', 'name', 'دكتور']):
                dcol = i + 1
            if any(x in h for x in ['id', 'رقم', 'بطاقة', 'national', 'identity']):
                icol = i + 1
        if not dcol or not icol:
            return {}, f"⚠️ Missing columns: {headers}"
        idd = {}
        total_processed = 0
        matched_entries = 0
        for row in ws.iter_rows(min_row=2, min_col=1, max_col=ws.max_column, values_only=False):
            total_processed += 1
            name_cell = row[dcol - 1]
            id_cell = row[icol - 1]
            n = name_cell.value
            v = id_cell.value
            if n is not None and v is not None:
                c = str(n).strip()
                s = str(v).strip()
                if c and s:
                    idd[c] = s
                    idd[c.lower()] = s
                    idd[c.replace(' ', '')] = s
                    matched_entries += 1
        return idd, f"✅ Loaded {matched_entries} doctor IDs from {total_processed} rows processed"
    except Exception as e:
        return {}, f"❌ Error: {e}"

# Optional animations
try:
    from streamlit_lottie import st_lottie
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
    page_title="Tricks For Excel — Split/Merge & AI Tools",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

LOTTIE_SPLIT = load_lottie_url("https://assets9.lottiefiles.com/packages/lf20_wx9z5gxb.json")
LOTTIE_MERGE = load_lottie_url("https://assets10.lottiefiles.com/packages/lf20_cg3rwjul.json")

# =================== THEME TOGGLE ===================
if 'ui_theme' not in st.session_state:
    st.session_state.ui_theme = 'Light'

with st.sidebar:
    st.markdown("### 🎨 Theme")
    st.session_state.ui_theme = st.radio("Choose theme", ["Light", "Dark"], index=0)

is_dark = st.session_state.ui_theme == 'Dark'

# ------------------ Custom CSS ------------------
colors_light = {
    'primary':  '#2563eb',
    'primary2': '#60a5fa',
    'accent':   '#22c55e',
    'bg':       '#f5f6fa',
    'card':     '#ffffff',
    'card2':    '#eef2ff',
    'text':     '#0f172a',
    'muted':    '#475569',
    'border':   '#e5e7eb',
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
html, body, [class^="css"] { font-family: 'Segoe UI', system-ui, -apple-system, Cairo, Tahoma, sans-serif; color: var(--text); }
.app-header {
  display:flex; align-items:center; justify-content:center; gap:16px; padding:18px 20px;
  background: linear-gradient(135deg, var(--bg-elev-2), #ffffff10);
  border: 1px solid var(--border); border-radius: 16px;
  box-shadow: 0 6px 18px rgba(17,24,39,.06);
}
.app-titlewrap { text-align:center; }
.app-title { margin:0; font-weight:800; letter-spacing:.2px; color: var(--text); }
.app-sub { margin:6px 0 0; color: var(--muted); font-size: 14.5px; }
.app-logo { width: 170px; height: auto; border-radius: 12px; box-shadow: 0 8px 22px rgba(0,0,0,.12); }
.card {
  border:1px solid var(--border); border-radius:16px; padding:18px 18px 10px; background: var(--bg-elev);
  margin: 10px 0 22px; box-shadow: 0 2px 12px rgba(17,24,39,.07);
}
.card h3, .card h2, .card h4 { margin-top:0; display:flex; align-items:center; gap:8px; color: var(--text); }
.hint {
  display:block; background: var(--bg-elev-2); border-left: 4px solid var(--primary);
  color: var(--muted); font-size: 14.5px; padding: 10px 12px;
  border-radius: 10px; margin: -4px 0 10px 0;
}
.chat-bubble-user {
  background: linear-gradient(135deg, var(--primary), var(--primary-2));
  color: white; padding: 10px 16px; border-radius: 18px 18px 4px 18px;
  margin: 6px 0 6px 40px; font-size: 14.5px; line-height: 1.5;
}
.chat-bubble-ai {
  background: var(--bg-elev-2); color: var(--text);
  padding: 10px 16px; border-radius: 18px 18px 18px 4px;
  margin: 6px 40px 6px 0; font-size: 14.5px; line-height: 1.5;
  border: 1px solid var(--border);
}
.chat-label { font-size: 12px; color: var(--muted); margin-bottom: 2px; }
.stButton > button {
  border-radius:12px; padding:10px 14px; font-weight:600; border:1px solid transparent;
  background: linear-gradient(135deg, var(--primary), var(--primary-2)); color:#fff;
  box-shadow: 0 4px 14px rgba(45,114,217,.25);
}
.stButton > button:hover { filter:brightness(1.06); transform: translateY(-1px); }
.stButton > button:active { transform: translateY(0); }
[data-testid="stFileUploader"] {
  background: var(--bg-elev-2); padding:12px; border-radius: 12px; border:1px dashed var(--border);
}
.stDownloadButton > button {
  border-radius:10px; border:1px solid var(--border); background: var(--bg-elev-2); color: var(--text);
}
.css-1m1b9qw, .stDataFrame { border-radius: 10px; overflow:hidden; border:1px solid var(--border); }
hr { border: none; height: 1px; background: var(--border); margin: 14px 0; }
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
    if src_cell.has_style:
        try:
            if src_cell.font:
                dst_cell.font = Font(
                    name=src_cell.font.name, size=src_cell.font.size,
                    bold=src_cell.font.bold, italic=src_cell.font.italic,
                    vertAlign=src_cell.font.vertAlign, underline=src_cell.font.underline,
                    strike=src_cell.font.strike, color=src_cell.font.color
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
                    left=src_cell.border.left, right=src_cell.border.right,
                    top=src_cell.border.top, bottom=src_cell.border.bottom,
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
    try:
        for col_letter in src_ws.column_dimensions:
            col_width = src_ws.column_dimensions[col_letter].width
            if col_width:
                dst_ws.column_dimensions[col_letter].width = col_width
    except Exception:
        pass

@st.cache_data(ttl=3600)
def load_bum_mapping():
    try:
        url = "https://docs.google.com/spreadsheets/d/1XQnQNDFHDKrWYn23ROAeFS2cELNbKurC/export?format=xlsx"
        response = requests.get(url)
        if response.status_code == 200:
            wb = load_workbook(filename=BytesIO(response.content))
            ws = wb.active
            data = []
            headers = [cell.value for cell in ws[1]]
            mr_idx = bum_idx = None
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
                        data.append({'MR': str(mr_value).strip(), 'BUM': str(bum_value).strip()})
            return pd.DataFrame(data)
    except Exception as e:
        st.warning(f"⚠️ Could not load BUM mapping: {e}")
        return pd.DataFrame()

def _is_match(cell_value, target_value):
    if cell_value is None and target_value is None:
        return True
    if cell_value is None or target_value is None:
        return False
    str_cell = str(cell_value).strip()
    str_target = str(target_value).strip()
    if str_cell == str_target:
        return True
    try:
        if float(cell_value) == float(target_value):
            return True
    except (ValueError, TypeError):
        pass
    if str_cell.lower() == str_target.lower():
        return True
    return False

# ===================== AI Chat Function =====================
def ask_claude(api_key: str, messages: list) -> str:
    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json={
                "model": "claude-sonnet-4-6",
                "max_tokens": 1024,
                "messages": messages,
            },
            timeout=30,
        )
        if resp.status_code == 200:
            return resp.json()["content"][0]["text"]
        else:
            return f"❌ API Error {resp.status_code}: {resp.text}"
    except Exception as e:
        return f"❌ Connection Error: {e}"

def df_to_context(df: pd.DataFrame, max_rows: int = 100) -> str:
    """تحويل الـ DataFrame لنص يفهمه الـ AI"""
    context = f"Dataset has {len(df)} rows and {len(df.columns)} columns.\n"
    context += f"Columns: {', '.join(df.columns.tolist())}\n\n"
    
    # إحصائيات سريعة للأعمدة الرقمية
    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    if numeric_cols:
        context += "Numeric columns summary:\n"
        for col in numeric_cols[:5]:
            context += f"  - {col}: min={df[col].min()}, max={df[col].max()}, avg={df[col].mean():.2f}, sum={df[col].sum():.2f}\n"
        context += "\n"
    
    # عينة من البيانات
    sample = df.head(max_rows)
    context += f"Data sample (first {min(max_rows, len(df))} rows):\n"
    context += sample.to_string(index=False)
    return context

def auto_generate_charts(df: pd.DataFrame, is_dark: bool):
    """توليد charts تلقائية بناءً على نوع البيانات"""
    plotly_theme = "plotly_dark" if is_dark else "plotly_white"
    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    categorical_cols = df.select_dtypes(include='object').columns.tolist()
    
    charts_made = 0
    
    # Chart 1: Bar chart لأول عمود categorical + أول عمود رقمي
    if categorical_cols and numeric_cols and charts_made < 3:
        cat_col = categorical_cols[0]
        num_col = numeric_cols[0]
        top_data = df.groupby(cat_col)[num_col].sum().nlargest(15).reset_index()
        fig = px.bar(
            top_data, x=cat_col, y=num_col,
            title=f"📊 {num_col} by {cat_col} (Top 15)",
            template=plotly_theme,
            color=num_col,
            color_continuous_scale="Blues"
        )
        fig.update_layout(showlegend=False, height=400)
        st.plotly_chart(fig, use_container_width=True)
        charts_made += 1

    # Chart 2: Pie chart لو في عمود categorical تاني
    if len(categorical_cols) > 1 and numeric_cols and charts_made < 3:
        cat_col2 = categorical_cols[1]
        num_col = numeric_cols[0]
        pie_data = df.groupby(cat_col2)[num_col].sum().nlargest(10).reset_index()
        fig2 = px.pie(
            pie_data, names=cat_col2, values=num_col,
            title=f"🥧 {num_col} Distribution by {cat_col2}",
            template=plotly_theme,
            hole=0.4
        )
        fig2.update_layout(height=400)
        st.plotly_chart(fig2, use_container_width=True)
        charts_made += 1

    # Chart 3: Line chart لو في أكتر من عمود رقمي
    if len(numeric_cols) >= 2 and charts_made < 3:
        fig3 = px.line(
            df.head(50), y=numeric_cols[:3],
            title=f"📈 Trend: {', '.join(numeric_cols[:3])}",
            template=plotly_theme
        )
        fig3.update_layout(height=400)
        st.plotly_chart(fig3, use_container_width=True)
        charts_made += 1

    if charts_made == 0:
        st.info("ℹ️ Upload a file with numeric columns to generate charts automatically.")

# ------------------ Header ------------------
logo_b64 = get_image_as_base64("logo.png")
header_html = f"""
<div class="app-header">
  {('<img class="app-logo" src="data:image/png;base64,' + logo_b64 + '" alt="Logo" />') if logo_b64 else ''}
  <div class="app-titlewrap">
    <h2 class="app-title">Tricks For Excel</h2>
    <p class="app-sub">Quick tools for Excel • Split • Merge • Processor • AI Dashboard</p>
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
        if uploaded_file.size > 50 * 1024 * 1024:
            st.error("❌ File too large! Max allowed: 50MB")
            st.stop()

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
            df.columns = df.columns.astype(str)
            col_to_split = st.selectbox("Select column to split by", df.columns)

            # Split Preview
            if col_to_split:
                preview_data = df[col_to_split].value_counts().reset_index()
                preview_data.columns = [col_to_split, 'Row Count']
                with st.expander(f"👁️ Preview: {len(preview_data)} unique values will create {len(preview_data)} files"):
                    st.dataframe(preview_data, use_container_width=True)

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
                        st.download_button("⬇️ Download (ZIP)", zip_buffer.getvalue(),
                            file_name=f"Split_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip",
                            mime="application/zip")
                    else:
                        ws = original_wb[selected_sheet]
                        if split_option == "Split by Column Values":
                            col_idx = df.columns.get_loc(col_to_split) + 1
                            unique_values = df[col_to_split].dropna().unique()
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
                                    for cell in ws[1]:
                                        dst = new_ws.cell(1, cell.column, cell.value)
                                        copy_cell_style(cell, dst)
                                    row_out = 2
                                    for row in ws.iter_rows(min_row=2):
                                        cell_value = row[col_idx - 1].value
                                        if _is_match(cell_value, value):
                                            for src in row:
                                                dst = new_ws.cell(row_out, src.column, src.value)
                                                copy_cell_style(src, dst)
                                            row_out += 1
                                    copy_column_widths(ws, new_ws)
                                    fb = BytesIO()
                                    new_wb.save(fb)
                                    fb.seek(0)
                                    zip_file.writestr(f"{clean_name(value)}.xlsx", fb.read())
                            status_text.empty()
                            progress_bar.empty()
                            zip_buffer.seek(0)
                            st.success("🎉 Split completed! ZIP is ready.")
                            st.download_button("⬇️ Download (ZIP)", zip_buffer.getvalue(),
                                file_name=f"Split_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip",
                                mime="application/zip")
                        else:
                            zip_buffer = BytesIO()
                            with ZipFile(zip_buffer, "w") as zip_file:
                                for sheet_name in original_wb.sheetnames:
                                    new_wb = Workbook()
                                    default_ws = new_wb.active
                                    new_wb.remove(default_ws)
                                    new_ws = new_wb.create_sheet(title=sheet_name)
                                    src_ws = original_wb[sheet_name]
                                    for row in src_ws.iter_rows():
                                        for src_cell in row:
                                            dst = new_ws.cell(src_cell.row, src_cell.column, src_cell.value)
                                            copy_cell_style(src_cell, dst)
                                    for merged_range in src_ws.merged_cells.ranges:
                                        new_ws.merge_cells(str(merged_range))
                                    copy_column_widths(src_ws, new_ws)
                                    fb = BytesIO()
                                    new_wb.save(fb)
                                    fb.seek(0)
                                    zip_file.writestr(f"{_safe_name(sheet_name)}.xlsx", fb.read())
                            zip_buffer.seek(0)
                            st.success("🎉 Split by sheets completed! ZIP is ready.")
                            st.download_button("⬇️ Download (ZIP)", zip_buffer.getvalue(),
                                file_name=f"SplitBySheets_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip",
                                mime="application/zip")
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
                            merged_wb = Workbook()
                            merged_ws = merged_wb.active
                            merged_ws.title = "Merged_Data"
                            current_row = 1
                            headers_copied = False
                            progress_bar = st.progress(0)
                            status_text = st.empty()
                            for idx, file in enumerate(merge_files):
                                status_text.text(f"Processing: {file.name}")
                                progress_bar.progress((idx + 1) / len(merge_files))
                                file_bytes = file.getvalue()
                                src_wb = load_workbook(filename=BytesIO(file_bytes), data_only=False)
                                src_ws = src_wb.active
                                if not headers_copied:
                                    for col, cell in enumerate(src_ws[1], start=1):
                                        dst_cell = merged_ws.cell(current_row, col, cell.value)
                                        copy_cell_style(cell, dst_cell)
                                    current_row += 1
                                    headers_copied = True
                                for row in src_ws.iter_rows(min_row=2):
                                    for col, cell in enumerate(row, start=1):
                                        if cell.value is not None:
                                            dst_cell = merged_ws.cell(current_row, col, cell.value)
                                            copy_cell_style(cell, dst_cell)
                                    current_row += 1
                            if merge_files:
                                first_file_bytes = merge_files[0].getvalue()
                                first_wb = load_workbook(filename=BytesIO(first_file_bytes), data_only=False)
                                first_ws = first_wb.active
                                copy_column_widths(first_ws, merged_ws)
                            status_text.empty()
                            progress_bar.empty()
                            out = BytesIO()
                            merged_wb.save(out)
                            out.seek(0)
                            st.success("✅ Merge completed with preserved formatting")
                            st.download_button("⬇️ Download merged file", out.getvalue(),
                                file_name="Merged_Consolidated_Formatted.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        else:
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
                            st.download_button("⬇️ Download file", out.getvalue(),
                                file_name="Merged_Consolidated.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    except Exception as e:
                        st.error(f"❌ Error while merging: {e}")
    st.markdown('</div>', unsafe_allow_html=True)


# ===================== Excel Processor Card =====================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 🧰 Excel Processor Service")
    st.markdown('<span class="hint">Process Excel file: Update BUM column, add ID Numbers, and move CRM Interval Date to the beginning.</span>', unsafe_allow_html=True)

    id_dict, id_message = load_online_doctor_ids()
    st.info(id_message)

    proc_file = st.file_uploader(
        "📂 Upload Excel file to process (xlsx/xlsm)",
        type=["xlsx", "xlsm"],
        accept_multiple_files=False,
        key=f"processor_uploader_{st.session_state.clear_counter}",
    )

    bum_df = load_bum_mapping()
    if not bum_df.empty:
        bum_dict = dict(zip(bum_df['MR'], bum_df['BUM']))
    else:
        bum_dict = {}

    COLUMN_RENAME_MAP = {
        "L1 Emp Name": "MR",
        "L2 Emp Name": "DM",
        "L3 Emp Name": "AM",
        "L4 Emp Name": "BUM"
    }

    FINAL_COLUMN_ORDER = [
        "CRM Interval Date", "Tracking Number", "MR", "DM", "AM", "BUM",
        "Line", "Activity", "Description", "Account Number", "Vendor", "Bank",
        "Cost", "Bricks", "Professionl Accounts", "Request Professionals",
        "Specialities", "Request Date",
    ]

    if proc_file:
        st.write("**File:**", proc_file.name)
        if id_dict:
            st.info("📊 Preview of uploaded file (first 5 rows):")
            try:
                sample_df = pd.read_excel(proc_file, nrows=5)
                st.dataframe(sample_df, use_container_width=True)
            except:
                pass

        if st.button("⚙️ Start processing"):
            try:
                wb = load_workbook(proc_file, data_only=False)
                ws = wb.active
                headers = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
                header_to_idx = {h: i+1 for i, h in enumerate(headers) if h is not None}
                new_wb = Workbook()
                new_ws = new_wb.active
                new_ws.title = "Processed_Data"
                mr_col_idx = bum_col_idx = crm_interval_idx = doctor_name_col_idx = None

                for old_name, new_name in COLUMN_RENAME_MAP.items():
                    if old_name in header_to_idx:
                        if new_name == "MR":
                            mr_col_idx = header_to_idx[old_name]
                        elif new_name == "BUM":
                            bum_col_idx = header_to_idx[old_name]

                st.write("**Searching for doctor name column:**")
                for col_name in headers:
                    if col_name:
                        col_name_str = str(col_name).strip()
                        if col_name_str == "Professionl Accounts":
                            doctor_name_col_idx = header_to_idx[col_name]
                            st.write(f"✅ Found: '{col_name}' at position {doctor_name_col_idx}")
                            break
                        elif any(keyword in col_name_str.lower() for keyword in ['professionl', 'professional', 'account', 'doctor', 'name']):
                            doctor_name_col_idx = header_to_idx[col_name]
                            st.write(f"⚠️ Found potential: '{col_name}' at position {doctor_name_col_idx}")

                if not doctor_name_col_idx:
                    st.warning("⚠️ Could not find 'Professionl Accounts' column. ID Numbers will not be added.")

                for col_name in headers:
                    if col_name and "CRM Interval Date" in str(col_name):
                        crm_interval_idx = header_to_idx[col_name]
                        break

                final_cols_info = []
                if crm_interval_idx:
                    final_cols_info.append({'name': 'CRM Interval Date', 'type': 'existing',
                        'source_col': crm_interval_idx, 'original_name': headers[crm_interval_idx - 1]})
                else:
                    final_cols_info.append({'name': 'CRM Interval Date', 'type': 'new', 'value': ''})

                for col_name in FINAL_COLUMN_ORDER[1:]:
                    if col_name == "BUM" and bum_col_idx:
                        final_cols_info.append({'name': 'BUM', 'type': 'bum',
                            'source_col': bum_col_idx, 'mr_col': mr_col_idx})
                    else:
                        found = False
                        for old_name, new_name in COLUMN_RENAME_MAP.items():
                            if col_name == new_name and old_name in header_to_idx:
                                final_cols_info.append({'name': col_name, 'type': 'existing',
                                    'source_col': header_to_idx[old_name], 'original_name': old_name})
                                found = True
                                break
                        if not found and col_name in header_to_idx:
                            final_cols_info.append({'name': col_name, 'type': 'existing',
                                'source_col': header_to_idx[col_name], 'original_name': col_name})

                if id_dict and doctor_name_col_idx:
                    final_cols_info.append({'name': 'ID Number', 'type': 'id_number',
                        'doctor_col': doctor_name_col_idx})
                    st.info(f"✅ ID Number column will be added. Found {len(id_dict)//3} doctor IDs.")

                for col_idx, col_info in enumerate(final_cols_info, start=1):
                    new_ws.cell(1, col_idx, col_info['name'])
                    if col_info.get('source_col'):
                        src_cell = ws.cell(1, col_info['source_col'])
                        dst_cell = new_ws.cell(1, col_idx)
                        copy_cell_style(src_cell, dst_cell)

                matched_count = 0
                unmatched_doctors = []

                for row_idx in range(2, ws.max_row + 1):
                    for col_idx, col_info in enumerate(final_cols_info, start=1):
                        if col_info['type'] == 'new':
                            new_ws.cell(row_idx, col_idx, '')
                        elif col_info['type'] == 'id_number':
                            if col_info.get('doctor_col'):
                                doctor_name = ws.cell(row_idx, col_info['doctor_col']).value
                                if doctor_name:
                                    clean_name_val = str(doctor_name).strip()
                                    found_id = (id_dict.get(clean_name_val) or
                                                id_dict.get(clean_name_val.lower()) or
                                                id_dict.get(clean_name_val.replace(" ", "")))
                                    if found_id:
                                        new_ws.cell(row_idx, col_idx, found_id)
                                        matched_count += 1
                                    else:
                                        new_ws.cell(row_idx, col_idx, '')
                                        if clean_name_val not in unmatched_doctors:
                                            unmatched_doctors.append(clean_name_val)
                                else:
                                    new_ws.cell(row_idx, col_idx, '')
                            else:
                                new_ws.cell(row_idx, col_idx, '')
                        elif col_info['type'] == 'bum':
                            src_cell = ws.cell(row_idx, col_info['source_col'])
                            if col_info.get('mr_col'):
                                mr_value = ws.cell(row_idx, col_info['mr_col']).value
                                if mr_value and str(mr_value).strip() in bum_dict:
                                    dst_cell = new_ws.cell(row_idx, col_idx, bum_dict[str(mr_value).strip()])
                                else:
                                    dst_cell = new_ws.cell(row_idx, col_idx, src_cell.value)
                            else:
                                dst_cell = new_ws.cell(row_idx, col_idx, src_cell.value)
                            copy_cell_style(src_cell, dst_cell)
                        else:
                            src_cell = ws.cell(row_idx, col_info['source_col'])
                            dst_cell = new_ws.cell(row_idx, col_idx, src_cell.value)
                            copy_cell_style(src_cell, dst_cell)

                if id_dict and doctor_name_col_idx:
                    total_doctors = ws.max_row - 1
                    if matched_count > 0:
                        st.success(f"✅ Matched {matched_count} out of {total_doctors} doctors")
                    else:
                        st.warning(f"⚠️ No matches found! Sample names:")
                        for name in unmatched_doctors[:10]:
                            st.write(f"- '{name}'")

                for col_idx, col_info in enumerate(final_cols_info, start=1):
                    if col_info.get('source_col'):
                        src_col_letter = get_column_letter(col_info['source_col'])
                        if src_col_letter in ws.column_dimensions:
                            width = ws.column_dimensions[src_col_letter].width
                            if width:
                                new_ws.column_dimensions[get_column_letter(col_idx)].width = width
                    else:
                        new_ws.column_dimensions[get_column_letter(col_idx)].width = 15

                out_buf = BytesIO()
                new_wb.save(out_buf)
                out_buf.seek(0)

                success_msg = "✅ Processing completed: "
                if bum_dict:
                    success_msg += "BUM updated, "
                if id_dict and doctor_name_col_idx:
                    success_msg += f"ID Numbers added ({matched_count} matched), "
                success_msg += "CRM Interval Date moved to beginning"
                st.success(success_msg)

                base = os.path.splitext(proc_file.name)[0]
                st.download_button("⬇️ Download processed file", out_buf.getvalue(),
                    file_name=f"{_safe_name(base)}_processed.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"❌ Error while processing: {e}")
    st.markdown('</div>', unsafe_allow_html=True)


# ===================== AI Dashboard & Chat Card =====================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 🤖 AI Dashboard & Chat")
    st.markdown('<span class="hint">Upload any Excel/CSV file — get automatic charts + chat with your data using AI. Ask questions in Arabic or English!</span>', unsafe_allow_html=True)

    # API Key input
    api_key_input = st.secrets["ANTHROPIC_API_KEY"]
        "🔑 Anthropic API Key",
        type="password",
        placeholder="sk-ant-api03-...",
        help="Your key is never stored or sent anywhere except Anthropic's API",
        key="anthropic_api_key"
    )

    ai_file = st.file_uploader(
        "📂 Upload Excel or CSV for AI analysis",
        type=["xlsx", "csv"],
        accept_multiple_files=False,
        key=f"ai_uploader_{st.session_state.clear_counter}",
    )

    if ai_file:
        if ai_file.size > 50 * 1024 * 1024:
            st.error("❌ File too large! Max allowed: 50MB")
        else:
            try:
                # قراءة الملف
                file_ext = ai_file.name.split(".")[-1].lower()
                if file_ext == "csv":
                    ai_df = pd.read_csv(ai_file)
                else:
                    ai_df = pd.read_excel(ai_file)

                # إحصائيات سريعة
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("📋 Rows", f"{len(ai_df):,}")
                col2.metric("📊 Columns", len(ai_df.columns))
                col3.metric("🔢 Numeric Cols", len(ai_df.select_dtypes(include='number').columns))
                col4.metric("📝 Text Cols", len(ai_df.select_dtypes(include='object').columns))

                st.markdown("---")

                # Auto Charts
                st.markdown("#### 📊 Auto-Generated Charts")
                auto_generate_charts(ai_df, is_dark)

                st.markdown("---")

                # Data Preview
                with st.expander("👁️ Data Preview"):
                    st.dataframe(ai_df.head(50), use_container_width=True)

                st.markdown("---")

                # AI Chat Section
                st.markdown("#### 💬 Chat with Your Data")

                if not api_key_input:
                    st.warning("⚠️ Please enter your Anthropic API Key above to use the AI chat.")
                else:
                    # Initialize chat history
                    if 'ai_chat_history' not in st.session_state:
                        st.session_state.ai_chat_history = []
                    if 'ai_df_context' not in st.session_state:
                        st.session_state.ai_df_context = ""

                    # تحديث الـ context لو الملف اتغيّر
                    current_context = df_to_context(ai_df)
                    if st.session_state.ai_df_context != current_context:
                        st.session_state.ai_df_context = current_context
                        st.session_state.ai_chat_history = []

                    # عرض المحادثة
                    chat_container = st.container()
                    with chat_container:
                        for msg in st.session_state.ai_chat_history:
                            if msg["role"] == "user":
                                st.markdown(f'<div class="chat-label">You</div><div class="chat-bubble-user">{msg["content"]}</div>', unsafe_allow_html=True)
                            else:
                                st.markdown(f'<div class="chat-label">🤖 AI</div><div class="chat-bubble-ai">{msg["content"]}</div>', unsafe_allow_html=True)

                    # Quick Suggestions
                    st.markdown("**💡 Quick questions:**")
                    suggestions = [
                        "ما هي أهم الإحصائيات في هذه البيانات؟",
                        "What are the top 5 values?",
                        "هل في قيم مكررة أو بيانات ناقصة؟",
                        "Summarize the data in 3 points",
                    ]
                    cols = st.columns(2)
                    for i, suggestion in enumerate(suggestions):
                        if cols[i % 2].button(suggestion, key=f"suggest_{i}"):
                            # إرسال السؤال المقترح
                            system_prompt = f"""You are a helpful data analyst assistant. 
The user has uploaded a dataset with the following information:

{st.session_state.ai_df_context}

Answer questions about this data clearly and concisely.
If the user writes in Arabic, respond in Arabic.
If the user writes in English, respond in English.
Use numbers and specific examples from the data when possible."""

                            messages_to_send = [{"role": "user", "content": system_prompt}]
                            for h in st.session_state.ai_chat_history:
                                messages_to_send.append(h)
                            messages_to_send.append({"role": "user", "content": suggestion})

                            with st.spinner("🤖 Thinking..."):
                                response = ask_claude(api_key_input, [
                                    {"role": "user", "content": system_prompt + "\n\nUser question: " + suggestion}
                                ])

                            st.session_state.ai_chat_history.append({"role": "user", "content": suggestion})
                            st.session_state.ai_chat_history.append({"role": "assistant", "content": response})
                            st.rerun()

                    # Chat Input
                    user_question = st.text_input(
                        "💬 Ask anything about your data...",
                        placeholder="مثال: ما هو متوسط التكلفة؟ / What is the total by region?",
                        key=f"ai_chat_input_{len(st.session_state.ai_chat_history)}"
                    )

                    c1, c2 = st.columns([1, 4])
                    with c1:
                        send_btn = st.button("📤 Send", key="send_ai_msg")
                    with c2:
                        if st.button("🗑️ Clear Chat", key="clear_ai_chat"):
                            st.session_state.ai_chat_history = []
                            st.rerun()

                    if send_btn and user_question.strip():
                        system_prompt = f"""You are a helpful data analyst assistant.
The user has uploaded a dataset with the following information:

{st.session_state.ai_df_context}

Answer questions about this data clearly and concisely.
If the user writes in Arabic, respond in Arabic.
If the user writes in English, respond in English.
Use numbers and specific examples from the data when possible."""

                        # بناء المحادثة الكاملة
                        full_messages = []
                        if st.session_state.ai_chat_history:
                            # أول رسالة تحتوي الـ context
                            full_messages.append({
                                "role": "user",
                                "content": system_prompt + f"\n\nUser question: {st.session_state.ai_chat_history[0]['content']}"
                            })
                            if len(st.session_state.ai_chat_history) > 1:
                                full_messages.append(st.session_state.ai_chat_history[1])
                            # باقي المحادثة
                            for h in st.session_state.ai_chat_history[2:]:
                                full_messages.append(h)
                        
                        full_messages.append({"role": "user", "content": user_question})

                        with st.spinner("🤖 Thinking..."):
                            response = ask_claude(api_key_input, full_messages if full_messages else [
                                {"role": "user", "content": system_prompt + f"\n\nUser question: {user_question}"}
                            ])

                        st.session_state.ai_chat_history.append({"role": "user", "content": user_question})
                        st.session_state.ai_chat_history.append({"role": "assistant", "content": response})
                        st.rerun()

            except Exception as e:
                st.error(f"❌ Error reading file: {e}")

    st.markdown('</div>', unsafe_allow_html=True)


# Footer
st.markdown("<hr>", unsafe_allow_html=True)
st.caption("© Tricks For Excel — Contact: WhatsApp 01554694554")
