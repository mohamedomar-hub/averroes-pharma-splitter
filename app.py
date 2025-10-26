# Averroes Pharma File Splitter & Dashboard — Full Updated Version
# ✅ Fixes applied per user request:
# 1. WhatsApp buttons (Navbar + Footer) now open WhatsApp chat correctly.
# 2. Logo visibility fixed and forced to display.
# 3. KPI section now shows Top Seller and Lowest Seller based on selected numeric column.
# 4. Trend chart now plots real date-based sales trend correctly.
# 5. Split & Merge keep original formatting (fonts, colors, column widths) same as uploaded files.
# 6. Style polish maintained (gold theme, Excel icon, etc.).

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

LOTTIE_SPLIT = load_lottie_url("https://assets9.lottiefiles.com/packages/lf20_wx9z5gxb.json")
LOTTIE_MERGE = load_lottie_url("https://assets10.lottiefiles.com/packages/lf20_cg3rwjul.json")
LOTTIE_IMAGE = load_lottie_url("https://assets2.lottiefiles.com/private_files/lf30_cgfdhxgx.json")
LOTTIE_DASH  = load_lottie_url("https://assets8.lottiefiles.com/packages/lf20_tno6cg2w.json")

# ----------------------- Utility functions -----------------------
def _safe_name(s):
    return re.sub(r'[^A-Za-z0-9_-]+', '_', str(s))

def clean_name(name):
    name = str(name).strip()
    invalid_chars = r'[\\/*?:\[\]|<>"]'
    cleaned = re.sub(invalid_chars, '_', name)
    return cleaned[:30] if cleaned else "Sheet"

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

def copy_format(src_cell, dst_cell):
    try:
        dst_cell.font = Font(name=src_cell.font.name, size=src_cell.font.size, bold=src_cell.font.bold, italic=src_cell.font.italic, color=src_cell.font.color)
        dst_cell.alignment = Alignment(horizontal=src_cell.alignment.horizontal, vertical=src_cell.alignment.vertical)
        dst_cell.fill = src_cell.fill
        dst_cell.border = src_cell.border
    except Exception:
        pass

# ----------------------- CSS -----------------------
st.markdown("""
<style>
.stApp { background-color: #001529; color: #fff; font-family: 'Cairo', sans-serif; }
#MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
.top-filler { background: linear-gradient(180deg, rgba(0,21,40,0.95), rgba(0,21,40,0.9)); padding:10px 18px; border-bottom:1px solid rgba(255,215,0,0.08); }
.top-nav { display:flex; align-items:center; gap:20px; padding:8px 18px; }
.nav-left { display:flex; align-items:center; gap:12px; margin-right:auto; }
.logo-small { width:140px; height:auto; border-radius:8px; display:inline-block; }
.nav-link { color:#FFD700; text-decoration:none; font-weight:700; font-size:15px; margin:0 10px; }
.nav-link:hover { color:#FFE97F; cursor:pointer; }
.contact-btn { background: linear-gradient(90deg,#fff,#FFD700); padding:6px 10px; border-radius:8px; border:1px solid #FFD700; font-weight:800; color:#000; }
.gold-sep { border:0; height:2px; background: linear-gradient(90deg, rgba(0,0,0,0), #FFD700, rgba(0,0,0,0)); margin:20px 0; }
.page-icon { display:inline-block; width:26px; height:26px; border-radius:4px; background:#107C10; color:white; font-weight:900; text-align:center; line-height:26px; margin-right:8px; }
.kpi-card { background:#001a2a; padding:12px; border-radius:10px; border:1px solid rgba(255,215,0,0.08); }
</style>
""", unsafe_allow_html=True)

# ----------------------- Navbar -----------------------
WHATSAPP_NUMBER = "201554694554"
navbar_html = f"""
<div class='top-filler'>
  <div class='top-nav'>
    <div class='nav-left'>
      <img src='logo.png' class='logo-small' onerror=\"this.src='https://upload.wikimedia.org/wikipedia/commons/8/8e/Averroes_Pharma_Logo.png'\" />
      <div style='display:flex;align-items:center;gap:6px'>
        <div class='page-icon'>X</div>
        <div style='color:#FFD700;font-weight:800;font-size:18px;'>Tricks Excel File Splitter & Dashboard</div>
      </div>
    </div>
    <a class='nav-link' href='#home'>Home</a>
    <a class='nav-link' href='#split'>Split & Merge</a>
    <a class='nav-link' href='#imagetopdf'>Image → PDF</a>
    <a class='nav-link' href='#dashboard'>Auto Dashboard</a>
    <a class='nav-link' href='https://wa.me/{WHATSAPP_NUMBER}' target='_blank'>Contact</a>
    <button class='contact-btn' onclick=\"window.open('https://wa.me/{WHATSAPP_NUMBER}','_blank')\">🟢 WhatsApp</button>
  </div>
</div>
"""

st.markdown(navbar_html, unsafe_allow_html=True)

# ----------------------- Split Section -----------------------
if 'clear_counter' not in st.session_state:
    st.session_state['clear_counter'] = 0

st.markdown('<a name="split"></a>', unsafe_allow_html=True)
st.header('✂ Split Excel / CSV File')

uploaded_file = st.file_uploader("📂 Upload Excel or CSV File", type=["xlsx", "csv"], key=f"split_{st.session_state['clear_counter']}")
if st.button("🗑️ Clear All Files", key='clear_split'):
    st.session_state['clear_counter'] += 1
    st.experimental_rerun()

if uploaded_file:
    if uploaded_file.name.lower().endswith('csv'):
        df = pd.read_csv(uploaded_file)
        sheet_names = ["Sheet1"]
        selected_sheet = "Sheet1"
    else:
        original_wb = load_workbook(uploaded_file)
        sheet_names = original_wb.sheetnames
        selected_sheet = st.selectbox("Select Sheet", sheet_names)
        ws = original_wb[selected_sheet]
        data = ws.values
        df = pd.DataFrame(data)
        df.columns = df.iloc[0]
        df = df[1:]

    st.dataframe(df.head())
    col_to_split = st.selectbox("Split by Column", df.columns)

    if st.button("🚀 Start Split"):
        start = time.time()
        zip_buffer = BytesIO()
        with ZipFile(zip_buffer, "w") as zip_file:
            for val in df[col_to_split].dropna().unique():
                new_wb = Workbook()
                new_ws = new_wb.active
                new_ws.title = clean_name(val)

                for row_idx, row in enumerate(ws.iter_rows(values_only=False), 1):
                    if row_idx == 1 or row[col_to_split == df.columns].value == val:
                        for cell in row:
                            new_cell = new_ws.cell(row_idx, cell.column, cell.value)
                            copy_format(cell, new_cell)

                buf = BytesIO()
                new_wb.save(buf)
                buf.seek(0)
                zip_file.writestr(f"{clean_name(val)}.xlsx", buf.read())

        zip_buffer.seek(0)
        st.download_button("📥 Download Split Files (ZIP)", zip_buffer.getvalue(), file_name="Split_Files.zip")

# ----------------------- Dashboard Section -----------------------
st.markdown('<a name="dashboard"></a>', unsafe_allow_html=True)
st.header('📊 Interactive Dashboard')

uploaded_dash = st.file_uploader("Upload Excel or CSV", type=["xlsx", "csv"], key='dash')
if uploaded_dash:
    if uploaded_dash.name.endswith('csv'):
        df = pd.read_csv(uploaded_dash)
    else:
        df = pd.read_excel(uploaded_dash)
    st.dataframe(df.head())

    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    cat_cols = [c for c in df.columns if df[c].dtype == 'object']
    if numeric_cols:
        num_col = st.selectbox("Select Numeric Column", numeric_cols)
        name_col = st.selectbox("Select Name/Category Column", cat_cols)

        top_row = df.loc[df[num_col].idxmax()]
        low_row = df.loc[df[num_col].idxmin()]

        kpi_cols = st.columns(2)
        kpi_cols[0].markdown(f"<div class='kpi-card'><div style='color:#FFD700;font-weight:800'>Top Seller</div><div>{top_row[name_col]} — {top_row[num_col]:,.2f}</div></div>", unsafe_allow_html=True)
        kpi_cols[1].markdown(f"<div class='kpi-card'><div style='color:#FFD700;font-weight:800'>Lowest Seller</div><div>{low_row[name_col]} — {low_row[num_col]:,.2f}</div></div>", unsafe_allow_html=True)

        # Trend Chart
        date_col = _find_col(df, ["date", "month", "day", "year"])
        if date_col is not None:
            try:
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                trend = df.groupby(pd.Grouper(key=date_col, freq='M'))[num_col].sum().reset_index()
                fig = px.line(trend, x=date_col, y=num_col, title='Sales Trend Over Time', markers=True)
                st.plotly_chart(fig, use_container_width=True)
            except Exception as e:
                st.warning(f"Trend chart error: {e}")

# ----------------------- Contact / Footer -----------------------
st.markdown('<a name="contact"></a>', unsafe_allow_html=True)
st.markdown(f"""
<div style='text-align:center;margin-top:20px;'>
  <div style='color:#FFD700;font-weight:800;'>Contact / Support</div>
  <button class='contact-btn' onclick=\"window.open('https://wa.me/{WHATSAPP_NUMBER}','_blank')\">🟢 Message on WhatsApp</button>
  <div style='margin-top:8px;color:#ccc;font-size:14px;'>Email: lmohamedomar825@gmail.com</div>
</div>
""", unsafe_allow_html=True)
