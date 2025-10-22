# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
import re
import os
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl import load_workbook, Workbook
# ====== Dashboard & Reporting ======
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image as RLImage, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from PIL import Image
from sklearn.linear_model import LinearRegression
import numpy as np
# === Lottie animation support ===
from streamlit_lottie import st_lottie
import requests
import json
import time  # ensure time is available for timers

def load_lottie_url(url: str):
    """تحميل Lottie JSON من رابط خارجي"""
    try:
        r = requests.get(url, timeout=8)
        if r.status_code == 200:
            return r.json()
    except Exception:
        return None
    return None

# Initialize session state
if 'clear_counter' not in st.session_state:
    st.session_state.clear_counter = 0

# ------------------ Page Setup ------------------
st.set_page_config(
    page_title="Averroes Pharma File Splitter & Dashboard",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

LOTTIE_SPLIT = load_lottie_url("https://assets9.lottiefiles.com/packages/lf20_wx9z5gxb.json")   # split
LOTTIE_MERGE = load_lottie_url("https://assets10.lottiefiles.com/packages/lf20_cg3rwjul.json")  # merge
LOTTIE_IMAGE = load_lottie_url("https://assets2.lottiefiles.com/private_files/lf30_cgfdhxgx.json")  # image/pdf
LOTTIE_DASH  = load_lottie_url("https://assets8.lottiefiles.com/packages/lf20_tno6cg2w.json")   # dashboard
LOTTIE_PDF   = load_lottie_url("https://assets1.lottiefiles.com/packages/lf20_zyu0ct3i.json")   # dashboard PDF

# Hide default Streamlit elements
hide_default = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
"""
st.markdown(hide_default, unsafe_allow_html=True)

# ------------------ Custom CSS ------------------
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
        gap: 20px;
        padding: 10px 30px;
        background-color: #001a33;
        border-bottom: 1px solid #FFD700;
        font-size: 18px;
        color: white;
    }
    .top-nav a {
        color: #FFD700;
        text-decoration: none;
        font-weight: bold;
        padding: 5px 10px;
        border-radius: 8px;
        transition: all 0.3s ease;
    }
    .top-nav a:hover {
        background-color: #FFD700;
        color: black;
    }
    label, .stSelectbox label, .stFileUploader label {
        color: #FFD700 !important;
        font-size: 18px !important;
        font-weight: bold !important;
    }
    .stButton>button, .stDownloadButton>button {
        background-color: #FFD700 !important;
        color: black !important;
        font-weight: bold !important;
        font-size: 18px !important;
        border-radius: 12px !important;
        padding: 12px 24px !important;
        border: none !important;
        box-shadow: 0 4px 8px rgba(0,0,0,0.3) !important;
        transition: all 0.3s ease !important;
        margin-top: 10px !important;
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        background-color: #FFC107 !important;
        transform: scale(1.08);
        box-shadow: 0 6px 12px rgba(0,0,0,0.4) !important;
    }
    .kpi-card {
        padding: 16px;
        border-radius: 12px;
        color: white;
        text-align: center;
        box-shadow: 0 6px 18px rgba(0,0,0,0.3);
        font-weight: 700;
        margin: 8px;
    }
    .kpi-title { font-size: 14px; opacity: 0.9; }
    .kpi-value { font-size: 22px; margin-top:6px; }
    hr.divider {
        border: 1px solid #FFD700;
        opacity: 0.6;
        margin: 30px 0;
    }
    hr.divider-dashed {
        border: 1px dashed #FFD700;
        opacity: 0.7;
        margin: 25px 0;
    }
    .stDataFrame {
        box-shadow: 0 4px 10px rgba(0,0,0,0.2);
        border-radius: 12px;
        overflow: hidden;
        margin: 10px 0;
    }
    .stFileUploader {
        border: 2px dashed #FFD700;
        border-radius: 10px;
        padding: 15px;
        background-color: rgba(255, 215, 0, 0.1);
    }
    .guide-title {
        color: #FFD700;
        font-weight: bold;
        font-size: 20px;
    }
    .chart-card {
        background-color: #00264d;
        border: 1px solid #FFD700;
        border-radius: 12px;
        padding: 12px;
        margin: 10px 0;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3);
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 20px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        padding-top: 10px;
        padding-bottom: 10px;
        font-size: 18px;
        font-weight: bold;
        border-radius: 10px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #FFD700 !important;
        color: black !important;
        border: 2px solid #FFC107 !important;
    }
    .stTabs [aria-selected="false"] {
        background-color: #003366 !important;
        color: #FFD700 !important;
        border: 1px solid #FFD700 !important;
    }
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ------------------ Helper Functions ------------------
def display_uploaded_files(file_list, file_type="Excel/CSV"):
    if file_list:
        st.markdown("### 📁 Uploaded Files:")
        for i, f in enumerate(file_list):
            st.markdown(
                f"<div style='background:#003366; color:white; padding:4px 8px; border-radius:4px; margin:2px 0; display:inline-block;'>"
                f"{i+1}. {f.name} ({f.size//1024} KB)</div>",
                unsafe_allow_html=True
            )

def _safe_name(s):
    return re.sub(r'[^A-Za-z0-9_-]+', '_', str(s))

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

def _format_millions(x, pos=None):
    try:
        x = float(x)
    except:
        return str(x)
    if abs(x) >= 1_000_000:
        return f"{x/1_000_000:.1f}M"
    if abs(x) >= 1_000:
        return f"{x/1_000:.0f}K"
    return f"{x:.0f}"

def build_pdf(sheet_title, charts_buffers, include_table=False, filtered_df=None, max_table_rows=200):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=20, rightMargin=20, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    elements = []
    elements.append(Paragraph(f"<para align='center'><b>{sheet_title} Report</b></para>", styles['Title']))
    elements.append(Spacer(1,12))
    elements.append(Paragraph("<para align='center'>Averroes Pharma - Auto Generated Dashboard</para>", styles['Heading3']))
    elements.append(Spacer(1,18))
    for img_buf, caption in charts_buffers:
        try:
            img_buf.seek(0)
            img = RLImage(img_buf, width=760, height=360)
            elements.append(img)
            elements.append(Spacer(1,6))
            elements.append(Paragraph(f"<para align='center'>{caption}</para>", styles['Normal']))
            elements.append(Spacer(1,12))
        except Exception:
            pass
    if include_table and (filtered_df is not None):
        table_df = filtered_df.copy().fillna("")
        if len(table_df) > max_table_rows:
            table_df = table_df.head(max_table_rows)
            elements.append(Paragraph(f"Showing first {max_table_rows} rows of filtered data", styles['Normal']))
            elements.append(Spacer(1,6))
        table_data = [table_df.columns.tolist()] + table_df.astype(str).values.tolist()
        tbl = Table(table_data, hAlign='CENTER')
        tbl.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#FFD700")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.black),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('GRID', (0,0), (-1,-1), 0.3, colors.grey),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ]))
        elements.append(tbl)
    doc.build(elements)
    buf.seek(0)
    return buf

def build_pptx(sheet_title, charts_buffers):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = slide.shapes.title
    title.text = f"{sheet_title} Dashboard"
    subtitle = slide.placeholders[1]
    subtitle.text = "Auto-generated by Averroes Pharma"
    for img_buf, caption in charts_buffers:
        try:
            img_buf.seek(0)
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            left = Inches(0.5)
            top = Inches(0.8)
            width = Inches(9)
            height = Inches(5)
            slide.shapes.add_picture(img_buf, left, top, width=width, height=height)
            txBox = slide.shapes.add_textbox(left, top + height + Inches(0.1), width, Inches(0.5))
            tf = txBox.text_frame
            p = tf.paragraphs[0]
            p.text = caption
            p.font.size = Pt(14)
            p.font.color.rgb = RGBColor(0, 0, 0)
            p.alignment = PP_ALIGN.CENTER
        except Exception:
            pass
    pptx_buffer = BytesIO()
    prs.save(pptx_buffer)
    pptx_buffer.seek(0)
    return pptx_buffer

# ---------- Shared progress UI (gold theme with shine + timer) ----------
def render_gold_progress(placeholder, percentage, message, elapsed_seconds):
    """Renders a gold animated progress bar with message and elapsed time."""
    html = f"""
    <style>
    .gold-container {{
        background-color: #00264d;
        border: 1px solid #FFD700;
        border-radius: 12px;
        padding: 14px;
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
        opacity: 0.8;
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

# ------------------ Navigation & Logo ------------------
st.markdown(
    """
    <div class="top-nav">
        <a href="#">Home</a>
        <a href="https://wa.me/201554694554" target="_blank">Contact</a>
        <a href="#info-section">Info</a>
    </div>
    """,
    unsafe_allow_html=True
)

logo_path = "logo.png"
if os.path.exists(logo_path):
    st.image(logo_path, width=200)
else:
    st.markdown('<div style="text-align:center; margin:20px 0; color:#FFD700; font-size:20px;">Averroes Pharma</div>', unsafe_allow_html=True)

st.markdown(
    """
    <div style="text-align:center; font-size:18px; color:#FFD700; margin-top:10px;">
        By <strong>Mohamed Abd ELGhany</strong> – 
        <a href="https://wa.me/201554694554" target="_blank" style="color:#FFD700; text-decoration:none;">
            01554694554 (WhatsApp)
        </a>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("<h1 style='text-align:center; color:#FFD700;'>💊 Averroes Pharma File Splitter & Dashboard</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align:center; color:white;'>✂ Split, Merge, Image-to-PDF & Auto Dashboard Generator</h3>", unsafe_allow_html=True)

# ------------------ Tabs ------------------
tab1, tab2, tab3, tab4 = st.tabs([
    "📂 Split & Merge", 
    "📷 Image to PDF", 
    "📊 Auto Dashboard", 
    "ℹ️ Info"
])

# ------------------ Tab 1: Split & Merge (Updated: Unified Gold Progress + Timer) ------------------
with tab1:
    st.markdown("### ✂ Split Excel/CSV File")
    uploaded_file = st.file_uploader(
        "📂 Upload Excel or CSV File (Splitter/Merge)",
        type=["xlsx", "csv"],
        accept_multiple_files=False,
        key=f"split_uploader_{st.session_state.clear_counter}"
    )
    if uploaded_file:
        display_uploaded_files([uploaded_file], "Excel/CSV")
        if st.button("🗑️ Clear Uploaded File", key="clear_split"):
            st.session_state.clear_counter += 1
            st.rerun()
        try:
            file_ext = uploaded_file.name.split('.')[-1].lower()
            if file_ext == "csv":
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
            col_to_split = st.selectbox(
                "Split by Column",
                df.columns,
                help="Select the column to split by, such as 'Brick' or 'Area Manager'"
            )
            st.markdown("### ⚙️ Split Options")
            split_option = st.radio(
                "Choose split method:",
                ["Split by Column Values", "Split Each Sheet into Separate File"],
                index=0,
                help="Choose 'Split by Column Values' to split the current sheet by column values. Choose 'Split Each Sheet into Separate File' to create a separate file for each sheet."
            )
            if st.button("🚀 Start Split"):
                with st.spinner("Splitting process in progress..."):
                    if LOTTIE_SPLIT:
                        st.markdown("<div style='text-align:center;'>", unsafe_allow_html=True)
                        st_lottie(LOTTIE_SPLIT, height=160, key="lottie_split_small")
                        st.markdown("</div>", unsafe_allow_html=True)
                    def clean_name(name):
                        name = str(name).strip()
                        invalid_chars = r'[\\/*?:\[\]|<>"]'
                        cleaned = re.sub(invalid_chars, '_', name)
                        return cleaned[:30] if cleaned else "Sheet"
                    # progress placeholder + timer start
                    progress_placeholder = st.empty()
                    start_time = time.time()
                    if file_ext == "csv":
                        # ===== CSV Split =====
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
                        # finish
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
                    else:
                        # ===== Excel Split =====
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
                                    # Copy headers (preserve style where possible)
                                    for cell in ws[1]:
                                        dst_cell = new_ws.cell(1, cell.column, cell.value)
                                        if cell.has_style:
                                            try:
                                                dst_cell.font = Font(
                                                    name=cell.font.name, size=cell.font.size,
                                                    bold=cell.font.bold, italic=cell.font.italic,
                                                    color=cell.font.color
                                                )
                                                if cell.fill and cell.fill.fill_type:
                                                    dst_cell.fill = PatternFill(
                                                        fill_type=cell.fill.fill_type,
                                                        start_color=cell.fill.start_color,
                                                        end_color=cell.fill.end_color
                                                    )
                                                dst_cell.alignment = Alignment(
                                                    horizontal=cell.alignment.horizontal,
                                                    vertical=cell.alignment.vertical,
                                                    wrap_text=cell.alignment.wrap_text
                                                )
                                                dst_cell.border = Border(
                                                    left=cell.border.left, right=cell.border.right,
                                                    top=cell.border.top, bottom=cell.border.bottom
                                                )
                                                dst_cell.number_format = cell.number_format
                                            except Exception:
                                                pass
                                    # Copy data rows
                                    row_idx = 2
                                    for row in ws.iter_rows(min_row=2):
                                        cell_in_col = row[col_idx - 1]
                                        if cell_in_col.value == value:
                                            for src_cell in row:
                                                dst_cell = new_ws.cell(row_idx, src_cell.column, src_cell.value)
                                                if src_cell.has_style:
                                                    try:
                                                        dst_cell.font = src_cell.font
                                                        dst_cell.fill = src_cell.fill
                                                        dst_cell.border = src_cell.border
                                                        dst_cell.alignment = src_cell.alignment
                                                        dst_cell.number_format = src_cell.number_format
                                                    except Exception:
                                                        pass
                                            row_idx += 1
                                    try:
                                        for col_letter in ws.column_dimensions:
                                            if ws.column_dimensions[col_letter].width:
                                                new_ws.column_dimensions[col_letter].width = ws.column_dimensions[col_letter].width
                                    except Exception:
                                        pass
                                    file_buffer = BytesIO()
                                    new_wb.save(file_buffer)
                                    file_buffer.seek(0)
                                    file_name = f"{clean_name(value)}.xlsx"
                                    zip_file.writestr(file_name, file_buffer.read())
                                    elapsed = time.time() - start_time
                                    pct = int(((i + 1) / total) * 100)
                                    render_gold_progress(progress_placeholder, pct, f"Splitting {i+1}/{total}: {value}", elapsed)
                            # finish
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
                                    new_ws = new_wb.create_sheet(title=sheet_name)
                                    for row in src_ws.iter_rows():
                                        for src_cell in row:
                                            dst_cell = new_ws.cell(src_cell.row, src_cell.column, src_cell.value)
                                            if src_cell.has_style:
                                                try:
                                                    dst_cell.font = src_cell.font
                                                    dst_cell.fill = src_cell.fill
                                                    dst_cell.border = src_cell.border
                                                    dst_cell.alignment = src_cell.alignment
                                                    dst_cell.number_format = src_cell.number_format
                                                except Exception:
                                                    pass
                                    # Copy merged cells + dimensions
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
                            # finish
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
        except Exception as e:
            st.error(f"❌ Error processing file: {e}")
    else:
        st.markdown("<p style='text-align:center; color:#FFD700;'>⚠️ No file uploaded yet for splitting.</p>", unsafe_allow_html=True)
    # ------------------ Merge (below Split UI in same tab) ------------------
    st.markdown("<hr class='divider-dashed'>", unsafe_allow_html=True)
    st.markdown("### 🔄 Merge Excel/CSV Files")
    merge_files = st.file_uploader(
        "📤 Upload Excel or CSV Files to Merge",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        key=f"merge_uploader_{st.session_state.clear_counter}"
    )
    if merge_files:
        display_uploaded_files(merge_files, "Excel/CSV")
        if st.button("🗑️ Clear All Merged Files", key="clear_merge"):
            st.session_state.clear_counter += 1
            st.rerun()
        if st.button("✨ Merge Files"):
            with st.spinner("Merging files..."):
                # show lottie if available
                if LOTTIE_MERGE:
                    st.markdown("<div style='text-align:center;'>", unsafe_allow_html=True)
                    st_lottie(LOTTIE_MERGE, height=140, key="lottie_merge_small")
                    st.markdown("</div>", unsafe_allow_html=True)
                # progress placeholder + timer
                merge_progress_placeholder = st.empty()
                merge_start = time.time()
                # shared helper for merge rendering
                def render_merge(pct, msg):
                    elapsed = time.time() - merge_start
                    render_gold_progress(merge_progress_placeholder, pct, msg, elapsed)
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
                            # Copy header only once
                            if idx == 0:
                                for row in src_ws.iter_rows(min_row=1, max_row=1):
                                    for cell in row:
                                        dst_cell = merged_ws.cell(current_row, cell.column, cell.value)
                                        if cell.has_style:
                                            try:
                                                dst_cell.font = cell.font
                                                dst_cell.fill = cell.fill
                                                dst_cell.border = cell.border
                                                dst_cell.alignment = cell.alignment
                                                dst_cell.number_format = cell.number_format
                                            except Exception:
                                                pass
                                current_row += 1
                            # Copy data rows (skip header)
                            for row in src_ws.iter_rows(min_row=2):
                                for cell in row:
                                    dst_cell = merged_ws.cell(current_row, cell.column, cell.value)
                                    if cell.has_style:
                                        try:
                                            dst_cell.font = cell.font
                                            dst_cell.fill = cell.fill
                                            dst_cell.border = cell.border
                                            dst_cell.alignment = cell.alignment
                                            dst_cell.number_format = cell.number_format
                                        except Exception:
                                            pass
                                current_row += 1
                            # Copy column widths
                            try:
                                for col_letter in src_ws.column_dimensions:
                                    if src_ws.column_dimensions[col_letter].width:
                                        merged_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width
                            except Exception:
                                pass
                            pct = int(((idx + 1) / total) * 100)
                            render_merge(pct, f"Merging file {idx+1}/{total}: {file.name}")
                        output_buffer = BytesIO()
                        merged_wb.save(output_buffer)
                        output_buffer.seek(0)
                        render_merge(100, "✅ Merging Completed Successfully!")
                        time.sleep(0.25)
                        merge_progress_placeholder.empty()
                        st.success("✅ Merged successfully with original formatting preserved!")
                        st.download_button(
                            label="📥 Download Merged File (Formatted)",
                            data=output_buffer.getvalue(),
                            file_name="Merged_Consolidated_Formatted.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        # Fallback for mixed or CSV
                        all_dfs = []
                        for i, file in enumerate(merge_files):
                            ext = file.name.split('.')[-1].lower()
                            df = pd.read_csv(file) if ext == "csv" else pd.read_excel(file)
                            all_dfs.append(df)
                            pct = int(((i + 1) / total) * 100)
                            render_merge(pct, f"Merging file {i+1}/{total}: {file.name}")
                        merged_df = pd.concat(all_dfs, ignore_index=True)
                        output_buffer = BytesIO()
                        merged_df.to_excel(output_buffer, index=False, engine='openpyxl')
                        output_buffer.seek(0)
                        render_merge(100, "✅ Merging Completed Successfully!")
                        time.sleep(0.25)
                        merge_progress_placeholder.empty()
                        st.success("✅ Merged successfully (formatting not preserved for CSV/mixed files).")
                        st.download_button(
                            label="📥 Download Merged File (Excel)",
                            data=output_buffer.getvalue(),
                            file_name="Merged_Consolidated.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except Exception as e:
                    st.error(f"❌ Error during merge: {e}")
                    merge_progress_placeholder.empty()

# ------------------ Tab 2: Image to PDF ------------------
with tab2:
    st.markdown("### 📷 Convert Images to PDF")
    uploaded_images = st.file_uploader(
        "📤 Upload JPG/JPEG/PNG Images to Convert to PDF",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key=f"image_uploader_{st.session_state.clear_counter}"
    )
    if uploaded_images:
        display_uploaded_files(uploaded_images, "Image")
        if st.button("🗑️ Clear All Images", key="clear_images"):
            st.session_state.clear_counter += 1
            st.rerun()
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
                bordered = cv2.copyMakeBorder(
                    enhanced,
                    top=border_size,
                    bottom=border_size,
                    left=border_size,
                    right=border_size,
                    borderType=cv2.BORDER_CONSTANT,
                    value=[255, 255, 255]
                )
                if bordered.dtype != np.uint8:
                    bordered = np.clip(bordered, 0, 255).astype(np.uint8)
                result = cv2.cvtColor(bordered, cv2.COLOR_GRAY2RGB)
                return Image.fromarray(result)
            if st.button("🖨️ Create PDF (CamScanner Style)"):
                with st.spinner("Enhancing images for PDF..."):
                    if LOTTIE_IMAGE:
                        st.markdown("<div style='text-align:center;'>", unsafe_allow_html=True)
                        st_lottie(LOTTIE_IMAGE, height=180, key="lottie_image_enhance")
                        st.markdown("</div>", unsafe_allow_html=True)
                    try:
                        first_image = Image.open(uploaded_images[0])
                        first_image_enhanced = enhance_image_for_pdf(first_image)
                        other_images = []
                        for img_file in uploaded_images[1:]:
                            img = Image.open(img_file)
                            enhanced_img = enhance_image_for_pdf(img)
                            other_images.append(enhanced_img.convert("RGB"))
                        pdf_buffer = BytesIO()
                        first_image_enhanced.save(pdf_buffer, format="PDF", save_all=True, append_images=other_images)
                        pdf_buffer.seek(0)
                        st.success("✅ Enhanced PDF created successfully!")
                        st.download_button(
                            label="📥 Download Enhanced PDF",
                            data=pdf_buffer.getvalue(),
                            file_name="Enhanced_Images_CamScanner.pdf",
                            mime="application/pdf"
                        )
                    except Exception as e:
                        st.error(f"❌ Error creating enhanced PDF: {e}")
        except ImportError:
            st.warning("⚠️ CamScanner effect requires 'opencv-python'. Install it to enable this feature.")
        if st.button("🖨️ Create PDF (Original Quality)"):
            with st.spinner("Converting images to PDF..."):
                try:
                    first_image = Image.open(uploaded_images[0]).convert("RGB")
                    other_images = []
                    for img_file in uploaded_images[1:]:
                        img = Image.open(img_file).convert("RGB")
                        other_images.append(img)
                    pdf_buffer = BytesIO()
                    first_image.save(pdf_buffer, format="PDF", save_all=True, append_images=other_images)
                    pdf_buffer.seek(0)
                    st.success("✅ PDF created successfully!")
                    st.download_button(
                        label="📥 Download Original PDF",
                        data=pdf_buffer.getvalue(),
                        file_name="Images_Combined.pdf",
                        mime="application/pdf"
                    )
                except Exception as e:
                    st.error(f"❌ Error creating PDF: {e}")
    else:
        st.info("📤 Please upload one or more JPG/JPEG/PNG images to convert them into a single PDF file.")

# ------------------ Tab 3: Dashboard ------------------
with tab3:
    st.markdown("### 📊 Interactive Auto Dashboard Generator")
    dashboard_file = st.file_uploader(
        "📊 Upload Excel or CSV File for Dashboard (Auto)",
        type=["xlsx", "csv"],
        key=f"dashboard_uploader_{st.session_state.clear_counter}"
    )
    if dashboard_file:
        display_uploaded_files([dashboard_file], "Excel/CSV")
        if st.button("🗑️ Clear Dashboard File", key="clear_dashboard"):
            st.session_state.clear_counter # تم دمج الكود من ملف `edit.txt` (الذي يحتوي على واجهة شريط التقدم الذهبية الموحّدة مع مؤقّت) داخل الكود الكامل من `Code Edit.txt`، مع الحفاظ على جميع الوظائف الأخرى (مثل تحويل الصور إلى PDF، لوحة التحكم التلقائية، إلخ) دون أي تغيير.

#التعديل الوحيد تم في **Tab 1: Split & Merge**، حيث تم استبدال قسم التقسيم والدمج الأصلي بنسخة مُحسّنة تُظهر شريط تقدم ذهبي موحّد مع مؤقّت زمني لكل من عمليات **Split** و**Merge**.

إليك الكود الكامل بعد الدمج:

```python
# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
import re
import os
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl import load_workbook, Workbook
# ====== Dashboard & Reporting ======
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image as RLImage, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from PIL import Image
from sklearn.linear_model import LinearRegression
import numpy as np
# === Lottie animation support ===
from streamlit_lottie import st_lottie
import requests
import json
import time  # Added for timer in progress bar

def load_lottie_url(url: str):
    """تحميل Lottie JSON من رابط خارجي"""
    try:
        r = requests.get(url, timeout=8)
        if r.status_code == 200:
            return r.json()
    except Exception:
        return None
    return None

# Initialize session state
if 'clear_counter' not in st.session_state:
    st.session_state.clear_counter = 0

# ------------------ Page Setup ------------------
st.set_page_config(
    page_title="Averroes Pharma File Splitter & Dashboard",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

LOTTIE_SPLIT = load_lottie_url("https://assets9.lottiefiles.com/packages/lf20_wx9z5gxb.json")   # split
LOTTIE_MERGE = load_lottie_url("https://assets10.lottiefiles.com/packages/lf20_cg3rwjul.json")  # merge
LOTTIE_IMAGE = load_lottie_url("https://assets2.lottiefiles.com/private_files/lf30_cgfdhxgx.json")  # image/pdf
LOTTIE_DASH  = load_lottie_url("https://assets8.lottiefiles.com/packages/lf20_tno6cg2w.json")   # dashboard
LOTTIE_PDF   = load_lottie_url("https://assets1.lottiefiles.com/packages/lf20_zyu0ct3i.json")   # dashboard PDF

# Hide default Streamlit elements
hide_default = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
"""
st.markdown(hide_default, unsafe_allow_html=True)

# ------------------ Custom CSS ------------------
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
        gap: 20px;
        padding: 10px 30px;
        background-color: #001a33;
        border-bottom: 1px solid #FFD700;
        font-size: 18px;
        color: white;
    }
    .top-nav a {
        color: #FFD700;
        text-decoration: none;
        font-weight: bold;
        padding: 5px 10px;
        border-radius: 8px;
        transition: all 0.3s ease;
    }
    .top-nav a:hover {
        background-color: #FFD700;
        color: black;
    }
    label, .stSelectbox label, .stFileUploader label {
        color: #FFD700 !important;
        font-size: 18px !important;
        font-weight: bold !important;
    }
    .stButton>button, .stDownloadButton>button {
        background-color: #FFD700 !important;
        color: black !important;
        font-weight: bold !important;
        font-size: 18px !important;
        border-radius: 12px !important;
        padding: 12px 24px !important;
        border: none !important;
        box-shadow: 0 4px 8px rgba(0,0,0,0.3) !important;
        transition: all 0.3s ease !important;
        margin-top: 10px !important;
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        background-color: #FFC107 !important;
        transform: scale(1.08);
        box-shadow: 0 6px 12px rgba(0,0,0,0.4) !important;
    }
    .kpi-card {
        padding: 16px;
        border-radius: 12px;
        color: white;
        text-align: center;
        box-shadow: 0 6px 18px rgba(0,0,0,0.3);
        font-weight: 700;
        margin: 8px;
    }
    .kpi-title { font-size: 14px; opacity: 0.9; }
    .kpi-value { font-size: 22px; margin-top:6px; }
    hr.divider {
        border: 1px solid #FFD700;
        opacity: 0.6;
        margin: 30px 0;
    }
    hr.divider-dashed {
        border: 1px dashed #FFD700;
        opacity: 0.7;
        margin: 25px 0;
    }
    .stDataFrame {
        box-shadow: 0 4px 10px rgba(0,0,0,0.2);
        border-radius: 12px;
        overflow: hidden;
        margin: 10px 0;
    }
    .stFileUploader {
        border: 2px dashed #FFD700;
        border-radius: 10px;
        padding: 15px;
        background-color: rgba(255, 215, 0, 0.1);
    }
    .guide-title {
        color: #FFD700;
        font-weight: bold;
        font-size: 20px;
    }
    .chart-card {
        background-color: #00264d;
        border: 1px solid #FFD700;
        border-radius: 12px;
        padding: 12px;
        margin: 10px 0;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3);
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 20px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        padding-top: 10px;
        padding-bottom: 10px;
        font-size: 18px;
        font-weight: bold;
        border-radius: 10px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #FFD700 !important;
        color: black !important;
        border: 2px solid #FFC107 !important;
    }
    .stTabs [aria-selected="false"] {
        background-color: #003366 !important;
        color: #FFD700 !important;
        border: 1px solid #FFD700 !important;
    }
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ------------------ Helper Functions ------------------
def display_uploaded_files(file_list, file_type="Excel/CSV"):
    if file_list:
        st.markdown("### 📁 Uploaded Files:")
        for i, f in enumerate(file_list):
            st.markdown(
                f"<div style='background:#003366; color:white; padding:4px 8px; border-radius:4px; margin:2px 0; display:inline-block;'>"
                f"{i+1}. {f.name} ({f.size//1024} KB)</div>",
                unsafe_allow_html=True
            )

def _safe_name(s):
    return re.sub(r'[^A-Za-z0-9_-]+', '_', str(s))

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

def _format_millions(x, pos=None):
    try:
        x = float(x)
    except:
        return str(x)
    if abs(x) >= 1_000_000:
        return f"{x/1_000_000:.1f}M"
    if abs(x) >= 1_000:
        return f"{x/1_000:.0f}K"
    return f"{x:.0f}"

def build_pdf(sheet_title, charts_buffers, include_table=False, filtered_df=None, max_table_rows=200):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=20, rightMargin=20, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    elements = []
    elements.append(Paragraph(f"<para align='center'><b>{sheet_title} Report</b></para>", styles['Title']))
    elements.append(Spacer(1,12))
    elements.append(Paragraph("<para align='center'>Averroes Pharma - Auto Generated Dashboard</para>", styles['Heading3']))
    elements.append(Spacer(1,18))
    for img_buf, caption in charts_buffers:
        try:
            img_buf.seek(0)
            img = RLImage(img_buf, width=760, height=360)
            elements.append(img)
            elements.append(Spacer(1,6))
            elements.append(Paragraph(f"<para align='center'>{caption}</para>", styles['Normal']))
            elements.append(Spacer(1,12))
        except Exception:
            pass
    if include_table and (filtered_df is not None):
        table_df = filtered_df.copy().fillna("")
        if len(table_df) > max_table_rows:
            table_df = table_df.head(max_table_rows)
            elements.append(Paragraph(f"Showing first {max_table_rows} rows of filtered data", styles['Normal']))
            elements.append(Spacer(1,6))
        table_data = [table_df.columns.tolist()] + table_df.astype(str).values.tolist()
        tbl = Table(table_data, hAlign='CENTER')
        tbl.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#FFD700")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.black),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('GRID', (0,0), (-1,-1), 0.3, colors.grey),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ]))
        elements.append(tbl)
    doc.build(elements)
    buf.seek(0)
    return buf

def build_pptx(sheet_title, charts_buffers):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = slide.shapes.title
    title.text = f"{sheet_title} Dashboard"
    subtitle = slide.placeholders[1]
    subtitle.text = "Auto-generated by Averroes Pharma"
    for img_buf, caption in charts_buffers:
        try:
            img_buf.seek(0)
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            left = Inches(0.5)
            top = Inches(0.8)
            width = Inches(9)
            height = Inches(5)
            slide.shapes.add_picture(img_buf, left, top, width=width, height=height)
            txBox = slide.shapes.add_textbox(left, top + height + Inches(0.1), width, Inches(0.5))
            tf = txBox.text_frame
            p = tf.paragraphs[0]
            p.text = caption
            p.font.size = Pt(14)
            p.font.color.rgb = RGBColor(0, 0, 0)
            p.alignment = PP_ALIGN.CENTER
        except Exception:
            pass
    pptx_buffer = BytesIO()
    prs.save(pptx_buffer)
    pptx_buffer.seek(0)
    return pptx_buffer

# ---------- Shared progress UI (gold theme with shine + timer) ----------
def render_gold_progress(placeholder, percentage, message, elapsed_seconds):
    """Renders a gold animated progress bar with message and elapsed time."""
    html = f"""
    <style>
    .gold-container {{
        background-color: #00264d;
        border: 1px solid #FFD700;
        border-radius: 12px;
        padding: 14px;
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
        opacity: 0.8;
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

# ------------------ Navigation & Logo ------------------
st.markdown(
    """
    <div class="top-nav">
        <a href="#">Home</a>
        <a href="https://wa.me/201554694554" target="_blank">Contact</a>
        <a href="#info-section">Info</a>
    </div>
    """,
    unsafe_allow_html=True
)

logo_path = "logo.png"
if os.path.exists(logo_path):
    st.image(logo_path, width=200)
else:
    st.markdown('<div style="text-align:center; margin:20px 0; color:#FFD700; font-size:20px;">Averroes Pharma</div>', unsafe_allow_html=True)

st.markdown(
    """
    <div style="text-align:center; font-size:18px; color:#FFD700; margin-top:10px;">
        By <strong>Mohamed Abd ELGhany</strong> – 
        <a href="https://wa.me/201554694554" target="_blank" style="color:#FFD700; text-decoration:none;">
            01554694554 (WhatsApp)
        </a>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("<h1 style='text-align:center; color:#FFD700;'>💊 Averroes Pharma File Splitter & Dashboard</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align:center; color:white;'>✂ Split, Merge, Image-to-PDF & Auto Dashboard Generator</h3>", unsafe_allow_html=True)

# ------------------ Tabs ------------------
tab1, tab2, tab3, tab4 = st.tabs([
    "📂 Split & Merge", 
    "📷 Image to PDF", 
    "📊 Auto Dashboard", 
    "ℹ️ Info"
])

# ------------------ Tab 1: Split & Merge (Updated: Unified Gold Progress + Timer) ------------------
with tab1:
    st.markdown("### ✂ Split Excel/CSV File")
    uploaded_file = st.file_uploader(
        "📂 Upload Excel or CSV File (Splitter/Merge)",
        type=["xlsx", "csv"],
        accept_multiple_files=False,
        key=f"split_uploader_{st.session_state.clear_counter}"
    )
    if uploaded_file:
        display_uploaded_files([uploaded_file], "Excel/CSV")
        if st.button("🗑️ Clear Uploaded File", key="clear_split"):
            st.session_state.clear_counter += 1
            st.rerun()
        try:
            file_ext = uploaded_file.name.split('.')[-1].lower()
            if file_ext == "csv":
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
            col_to_split = st.selectbox(
                "Split by Column",
                df.columns,
                help="Select the column to split by, such as 'Brick' or 'Area Manager'"
            )
            st.markdown("### ⚙️ Split Options")
            split_option = st.radio(
                "Choose split method:",
                ["Split by Column Values", "Split Each Sheet into Separate File"],
                index=0,
                help="Choose 'Split by Column Values' to split the current sheet by column values. Choose 'Split Each Sheet into Separate File' to create a separate file for each sheet."
            )
            if st.button("🚀 Start Split"):
                with st.spinner("Splitting process in progress..."):
                    if LOTTIE_SPLIT:
                        st.markdown("<div style='text-align:center;'>", unsafe_allow_html=True)
                        st_lottie(LOTTIE_SPLIT, height=160, key="lottie_split_small")
                        st.markdown("</div>", unsafe_allow_html=True)
                    def clean_name(name):
                        name = str(name).strip()
                        invalid_chars = r'[\\/*?:\[\]|<>"]'
                        cleaned = re.sub(invalid_chars, '_', name)
                        return cleaned[:30] if cleaned else "Sheet"
                    # progress placeholder + timer start
                    progress_placeholder = st.empty()
                    start_time = time.time()
                    if file_ext == "csv":
                        # ===== CSV Split =====
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
                        # finish
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
                    else:
                        # ===== Excel Split =====
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
                                    # Copy headers (preserve style where possible)
                                    for cell in ws[1]:
                                        dst_cell = new_ws.cell(1, cell.column, cell.value)
                                        if cell.has_style:
                                            try:
                                                dst_cell.font = Font(
                                                    name=cell.font.name, size=cell.font.size,
                                                    bold=cell.font.bold, italic=cell.font.italic,
                                                    color=cell.font.color
                                                )
                                                if cell.fill and cell.fill.fill_type:
                                                    dst_cell.fill = PatternFill(
                                                        fill_type=cell.fill.fill_type,
                                                        start_color=cell.fill.start_color,
                                                        end_color=cell.fill.end_color
                                                    )
                                                dst_cell.alignment = Alignment(
                                                    horizontal=cell.alignment.horizontal,
                                                    vertical=cell.alignment.vertical,
                                                    wrap_text=cell.alignment.wrap_text
                                                )
                                                dst_cell.border = Border(
                                                    left=cell.border.left, right=cell.border.right,
                                                    top=cell.border.top, bottom=cell.border.bottom
                                                )
                                                dst_cell.number_format = cell.number_format
                                            except Exception:
                                                pass
                                    # Copy data rows
                                    row_idx = 2
                                    for row in ws.iter_rows(min_row=2):
                                        cell_in_col = row[col_idx - 1]
                                        if cell_in_col.value == value:
                                            for src_cell in row:
                                                dst_cell = new_ws.cell(row_idx, src_cell.column, src_cell.value)
                                                if src_cell.has_style:
                                                    try:
                                                        dst_cell.font = src_cell.font
                                                        dst_cell.fill = src_cell.fill
                                                        dst_cell.border = src_cell.border
                                                        dst_cell.alignment = src_cell.alignment
                                                        dst_cell.number_format = src_cell.number_format
                                                    except Exception:
                                                        pass
                                            row_idx += 1
                                    try:
                                        for col_letter in ws.column_dimensions:
                                            if ws.column_dimensions[col_letter].width:
                                                new_ws.column_dimensions[col_letter].width = ws.column_dimensions[col_letter].width
                                    except Exception:
                                        pass
                                    file_buffer = BytesIO()
                                    new_wb.save(file_buffer)
                                    file_buffer.seek(0)
                                    file_name = f"{clean_name(value)}.xlsx"
                                    zip_file.writestr(file_name, file_buffer.read())
                                    elapsed = time.time() - start_time
                                    pct = int(((i + 1) / total) * 100)
                                    render_gold_progress(progress_placeholder, pct, f"Splitting {i+1}/{total}: {value}", elapsed)
                            # finish
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
                                    new_ws = new_wb.create_sheet(title=sheet_name)
                                    for row in src_ws.iter_rows():
                                        for src_cell in row:
                                            dst_cell = new_ws.cell(src_cell.row, src_cell.column, src_cell.value)
                                            if src_cell.has_style:
                                                try:
                                                    dst_cell.font = src_cell.font
                                                    dst_cell.fill = src_cell.fill
                                                    dst_cell.border = src_cell.border
                                                    dst_cell.alignment = src_cell.alignment
                                                    dst_cell.number_format = src_cell.number_format
                                                except Exception:
                                                    pass
                                    # Copy merged cells + dimensions
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
                            # finish
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
        except Exception as e:
            st.error(f"❌ Error processing file: {e}")
    else:
        st.markdown("<p style='text-align:center; color:#FFD700;'>⚠️ No file uploaded yet for splitting.</p>", unsafe_allow_html=True)
    # ------------------ Merge (below Split UI in same tab) ------------------
    st.markdown("<hr class='divider-dashed'>", unsafe_allow_html=True)
    st.markdown("### 🔄 Merge Excel/CSV Files")
    merge_files = st.file_uploader(
        "📤 Upload Excel or CSV Files to Merge",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        key=f"merge_uploader_{st.session_state.clear_counter}"
    )
    if merge_files:
        display_uploaded_files(merge_files, "Excel/CSV")
        if st.button("🗑️ Clear All Merged Files", key="clear_merge"):
            st.session_state.clear_counter += 1
            st.rerun()
        if st.button("✨ Merge Files"):
            with st.spinner("Merging files..."):
                # show lottie if available
                if LOTTIE_MERGE:
                    st.markdown("<div style='text-align:center;'>", unsafe_allow_html=True)
                    st_lottie(LOTTIE_MERGE, height=140, key="lottie_merge_small")
                    st.markdown("</div>", unsafe_allow_html=True)
                # progress placeholder + timer
                merge_progress_placeholder = st.empty()
                merge_start = time.time()
                # shared helper for merge rendering
                def render_merge(pct, msg):
                    elapsed = time.time() - merge_start
                    render_gold_progress(merge_progress_placeholder, pct, msg, elapsed)
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
                            # Copy header only once
                            if idx == 0:
                                for row in src_ws.iter_rows(min_row=1, max_row=1):
                                    for cell in row:
                                        dst_cell = merged_ws.cell(current_row, cell.column, cell.value)
                                        if cell.has_style:
                                            try:
                                                dst_cell.font = cell.font
                                                dst_cell.fill = cell.fill
                                                dst_cell.border = cell.border
                                                dst_cell.alignment = cell.alignment
                                                dst_cell.number_format = cell.number_format
                                            except Exception:
                                                pass
                                current_row += 1
                            # Copy data rows (skip header)
                            for row in src_ws.iter_rows(min_row=2):
                                for cell in row:
                                    dst_cell = merged_ws.cell(current_row, cell.column, cell.value)
                                    if cell.has_style:
                                        try:
                                            dst_cell.font = cell.font
                                            dst_cell.fill = cell.fill
                                            dst_cell.border = cell.border
                                            dst_cell.alignment = cell.alignment
                                            dst_cell.number_format = cell.number_format
                                        except Exception:
                                            pass
                                current_row += 1
                            # Copy column widths
                            try:
                                for col_letter in src_ws.column_dimensions:
                                    if src_ws.column_dimensions[col_letter].width:
                                        merged_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width
                            except Exception:
                                pass
                            pct = int(((idx + 1) / total) * 100)
                            render_merge(pct, f"Merging file {idx+1}/{total}: {file.name}")
                        output_buffer = BytesIO()
                        merged_wb.save(output_buffer)
                        output_buffer.seek(0)
                        render_merge(100, "✅ Merging Completed Successfully!")
                        time.sleep(0.25)
                        merge_progress_placeholder.empty()
                        st.success("✅ Merged successfully with original formatting preserved!")
                        st.download_button(
                            label="📥 Download Merged File (Formatted)",
                            data=output_buffer.getvalue(),
                            file_name="Merged_Consolidated_Formatted.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        # Fallback for mixed or CSV
                        all_dfs = []
                        for i, file in enumerate(merge_files):
                            ext = file.name.split('.')[-1].lower()
                            df = pd.read_csv(file) if ext == "csv" else pd.read_excel(file)
                            all_dfs.append(df)
                            pct = int(((i + 1) / total) * 100)
                            render_merge(pct, f"Merging file {i+1}/{total}: {file.name}")
                        merged_df = pd.concat(all_dfs, ignore_index=True)
                        output_buffer = BytesIO()
                        merged_df.to_excel(output_buffer, index=False, engine='openpyxl')
                        output_buffer.seek(0)
                        render_merge(100, "✅ Merging Completed Successfully!")
                        time.sleep(0.25)
                        merge_progress_placeholder.empty()
                        st.success("✅ Merged successfully (formatting not preserved for CSV/mixed files).")
                        st.download_button(
                            label="📥 Download Merged File (Excel)",
                            data=output_buffer.getvalue(),
                            file_name="Merged_Consolidated.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except Exception as e:
                    st.error(f"❌ Error during merge: {e}")
                    merge_progress_placeholder.empty()

# ------------------ Tab 2: Image to PDF ------------------
with tab2:
    st.markdown("### 📷 Convert Images to PDF")
    uploaded_images = st.file_uploader(
        "📤 Upload JPG/JPEG/PNG Images to Convert to PDF",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key=f"image_uploader_{st.session_state.clear_counter}"
    )
    if uploaded_images:
        display_uploaded_files(uploaded_images, "Image")
        if st.button("🗑️ Clear All Images", key="clear_images"):
            st.session_state.clear_counter += 1
            st.rerun()
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
                bordered = cv2.copyMakeBorder(
                    enhanced,
                    top=border_size,
                    bottom=border_size,
                    left=border_size,
                    right=border_size,
                    borderType=cv2.BORDER_CONSTANT,
                    value=[255, 255, 255]
                )
                if bordered.dtype != np.uint8:
                    bordered = np.clip(bordered, 0, 255).astype(np.uint8)
                result = cv2.cvtColor(bordered, cv2.COLOR_GRAY2RGB)
                return Image.fromarray(result)
            if st.button("🖨️ Create PDF (CamScanner Style)"):
                with st.spinner("Enhancing images for PDF..."):
                    if LOTTIE_IMAGE:
                        st.markdown("<div style='text-align:center;'>", unsafe_allow_html=True)
                        st_lottie(LOTTIE_IMAGE, height=180, key="lottie_image_enhance")
                        st.markdown("</div>", unsafe_allow_html=True)
                    try:
                        first_image = Image.open(uploaded_images[0])
                        first_image_enhanced = enhance_image_for_pdf(first_image)
                        other_images = []
                        for img_file in uploaded_images[1:]:
                            img = Image.open(img_file)
                            enhanced_img = enhance_image_for_pdf(img)
                            other_images.append(enhanced_img.convert("RGB"))
                        pdf_buffer = BytesIO()
                        first_image_enhanced.save(pdf_buffer, format="PDF", save_all=True, append_images=other_images)
                        pdf_buffer.seek(0)
                        st.success("✅ Enhanced PDF created successfully!")
                        st.download_button(
                            label="📥 Download Enhanced PDF",
                            data=pdf_buffer.getvalue(),
                            file_name="Enhanced_Images_CamScanner.pdf",
                            mime="application/pdf"
                        )
                    except Exception as e:
                        st.error(f"❌ Error creating enhanced PDF: {e}")
        except ImportError:
            st.warning("⚠️ CamScanner effect requires 'opencv-python'. Install it to enable this feature.")
        if st.button("🖨️ Create PDF (Original Quality)"):
            with st.spinner("Converting images to PDF..."):
                try:
                    first_image = Image.open(uploaded_images[0]).convert("RGB")
                    other_images = []
                    for img_file in uploaded_images[1:]:
                        img = Image.open(img_file).convert("RGB")
                        other_images.append(img)
                    pdf_buffer = BytesIO()
                    first_image.save(pdf_buffer, format="PDF", save_all=True, append_images=other_images)
                    pdf_buffer.seek(0)
                    st.success("✅ PDF created successfully!")
                    st.download_button(
                        label="📥 Download Original PDF",
                        data=pdf_buffer.getvalue(),
                        file_name="Images_Combined.pdf",
                        mime="application/pdf"
                    )
                except Exception as e:
                    st.error(f"❌ Error creating PDF: {e}")
    else:
        st.info("📤 Please upload one or more JPG/JPEG/PNG images to convert them into a single PDF file.")

# ------------------ Tab 3: Dashboard ------------------
with tab3:
    st.markdown("### 📊 Interactive Auto Dashboard Generator")
    dashboard_file = st.file_uploader(
        "📊 Upload Excel or CSV File for Dashboard (Auto)",
        type=["xlsx", "csv"],
        key=f"dashboard_uploader_{st.session_state.clear_counter}"
    )
    if dashboard_file:
        display_uploaded_files([dashboard_file], "Excel/CSV")
        if st.button("🗑️ Clear Dashboard File", key="clear_dashboard"):
            st.session_state.clear_counter += 1
            st.rerun()
        try:
            # =============== Progress Bar ===============
            progress_bar = st.progress(0)
            status_text = st.empty()
            def update_progress(pct, msg=""):
                progress_bar.progress(pct)
                status_text.text(f"🔄 {msg}... {pct}%")
            update_progress(10, "Loading file")
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
            update_progress(30, "Analyzing data")
            st.markdown("### 🔍 Data Preview (original)")
            st.dataframe(df0.head(), use_container_width=True)
            # =============== Detect Period Columns ===============
            numeric_cols = df0.select_dtypes(include='number').columns.tolist()
            period_cols = []
            base_names = {}
            for col in numeric_cols:
                match = re.search(r'(.+?)[_\s\-](\d{4}|Q[1-4]|[A-Za-z]+_\d{4})$', col.strip())
                if match:
                    base = match.group(1).strip()
                    period = match.group(2)
                    if base not in base_names:
                        base_names[base] = []
                    base_names[base].append((col, period))
            valid_periods = {}
            for base, cols in base_names.items():
                if len(cols) >= 2:
                    valid_periods[base] = sorted(cols, key=lambda x: x[1])
            period_comparison = None
            if valid_periods:
                base_key = list(valid_periods.keys())[0]
                cols_info = valid_periods[base_key]
                col1, period1 = cols_info[-2]
                col2, period2 = cols_info[-1]
                df0['__abs_change__'] = df0[col2] - df0[col1]
                df0['__pct_change__'] = df0['__abs_change__'] / df0[col1].replace(0, pd.NA)
                period_comparison = {'col1': col1, 'col2': col2, 'period1': period1, 'period2': period2, 'base': base_key}
                st.success(f"✅ Detected period comparison: {period1} vs {period2} for '{base_key}'")
            update_progress(50, "Processing filters")
            # =============== Handle Month Columns ===============
            month_names = ["jan","feb","mar","apr","may","jun","jul","aug","sep","sept","oct","nov","dec"]
            potential_months = [c for c in df0.columns if c.strip().lower() in month_names]
            if potential_months:
                id_vars = [c for c in df0.columns if c not in potential_months]
                value_vars = potential_months
                df_long = df0.melt(id_vars=id_vars, value_vars=value_vars, var_name="Month", value_name="Value")
                df_long["Month"] = df_long["Month"].astype(str)
                measure_col = "Value"
            else:
                numeric_cols = df0.select_dtypes(include='number').columns.tolist()
                measure_col = numeric_cols[0] if numeric_cols else None
                df_long = df0.copy()
            # =============== Select Measure Column ===============
            numeric_cols_in_long = df_long.select_dtypes(include='number').columns.tolist()
            if numeric_cols_in_long:
                user_measure_col = st.selectbox(
                    "🎯 Select Sales/Value Column (for KPIs & Charts)",
                    numeric_cols_in_long,
                    index=numeric_cols_in_long.index(measure_col) if measure_col in numeric_cols_in_long else 0
                )
                kpi_measure_col = user_measure_col
            else:
                kpi_measure_col = measure_col
            # =============== Identify Categorical Columns ===============
            cat_cols = [c for c in df_long.columns if df_long[c].dtype == "object" or df_long[c].dtype.name.startswith("category")]
            for c in df_long.columns:
                if c not in cat_cols and df_long[c].nunique(dropna=True) <= 100 and df_long[c].dtype not in ["float64", "int64"]:
                    cat_cols.append(c)
            cat_cols = [c for c in cat_cols if c is not None]
            # =============== Sidebar Filters ===============
            st.sidebar.header("🔍 Dynamic Filters")
            primary_filter_col = None
            if len(cat_cols) > 0:
                primary_filter_col = st.sidebar.selectbox("Primary Filter Column", ["-- None --"] + cat_cols, index=0)
                if primary_filter_col == "-- None --":
                    primary_filter_col = None
            primary_values = None
            if primary_filter_col:
                vals = df_long[primary_filter_col].dropna().astype(str).unique().tolist()
                try:
                    vals = sorted(vals)
                except:
                    pass
                primary_values = st.sidebar.multiselect(f"Filter values for {primary_filter_col}", vals, default=vals)
            other_filter_cols = st.sidebar.multiselect("Additional filter columns", [c for c in cat_cols if c != primary_filter_col], default=[])
            active_filters = {}
            for fc in other_filter_cols:
                opts = df_long[fc].dropna().astype(str).unique().tolist()
                try:
                    opts = sorted(opts)
                except:
                    pass
                sel = st.sidebar.multiselect(f"Filter: {fc}", opts, default=opts)
                active_filters[fc] = sel
            # =============== Apply Filters ===============
            filtered = df_long.copy()
            if primary_filter_col and primary_values is not None and len(primary_values) > 0:
                filtered = filtered[filtered[primary_filter_col].astype(str).isin(primary_values)]
            for fc, sel in active_filters.items():
                if sel is not None and len(sel) > 0:
                    filtered = filtered[filtered[fc].astype(str).isin(sel)]
            update_progress(70, "Building KPIs")
            # === Auto Group Low-Performers ===
            rep_col = _find_col(filtered, ["rep", "representative", "salesman", "employee", "name", "mr"])
            performance_group_col = None
            filtered_with_group = filtered.copy()
            if rep_col and kpi_measure_col and rep_col in filtered.columns and kpi_measure_col in filtered.columns:
                sales_by_rep = filtered.groupby(rep_col)[kpi_measure_col].sum().sort_values(ascending=False)
                total_reps = len(sales_by_rep)
                if total_reps >= 5:
                    high_idx = int(0.2 * total_reps)
                    low_idx = int(0.8 * total_reps)
                    high_reps = set(sales_by_rep.index[:high_idx])
                    low_reps = set(sales_by_rep.index[low_idx:])
                    def assign_group(rep):
                        if rep in high_reps:
                            return "High Performer"
                        elif rep in low_reps:
                            return "Needs Support"
                        else:
                            return "Medium Performer"
                    filtered_with_group['Performance Group'] = filtered_with_group[rep_col].apply(assign_group)
                    performance_group_col = 'Performance Group'
            final_df = filtered_with_group
            # === KPIs ===
            found_dims = {}
            for dim_key, aliases in {"area": ["area", "region"], "branch": ["branch", "location"], "rep": ["rep", "representative"]}.items():
                col = _find_col(final_df, aliases)
                if col:
                    found_dims[dim_key] = col
            kpi_values = {}
            if kpi_measure_col and kpi_measure_col in final_df.columns:
                kpi_values['total'] = final_df[kpi_measure_col].sum()
                date_cols = [c for c in final_df.columns if any(d in c.lower() for d in ["date", "month", "year", "day"])]
                if date_cols:
                    unique_dates = final_df[date_cols[0]].nunique()
                    kpi_values['avg_per_date'] = kpi_values['total'] / unique_dates if unique_dates > 0 else None
                else:
                    kpi_values['avg_per_date'] = None
            else:
                kpi_values['total'] = None
                kpi_values['avg_per_date'] = None
            for dim_key, col_name in found_dims.items():
                kpi_values[f'unique_{dim_key}'] = final_df[col_name].nunique()
            # Calculate growth as direct difference (not average)
            if period_comparison and '__pct_change__' in final_df.columns:
                # Use sum of last period minus sum of previous period
                col1_sum = final_df[period_comparison['col1']].sum()
                col2_sum = final_df[period_comparison['col2']].sum()
                if col1_sum != 0:
                    growth_pct = ((col2_sum - col1_sum) / col1_sum) * 100
                else:
                    growth_pct = 0
                kpi_values['growth_pct'] = growth_pct
            kpi_cards = []
            if kpi_values.get('total') is not None:
                kpi_cards.append({'title': f'Total {kpi_measure_col}', 'value': f"{kpi_values['total']:,.0f}", 'color': 'linear-gradient(135deg, #28a745, #85e085)', 'icon': '📈'})
            # Removed Average KPI as requested
            if kpi_values.get('avg_per_date') is not None:
                kpi_cards.append({'title': 'Monthly Avg', 'value': f"{kpi_values['avg_per_date']:,.0f}", 'color': 'linear-gradient(135deg, #17a2b8, #66d9b3)', 'icon': '📅'})
            if kpi_values.get('growth_pct') is not None:
                color = 'linear-gradient(135deg, #28a745, #85e085)' if kpi_values['growth_pct'] >= 0 else 'linear-gradient(135deg, #dc3545, #ff6b6b)'
                kpi_cards.append({'title': 'Growth', 'value': f"{kpi_values['growth_pct']:.1f}%", 'color': color, 'icon': '↗️' if kpi_values['growth_pct'] >= 0 else '↘️'})
            if kpi_values.get('unique_area') is not None:
                kpi_cards.append({'title': 'Number of Areas', 'value': f"{kpi_values['unique_area']}", 'color': 'linear-gradient(135deg, #6f42c1, #a779e9)', 'icon': '🌍'})
            if kpi_values.get('unique_rep') is not None:
                kpi_cards.append({'title': 'Number of Reps', 'value': f"{kpi_values['unique_rep']}", 'color': 'linear-gradient(135deg, #ffc107, #ff8a00)', 'icon': '👥'})
            if kpi_values.get('unique_branch') is not None:
                kpi_cards.append({'title': 'Number of Branches', 'value': f"{kpi_values['unique_branch']}", 'color': 'linear-gradient(135deg, #20c997, #66d9b3)', 'icon': '🏢'})
            if performance_group_col:
                num_needs_support = len(final_df[final_df['Performance Group'] == 'Needs Support'][rep_col].unique())
                kpi_cards.append({'title': 'Needs Support', 'value': f"{num_needs_support}", 'color': 'linear-gradient(135deg, #dc3545, #ff6b6b)', 'icon': '🆘'})
            st.markdown("### 🚀 KPIs")
            cols = st.columns(min(6, len(kpi_cards)))
            for i, card in enumerate(kpi_cards[:6]):
                with cols[i]:
                    st.markdown(f"""
                    <div class='kpi-card' style='background:{card['color']};'>
                        <div class='kpi-title'>{card['icon']} {card['title']}</div>
                        <div class='kpi-value'>{card['value']}</div>
                    </div>
                    """, unsafe_allow_html=True)
            update_progress(90, "Rendering charts")
            # === Auto Charts ===
            st.markdown("### 📊 Auto Charts")
            charts_buffers = []
            plotly_figs = []
            # Performance Group Chart
            if performance_group_col:
                try:
                    group_total = final_df.groupby('Performance Group')[kpi_measure_col].sum().reset_index()
                    color_map = {'High Performer': '#28a745', 'Medium Performer': '#ffc107', 'Needs Support': '#dc3545'}
                    fig = px.bar(group_total, x='Performance Group', y=kpi_measure_col, title="📊 Performance Groups - Total Sales",
                                 color='Performance Group', color_discrete_map=color_map, text=kpi_measure_col)
                    fig.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
                    fig.update_layout(margin=dict(t=40,b=20,l=10,r=10), template="plotly_white")
                    plotly_figs.append((fig, "Performance Groups"))
                except Exception as e:
                    st.warning(f"⚠️ Could not generate Performance Groups chart: {e}")
            # Top/Bottom Employees
            if rep_col and kpi_measure_col and rep_col in final_df.columns and kpi_measure_col in final_df.columns:
                rep_data = final_df.groupby(rep_col)[kpi_measure_col].sum()
                if len(rep_data) >= 10:
                    top10 = rep_data.sort_values(ascending=False).head(10)
                    df_top = top10.reset_index().rename(columns={kpi_measure_col: "value"})
                    df_top[rep_col] = df_top[rep_col].astype(str).str.strip()
                    fig_top = px.bar(df_top, x=rep_col, y="value", title="🥇 Top 10 Employees", text="value")
                    fig_top.update_layout(margin=dict(t=40,b=20,l=10,r=10), template="plotly_white")
                    fig_top.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
                    plotly_figs.append((fig_top, "Top 10 Employees"))
                    bottom10 = rep_data.sort_values(ascending=True).head(10)
                    df_bottom = bottom10.reset_index().rename(columns={kpi_measure_col: "value"})
                    df_bottom[rep_col] = df_bottom[rep_col].astype(str).str.strip()
                    fig_bottom = px.bar(df_bottom, x=rep_col, y="value", title="📉 Bottom 10 Employees", text="value")
                    fig_bottom.update_layout(margin=dict(t=40,b=20,l=10,r=10), template="plotly_white")
                    fig_bottom.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
                    plotly_figs.append((fig_bottom, "Bottom 10 Employees"))
            # Other charts
            possible_dims = [c for c in final_df.columns if c != kpi_measure_col and c not in date_cols and c != rep_col]
            chosen_dim = None
            for alias in ["area", "region", "branch", "product", "item"]:
                col = _find_col(final_df, [alias])
                if col:
                    chosen_dim = col
                    break
            if not chosen_dim and possible_dims:
                chosen_dim = min(possible_dims, key=lambda x: final_df[x].nunique(dropna=True))
            if chosen_dim and kpi_measure_col and chosen_dim in final_df.columns:
                try:
                    series = final_df.groupby(chosen_dim)[kpi_measure_col].sum().sort_values(ascending=False).head(10)
                    df_series = series.reset_index().rename(columns={kpi_measure_col: "value"})
                    df_series[chosen_dim] = df_series[chosen_dim].astype(str).str.strip()
                    fig_bar = px.bar(df_series, x=chosen_dim, y="value", title=f"Top by {chosen_dim}", text="value")
                    fig_bar.update_xaxes(type='category')
                    fig_bar.update_layout(margin=dict(t=40,b=20,l=10,r=10), template="plotly_white")
                    fig_bar.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
                    plotly_figs.append((fig_bar, f"Top by {chosen_dim}"))
                except Exception as e:
                    st.warning(f"⚠️ Could not generate chart for {chosen_dim}: {e}")
            # Display charts
            st.markdown("#### Dashboard — Charts (3 columns × up to 2 rows)")
            plotly_figs = plotly_figs[:6]
            while len(plotly_figs) < 6:
                plotly_figs.append((None, None))
            for i in range(0, 6, 3):
                cols = st.columns(3)
                for j in range(3):
                    if i+j < len(plotly_figs):
                        fig, caption = plotly_figs[i+j]
                        with cols[j]:
                            if fig is not None:
                                st.markdown('<div class="chart-card">', unsafe_allow_html=True)
                                st.plotly_chart(fig, use_container_width=True, theme="streamlit")
                                st.markdown(f'<div style="text-align:center; color:#FFD700; font-size:14px; margin-top:4px;">{caption}</div>', unsafe_allow_html=True)
                                st.markdown('</div>', unsafe_allow_html=True)
            update_progress(100, "Dashboard ready!")
            st.success("🎉 Dashboard generated successfully!")
            # === Export Section ===
            st.markdown("### 💾 Export Report / Data")
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Filtered_Data')
            excel_data = excel_buffer.getvalue()
            st.download_button(
                label="⬇️ Download Filtered Data (Excel)",
                data=excel_data,
                file_name=f"{_safe_name(sheet_title)}_Filtered.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            if st.button("📥 Generate Dashboard PDF (charts only)"):
                with st.spinner("Generating Dashboard PDF..."):
                    if LOTTIE_PDF:
                        st.markdown("<div style='text-align:center;'>", unsafe_allow_html=True)
                        st_lottie(LOTTIE_PDF, height=180, key="lottie_dashboard_pdf")
                        st.markdown("</div>", unsafe_allow_html=True)
                    try:
                        pdf_buffer = build_pdf(sheet_title, charts_buffers, include_table=False)
                        st.success("✅ Dashboard PDF ready.")
                        st.download_button(
                            label="⬇️ Download Dashboard PDF",
                            data=pdf_buffer,
                            file_name=f"{_safe_name(sheet_title)}_Dashboard.pdf",
                            mime="application/pdf"
                        )
                    except Exception as e:
                        st.error(f"❌ PDF generation failed: {e}")
            progress_bar.empty()
            status_text.empty()
        except Exception as e:
            st.error(f"❌ Error generating dashboard: {e}")
            if 'progress_bar' in locals():
                progress_bar.empty()
                status_text.empty()

# ------------------ Tab 4: Info ------------------
with tab4:
    st.markdown("""
    <div class='guide-title'>🎯 Welcome to a free tool provided by the company admin.!</div>
    <br>
    <h3 style='color:#FFD700;'>📌 How to Use</h3>
    <ol style='color:white; font-size:16px; line-height:1.6;'>
        <li><strong>Upload Excel/CSV File (Splitter/Merge)</strong>: ... </li>
        <li><strong>Merge Excel/CSV Files</strong>: ... </li>
        <li><strong>Convert Images to PDF</strong>: ... </li>
        <li><strong>Auto Dashboard Generator</strong>: 
            <ul>
                <li>Upload an Excel or CSV file for dashboard.</li>
                <li>Select the sheet (if Excel).</li>
                <li>Use the sidebar to apply filters.</li>
                <li>The dashboard will auto-generate KPIs and charts.</li>
                <li><strong>New:</strong> Auto Group Low-Performers.</li>
            </ul>
        </li>
    </ol>
    <br>
    <h3 style='color:#FFD700;'>💡 Tips</h3>
    <ul>
        <li>Performance grouping requires at least 5 representatives.</li>
    </ul>
    """, unsafe_allow_html=True)


