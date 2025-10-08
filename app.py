# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
import re
import os
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl import load_workbook, Workbook
# ====== إضافات للداش بورد والتقارير ======
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image as RLImage, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
# Plotly for modern interactive charts (on-screen only)
import plotly.express as px
import plotly.graph_objects as go
# PowerPoint export
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
# ------------------ إضافة PIL لتحويل الصور إلى PDF ------------------
from PIL import Image
# Initialize clear counter in session state
if 'clear_counter' not in st.session_state:
    st.session_state.clear_counter = 0
# ------------------ ربط بخط عربي جميل (Cairo) ------------------
st.markdown(
    '<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap" rel="stylesheet">',
    unsafe_allow_html=True
)
# ------------------ إعدادات الصفحة ------------------
st.set_page_config(
    page_title="Averroes Pharma File Splitter & Dashboard",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="collapsed"
)
# ------------------ إخفاء شعار Streamlit والفوتر ------------------
hide_default = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
"""
st.markdown(hide_default, unsafe_allow_html=True)
# ------------------ ستايل مخصص ------------------
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
        background: linear-gradient(135deg, #6a11cb 0%, #2575fc 100%);
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
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)
# ------------------ دالة لعرض أسماء الملفات بلون فاتح ------------------
def display_uploaded_files(file_list, file_type="Excel"):
    if file_list:
        st.markdown("### 📁 Uploaded Files:")
        for i, f in enumerate(file_list):
            st.markdown(
                f"<div style='background:#003366; color:white; padding:4px 8px; border-radius:4px; margin:2px 0; display:inline-block;'>"
                f"{i+1}. {f.name} ({f.size//1024} KB)</div>",
                unsafe_allow_html=True
            )
# ------------------ شريط التنقل العلوي ------------------
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
# ------------------ عرض اللوجو ------------------
logo_path = "logo.png"
if os.path.exists(logo_path):
    st.markdown('<div style="text-align:center; margin:20px 0;">', unsafe_allow_html=True)
    st.image(logo_path, width=200)
    st.markdown('</div>', unsafe_allow_html=True)
else:
    st.markdown('<div style="text-align:center; margin:20px 0; color:#FFD700; font-size:20px;">Averroes Pharma</div>', unsafe_allow_html=True)
# ------------------ معلومات المطور ------------------
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
# ------------------ العنوان ------------------
st.markdown("<h1 style='text-align:center; color:#FFD700;'>💊 Averroes Pharma File Splitter & Dashboard</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align:center; color:white;'>✂ Split, Merge, Image-to-PDF & Auto Dashboard Generator</h3>", unsafe_allow_html=True)
# ------------------ Utility functions ------------------
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
            slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout
            left = Inches(0.5)
            top = Inches(0.8)
            width = Inches(9)
            height = Inches(5)
            slide.shapes.add_picture(img_buf, left, top, width=width, height=height)
            # Add caption
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
# ------------------ قسم التقسيم (Splitter) ------------------
st.markdown("### ✂ Split Excel File")
uploaded_file = st.file_uploader(
    "📂 Upload Excel File (Splitter/Merge)",
    type=["xlsx"],
    accept_multiple_files=False,
    key=f"split_uploader_{st.session_state.clear_counter}"
)
if uploaded_file:
    display_uploaded_files([uploaded_file], "Excel")
    if st.button("🗑️ Clear Uploaded File", key="clear_split"):
        st.session_state.clear_counter += 1
        st.rerun()
    try:
        input_bytes = uploaded_file.getvalue()
        original_wb = load_workbook(filename=BytesIO(input_bytes), data_only=False)
        st.success(f"✅ The file has been uploaded successfully. Number of sheets: {len(original_wb.sheetnames)}")
        selected_sheet = st.selectbox("Select Sheet (for Split)", original_wb.sheetnames)
        if selected_sheet:
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
                help="اختر 'Split by Column Values' لتقسيم الشيت الحالي حسب قيم عمود. اختر 'Split Each Sheet into Separate File' لإنشاء ملف منفصل لكل شيت في الـ Workbook."
            )
            if st.button("🚀 Start Split"):
                with st.spinner("Splitting process in progress while preserving original format..."):
                    def clean_name(name):
                        name = str(name).strip()
                        invalid_chars = r'[\\/*?:\[\]|<>"]'
                        cleaned = re.sub(invalid_chars, '_', name)
                        return cleaned[:30] if cleaned else "Sheet"
                    if split_option == "Split by Column Values":
                        ws = original_wb[selected_sheet]
                        col_idx = df.columns.get_loc(col_to_split) + 1
                        unique_values = df[col_to_split].dropna().unique()
                        zip_buffer = BytesIO()
                        with ZipFile(zip_buffer, "w") as zip_file:
                            for value in unique_values:
                                new_wb = Workbook()
                                default_ws = new_wb.active
                                new_wb.remove(default_ws)
                                new_ws = new_wb.create_sheet(title=clean_name(value))
                                for cell in ws[1]:
                                    dst_cell = new_ws.cell(1, cell.column, cell.value)
                                    if cell.has_style:
                                        try:
                                            if cell.font:
                                                dst_cell.font = Font(
                                                    name=cell.font.name,
                                                    size=cell.font.size,
                                                    bold=cell.font.bold,
                                                    italic=cell.font.italic,
                                                    color=cell.font.color
                                                )
                                            if cell.fill and cell.fill.fill_type:
                                                dst_cell.fill = PatternFill(
                                                    fill_type=cell.fill.fill_type,
                                                    start_color=cell.fill.start_color,
                                                    end_color=cell.fill.end_color
                                                )
                                            if cell.border:
                                                dst_cell.border = Border(
                                                    left=cell.border.left,
                                                    right=cell.border.right,
                                                    top=cell.border.top,
                                                    bottom=cell.border.bottom
                                                )
                                            if cell.alignment:
                                                dst_cell.alignment = Alignment(
                                                    horizontal=cell.alignment.horizontal,
                                                    vertical=cell.alignment.vertical,
                                                    wrap_text=cell.alignment.wrap_text
                                                )
                                            dst_cell.number_format = cell.number_format
                                        except Exception:
                                            pass
                                row_idx = 2
                                for row in ws.iter_rows(min_row=2):
                                    cell_in_col = row[col_idx - 1]
                                    if cell_in_col.value == value:
                                        for src_cell in row:
                                            dst_cell = new_ws.cell(row_idx, src_cell.column, src_cell.value)
                                            if src_cell.has_style:
                                                try:
                                                    if src_cell.font:
                                                        dst_cell.font = Font(
                                                            name=src_cell.font.name,
                                                            size=src_cell.font.size,
                                                            bold=src_cell.font.bold,
                                                            italic=src_cell.font.italic,
                                                            color=src_cell.font.color
                                                        )
                                                    if src_cell.fill and src_cell.fill.fill_type:
                                                        dst_cell.fill = PatternFill(
                                                            fill_type=src_cell.fill.fill_type,
                                                            start_color=src_cell.fill.start_color,
                                                            end_color=src_cell.fill.end_color
                                                        )
                                                    if src_cell.border:
                                                        dst_cell.border = Border(
                                                            left=src_cell.border.left,
                                                            right=src_cell.border.right,
                                                            top=src_cell.border.top,
                                                            bottom=src_cell.border.bottom
                                                        )
                                                    if src_cell.alignment:
                                                        dst_cell.alignment = Alignment(
                                                            horizontal=src_cell.alignment.horizontal,
                                                            vertical=src_cell.alignment.vertical,
                                                            wrap_text=src_cell.alignment.wrap_text
                                                        )
                                                    dst_cell.number_format = src_cell.number_format
                                                except Exception:
                                                    pass
                                        row_idx += 1
                                try:
                                    for col_letter in ws.column_dimensions:
                                        new_ws.column_dimensions[col_letter].width = ws.column_dimensions[col_letter].width
                                except Exception:
                                    pass
                                file_buffer = BytesIO()
                                new_wb.save(file_buffer)
                                file_buffer.seek(0)
                                file_name = f"{clean_name(value)}.xlsx"
                                zip_file.writestr(file_name, file_buffer.read())
                                st.write(f"📁 Created file: `{value}`")
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
                        with ZipFile(zip_buffer, "w") as zip_file:
                            for sheet_name in original_wb.sheetnames:
                                new_wb = Workbook()
                                default_ws = new_wb.active
                                new_wb.remove(default_ws)
                                new_ws = new_wb.create_sheet(title=sheet_name)
                                src_ws = original_wb[sheet_name]
                                for row in src_ws.iter_rows():
                                    for src_cell in row:
                                        dst_cell = new_ws.cell(src_cell.row, src_cell.column, src_cell.value)
                                        if src_cell.has_style:
                                            try:
                                                if src_cell.font:
                                                    dst_cell.font = Font(
                                                        name=src_cell.font.name,
                                                        size=src_cell.font.size,
                                                        bold=src_cell.font.bold,
                                                        italic=src_cell.font.italic,
                                                        color=src_cell.font.color
                                                    )
                                                if src_cell.fill and src_cell.fill.fill_type:
                                                    dst_cell.fill = PatternFill(
                                                        fill_type=src_cell.fill.fill_type,
                                                        start_color=src_cell.fill.start_color,
                                                        end_color=src_cell.fill.end_color
                                                    )
                                                if src_cell.border:
                                                    dst_cell.border = Border(
                                                        left=src_cell.border.left,
                                                        right=src_cell.border.right,
                                                        top=src_cell.border.top,
                                                        bottom=src_cell.border.bottom
                                                    )
                                                if src_cell.alignment:
                                                    dst_cell.alignment = Alignment(
                                                        horizontal=src_cell.alignment.horizontal,
                                                        vertical=src_cell.alignment.vertical,
                                                        wrap_text=src_cell.alignment.wrap_text,
                                                        indent=src_cell.alignment.indent
                                                    )
                                                dst_cell.number_format = src_cell.number_format
                                            except Exception:
                                                pass
                                if src_ws.merged_cells.ranges:
                                    for merged_range in src_ws.merged_cells.ranges:
                                        new_ws.merge_cells(str(merged_range))
                                        top_left_cell = src_ws.cell(merged_range.min_row, merged_range.min_col)
                                        merged_value = top_left_cell.value
                                        new_ws.cell(merged_range.min_row, merged_range.min_col, merged_value)
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
                                st.write(f"📁 Created file: `{sheet_name}`")
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
# -----------------------------------------------
# Merge area
# -----------------------------------------------
st.markdown("<hr class='divider-dashed'>", unsafe_allow_html=True)
st.markdown("### 🔄 Merge Excel Files (Keep Original Format & Merged Cells)")
merge_files = st.file_uploader(
    "📤 Upload Excel Files to Merge",
    type=["xlsx"],
    accept_multiple_files=True,
    key=f"merge_uploader_{st.session_state.clear_counter}"
)
if merge_files:
    display_uploaded_files(merge_files, "Excel")
    if st.button("🗑️ Clear All Merged Files", key="clear_merge"):
        st.session_state.clear_counter += 1
        st.rerun()
    if st.button("✨ Merge Files with Format"):
        with st.spinner("Merging files while preserving formatting and merged cells..."):
            try:
                combined_wb = Workbook()
                combined_ws = combined_wb.active
                combined_ws.title = "Consolidated"
                first_file = merge_files[0]
                first_wb = load_workbook(filename=BytesIO(first_file.getvalue()), data_only=False)
                first_ws = first_wb.active
                for cell in first_ws[1]:
                    new_cell = combined_ws.cell(1, cell.column, cell.value)
                    if cell.has_style:
                        try:
                            if cell.font:
                                new_cell.font = Font(
                                    name=cell.font.name,
                                    size=cell.font.size,
                                    bold=cell.font.bold,
                                    italic=cell.font.italic,
                                    color=cell.font.color
                                )
                            if cell.fill and cell.fill.fill_type:
                                new_cell.fill = PatternFill(
                                    fill_type=cell.fill.fill_type,
                                    start_color=cell.fill.start_color,
                                    end_color=cell.fill.end_color
                                )
                            if cell.border:
                                new_cell.border = Border(
                                    left=cell.border.left,
                                    right=cell.border.right,
                                    top=cell.border.top,
                                    bottom=cell.border.bottom
                                )
                            if cell.alignment:
                                new_cell.alignment = Alignment(
                                    horizontal=cell.alignment.horizontal,
                                    vertical=cell.alignment.vertical,
                                    wrap_text=cell.alignment.wrap_text
                                )
                            new_cell.number_format = cell.number_format
                        except Exception:
                            pass
                if first_ws.merged_cells.ranges:
                    for merged_range in first_ws.merged_cells.ranges:
                        combined_ws.merge_cells(str(merged_range))
                        top_left_cell = first_ws.cell(merged_range.min_row, merged_range.min_col)
                        combined_ws.cell(merged_range.min_row, merged_range.min_col, top_left_cell.value)
                try:
                    for col_letter in first_ws.column_dimensions:
                        combined_ws.column_dimensions[col_letter].width = first_ws.column_dimensions[col_letter].width
                except Exception:
                    pass
                row_idx = 2
                for file in merge_files:
                    wb = load_workbook(filename=BytesIO(file.getvalue()), data_only=True)
                    ws = wb.active
                    for row in ws.iter_rows(min_row=2):
                        for cell in row:
                            if cell.value is not None:
                                new_cell = combined_ws.cell(row_idx, cell.column, cell.value)
                                if cell.has_style:
                                    try:
                                        if cell.font:
                                            new_cell.font = Font(
                                                name=cell.font.name,
                                                size=cell.font.size,
                                                bold=cell.font.bold,
                                                italic=cell.font.italic,
                                                color=cell.font.color
                                            )
                                        if cell.fill and cell.fill.fill_type:
                                            new_cell.fill = PatternFill(
                                                fill_type=cell.fill.fill_type,
                                                start_color=cell.fill.start_color,
                                                end_color=cell.fill.end_color
                                            )
                                        if cell.border:
                                            new_cell.border = Border(
                                                left=cell.border.left,
                                                right=cell.border.right,
                                                top=cell.border.top,
                                                bottom=cell.border.bottom
                                            )
                                        if cell.alignment:
                                            new_cell.alignment = Alignment(
                                                horizontal=cell.alignment.horizontal,
                                                vertical=cell.alignment.vertical,
                                                wrap_text=cell.alignment.wrap_text
                                            )
                                        new_cell.number_format = cell.number_format
                                    except Exception:
                                        pass
                        row_idx += 1
                output_buffer = BytesIO()
                combined_wb.save(output_buffer)
                output_buffer.seek(0)
                st.success("✅ Merged successfully with full format preserved!")
                st.download_button(
                    label="📥 Download Merged File (with Format)",
                    data=output_buffer.getvalue(),
                    file_name="Merged_Consolidated_With_Format.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ Error during merge: {e}")
# ====================================================================================
# 📷 Image to PDF Converter with CamScanner Effect
# ====================================================================================
st.markdown("<hr class='divider'>", unsafe_allow_html=True)
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
    # --- دالة لتحسين الصورة مثل CamScanner ---
    try:
        import cv2
        import numpy as np
        def enhance_image_for_pdf(image_pil):
            """تحسين الصورة لتكون مثل ما يفعله CamScanner: اكتشاف الحواف، القص، تحسين الألوان"""
            # تحويل PIL إلى OpenCV
            image = np.array(image_pil)
            if image.shape[2] == 4:  # RGBA
                image = cv2.cvtColor(image, cv2.COLOR_RGBA2RGB)
            original_color = image.copy()

            # --- 1. اكتشاف الحواف وقص الورقة (Document Detection) ---
            def find_document_contour(img):
                gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
                blur = cv2.GaussianBlur(gray, (5,5), 0)
                # استخدام Canny مع حدود ديناميكية
                v = np.median(gray)
                lower = int(max(0, (1.0 - 0.33) * v))
                upper = int(min(255, (1.0 + 0.33) * v))
                edged = cv2.Canny(blur, lower, upper)

                contours, _ = cv2.findContours(edged.copy(), cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
                contours = sorted(contours, key=cv2.contourArea, reverse=True)[:10]

                for c in contours:
                    peri = cv2.arcLength(c, True)
                    approx = cv2.approxPolyDP(c, 0.02 * peri, True)
                    if len(approx) == 4:
                        return approx
                return None

            doc_contour = find_document_contour(original_color)
            if doc_contour is not None:
                # حساب الإحداثيات للقص
                pts = doc_contour.reshape(4, 2)
                rect = np.zeros((4, 2), dtype="float32")
                s = pts.sum(axis=1)
                rect[0] = pts[np.argmin(s)]   # top-left
                rect[2] = pts[np.argmax(s)]   # bottom-right
                diff = np.diff(pts, axis=1)
                rect[1] = pts[np.argmin(diff)]  # top-right
                rect[3] = pts[np.argmax(diff)]  # bottom-left

                # حساب الأبعاد الجديدة (عرض وارتفاع)
                widthA = np.sqrt(((rect[2][0] - rect[3][0]) ** 2) + ((rect[2][1] - rect[3][1]) ** 2))
                widthB = np.sqrt(((rect[1][0] - rect[0][0]) ** 2) + ((rect[1][1] - rect[0][1]) ** 2))
                maxWidth = max(int(widthA), int(widthB))

                heightA = np.sqrt(((rect[1][0] - rect[2][0]) ** 2) + ((rect[1][1] - rect[2][1]) ** 2))
                heightB = np.sqrt(((rect[0][0] - rect[3][0]) ** 2) + ((rect[0][1] - rect[3][1]) ** 2))
                maxHeight = max(int(heightA), int(heightB))

                dst = np.array([
                    [0, 0],
                    [maxWidth - 1, 0],
                    [maxWidth - 1, maxHeight - 1],
                    [0, maxHeight - 1]], dtype="float32")

                M = cv2.getPerspectiveTransform(rect, dst)
                warped = cv2.warpPerspective(original_color, M, (maxWidth, maxHeight))
                image = warped  # نستبدل الصورة الأصلية بالصورة المقطوعة والمعدلة

            # --- 2. تحسين الألوان والتباين (بدون تحويل للرمادي) ---
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
            if len(image.shape) == 3 and image.shape[2] == 3:
                channels = cv2.split(image)
                enhanced_channels = []
                for channel in channels:
                    enhanced = clahe.apply(channel)
                    enhanced_channels.append(enhanced)
                image = cv2.merge(enhanced_channels)

            # --- 3. إضافة إطار أبيض (اختياري) ---
            border_size = 20
            bordered = cv2.copyMakeBorder(
                image,
                top=border_size,
                bottom=border_size,
                left=border_size,
                right=border_size,
                borderType=cv2.BORDER_CONSTANT,
                value=[255, 255, 255]
            )

            # --- 4. التأكد من العمق (uint8) ---
            if bordered.dtype != np.uint8:
                bordered = np.clip(bordered, 0, 255).astype(np.uint8)

            # --- 5. تحويل إلى PIL (RGB) ---
            result = cv2.cvtColor(bordered, cv2.COLOR_BGR2RGB)
            return Image.fromarray(result)

        if st.button("🖨️ Create PDF (CamScanner Style)"):
            with st.spinner("Enhancing images for PDF..."):
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
    # --- الزر العادي (بدون تحسين) ---
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
# ====================================================================================
# 📊 Dashboard Generator
# ====================================================================================
st.markdown("<hr class='divider' id='dashboard-section'>", unsafe_allow_html=True)
st.markdown("### 📊 Interactive Auto Dashboard Generator")
dashboard_file = st.file_uploader(
    "📊 Upload Excel File for Dashboard (Auto)",
    type=["xlsx"],
    key=f"dashboard_uploader_{st.session_state.clear_counter}"
)
if dashboard_file:
    display_uploaded_files([dashboard_file], "Excel")
    if st.button("🗑️ Clear Dashboard File", key="clear_dashboard"):
        st.session_state.clear_counter += 1
        st.rerun()
    try:
        df_dict = pd.read_excel(dashboard_file, sheet_name=None)
        sheet_names = list(df_dict.keys())
        selected_sheet_dash = st.selectbox("Select Sheet for Dashboard", sheet_names, key="sheet_dash")
        if selected_sheet_dash:
            sheet_title = selected_sheet_dash
            df0 = df_dict[selected_sheet_dash].copy()
            st.markdown("### 🔍 Data Preview (original)")
            st.dataframe(df0.head(), use_container_width=True)
            month_names = ["jan","feb","mar","apr","may","jun","jul","aug","sep","sept","oct","nov","dec"]
            cols_lower = [c.strip().lower() for c in df0.columns]
            potential_months = [c for c in df0.columns if c.strip().lower() in month_names]
            numeric_cols = df0.select_dtypes(include='number').columns.tolist()
            if potential_months:
                id_vars = [c for c in df0.columns if c not in potential_months]
                value_vars = potential_months
                df_long = df0.melt(id_vars=id_vars, value_vars=value_vars, var_name="Month", value_name="Value")
                df_long["Month"] = df_long["Month"].astype(str)
                measure_col = "Value"
            else:
                if len(numeric_cols) == 1:
                    measure_col = numeric_cols[0]
                    df_long = df0.copy()
                elif len(numeric_cols) > 1:
                    df_long = df0.copy()
                    df_long["__auto_sales__"] = df_long[numeric_cols].sum(axis=1, numeric_only=True)
                    measure_col = "__auto_sales__"
                else:
                    measure_col = None
                    df_long = df0.copy()
            cat_cols = [c for c in df_long.columns if df_long[c].dtype == "object" or df_long[c].dtype.name.startswith("category")]
            for c in df_long.columns:
                if c not in cat_cols and df_long[c].nunique(dropna=True) <= 100 and df_long[c].dtype != "float64" and df_long[c].dtype != "int64":
                    cat_cols.append(c)
            cat_cols = [c for c in cat_cols if c is not None]
            st.sidebar.header("🔍 Dynamic Filters")
            primary_filter_col = None
            if len(cat_cols) > 0:
                primary_filter_col = st.sidebar.selectbox("Primary Filter Column (drop-list)", ["-- None --"] + cat_cols, index=0)
                if primary_filter_col == "-- None --":
                    primary_filter_col = None
            primary_values = None
            if primary_filter_col:
                vals = df_long[primary_filter_col].dropna().astype(str).unique().tolist()
                try:
                    vals = sorted(vals)
                except Exception:
                    pass
                primary_values = st.sidebar.multiselect(f"Filter values for {primary_filter_col}", vals, default=vals)
            other_filter_cols = st.sidebar.multiselect("Choose additional filter columns (optional)", [c for c in cat_cols if c != primary_filter_col], default=[])
            active_filters = {}
            for fc in other_filter_cols:
                opts = df_long[fc].dropna().astype(str).unique().tolist()
                try:
                    opts = sorted(opts)
                except Exception:
                    pass
                sel = st.sidebar.multiselect(f"Filter: {fc}", opts, default=opts)
                active_filters[fc] = sel
            filtered = df_long.copy()
            if primary_filter_col and primary_values is not None:
                if len(primary_values) > 0:
                    filtered = filtered[filtered[primary_filter_col].astype(str).isin(primary_values)]
            for fc, sel in active_filters.items():
                if sel is not None and len(sel) > 0:
                    filtered = filtered[filtered[fc].astype(str).isin(sel)]
            st.markdown("### 📈 Filtered Data Preview")
            st.dataframe(filtered.head(200), use_container_width=True)
            # === KPIs with Icons & Gradients ===
            kpi_measure_col = measure_col
            possible_measure_aliases = ["sales", "amount", "value", "total", "revenue", "target", "achievement"]
            for alias in possible_measure_aliases:
                col = _find_col(filtered, [alias])
                if col:
                    kpi_measure_col = col
                    break
            possible_dim_aliases = {
                "area": ["area", "region", "territory"],
                "branch": ["branch", "location", "store"],
                "rep": ["rep", "representative", "salesman", "employee", "name"]
            }
            found_dims = {}
            for dim_key, aliases in possible_dim_aliases.items():
                col = _find_col(filtered, aliases)
                if col:
                    found_dims[dim_key] = col
            kpi_values = {}
            if kpi_measure_col and kpi_measure_col in filtered.columns:
                kpi_values['total'] = filtered[kpi_measure_col].sum()
                kpi_values['avg'] = filtered[kpi_measure_col].mean()
                if "Month" in filtered.columns:
                    unique_dates = filtered["Month"].nunique()
                    if unique_dates > 0:
                        kpi_values['avg_per_date'] = kpi_values['total'] / unique_dates
                    else:
                        kpi_values['avg_per_date'] = None
                else:
                    kpi_values['avg_per_date'] = None
            else:
                kpi_values['total'] = None
                kpi_values['avg'] = None
                kpi_values['avg_per_date'] = None
            for dim_key, col_name in found_dims.items():
                kpi_values[f'unique_{dim_key}'] = filtered[col_name].nunique()
            kpi_cards = []
            if kpi_values.get('total') is not None:
                kpi_cards.append({
                    'title': f'إجمالي {kpi_measure_col}',
                    'value': f"{kpi_values['total']:,.0f}",
                    'color': 'linear-gradient(135deg, #ff8a00, #ffc107)',
                    'icon': '💰'
                })
            if kpi_values.get('avg') is not None:
                kpi_cards.append({
                    'title': f'Average {kpi_measure_col}',
                    'value': f"{kpi_values['avg']:,.0f}",
                    'color': 'linear-gradient(135deg, #00c0ff, #007bff)',
                    'icon': '📊'
                })
            if kpi_values.get('avg_per_date') is not None:
                kpi_cards.append({
                    'title': f'Average per Date',
                    'value': f"{kpi_values['avg_per_date']:,.0f}",
                    'color': 'linear-gradient(135deg, #28a745, #85e085)',
                    'icon': '📅'
                })
            if kpi_values.get('unique_area') is not None:
                kpi_cards.append({
                    'title': f'عدد المناطق',
                    'value': f"{kpi_values['unique_area']}",
                    'color': 'linear-gradient(135deg, #6f42c1, #a779e9)',
                    'icon': '🌍'
                })
            if kpi_values.get('unique_rep') is not None:
                kpi_cards.append({
                    'title': f'Count Mr ',
                    'value': f"{kpi_values['unique_rep']}",
                    'color': 'linear-gradient(135deg, #dc3545, #ff6b6b)',
                    'icon': '👨‍💼'
                })
            if kpi_values.get('unique_branch') is not None:
                kpi_cards.append({
                    'title': f'عدد الفروع',
                    'value': f"{kpi_values['unique_branch']}",
                    'color': 'linear-gradient(135deg, #20c997, #66d9b3)',
                    'icon': '🏢'
                })
            st.markdown("### 🚀 KPIs")
            cols = st.columns(min(6, len(kpi_cards)))
            for i, card in enumerate(kpi_cards[:6]):
                with cols[i]:
                    kpi_html = f"""
                    <div class='kpi-card' style='background:{card['color']};'>
                        <div class='kpi-title'>{card['icon']} {card['title']}</div>
                        <div class='kpi-value'>{card['value']}</div>
                    </div>
                    """
                    st.markdown(kpi_html, unsafe_allow_html=True)
            # === Auto Charts ===
            st.markdown("### 📊 Auto Charts (built from data)")
            charts_buffers = []
            plotly_figs = []
            possible_dims = [c for c in filtered.columns if c not in [measure_col, "Month"]]
            pie_prefer_order = ["area", "region", "territory", "branch", "location", "city"]
            bar_prefer_order = ["rep", "representative", "salesman", "employee", "name", "item", "product", "sku"]
            pie_dim = None
            bar_dim = None
            for p in pie_prefer_order:
                for c in possible_dims:
                    if p in c.lower():
                        pie_dim = c
                        break
                if pie_dim:
                    break
            for p in bar_prefer_order:
                for c in possible_dims:
                    if p in c.lower():
                        bar_dim = c
                        break
                if bar_dim:
                    break
            if not pie_dim and not bar_dim and len(possible_dims) > 0:
                lens = [(c, filtered[c].nunique(dropna=True)) for c in possible_dims]
                lens = sorted([x for x in lens if x[1] > 1], key=lambda x: x[1])
                if lens:
                    pie_dim = lens[0][0]
                    bar_dim = lens[-1][0] if len(lens) > 1 else pie_dim
            elif pie_dim and not bar_dim:
                bar_dim = pie_dim
            elif bar_dim and not pie_dim:
                pie_dim = bar_dim
            chosen_dim = bar_dim
            dims_for_charts = []
            if chosen_dim:
                dims_for_charts.append(chosen_dim)
            remaining = [c for c in possible_dims if c not in dims_for_charts]
            rem_sorted = sorted(remaining, key=lambda x: filtered[x].nunique(dropna=True))
            for r in rem_sorted[:4]:
                dims_for_charts.append(r)
            dims_for_charts = dims_for_charts[:5]
            # Chart A: Bar
            if chosen_dim and measure_col and chosen_dim in filtered.columns:
                try:
                    series = filtered.groupby(chosen_dim)[measure_col].sum().sort_values(ascending=False).head(10)
                    df_series = series.reset_index().rename(columns={measure_col: "value"})
                    df_series[chosen_dim] = df_series[chosen_dim].astype(str).str.strip()
                    fig_bar = px.bar(df_series, x=chosen_dim, y="value", title=f"Top by {chosen_dim}", text="value")
                    fig_bar.update_xaxes(type='category')
                    fig_bar.update_layout(margin=dict(t=40,b=20,l=10,r=10), template="plotly_white")
                    fig_bar.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
                    plotly_figs.append((fig_bar, f"Top by {chosen_dim}"))
                    fig_m, ax = plt.subplots(figsize=(10, 5))
                    x_labels = series.index.astype(str).str.strip()
                    bars = ax.bar(x_labels, series.values)
                    ax.set_title(f"Top by {chosen_dim}", fontsize=14, fontweight='bold')
                    ax.yaxis.set_major_formatter(FuncFormatter(_format_millions))
                    ax.tick_params(axis='x', rotation=45, labelsize=10)
                    for label in ax.get_xticklabels():
                        label.set_ha('right')
                    ax.set_xlabel(chosen_dim, fontsize=10, fontweight='bold')
                    for b in bars:
                        h = b.get_height()
                        if pd.isna(h):
                            continue
                        ax.annotate(f"{h:,.0f}", 
                                    xy=(b.get_x() + b.get_width()/2, h),
                                    xytext=(0, 5), 
                                    textcoords="offset points",
                                    ha='center', va='bottom',
                                    fontsize=9, fontweight='bold')
                    fig_m.tight_layout()
                    img_buf = BytesIO()
                    fig_m.savefig(img_buf, format="png", dpi=200, bbox_inches="tight")
                    img_buf.seek(0)
                    charts_buffers.append((img_buf, f"Top by {chosen_dim}"))
                    plt.close(fig_m)
                except Exception as e:
                    st.warning(f"⚠️ Could not generate chart for {chosen_dim}: {e}")
            # Chart B: Pie
            if len(dims_for_charts) >= 2 and measure_col:
                dim2 = dims_for_charts[1]
                try:
                    series2 = filtered.groupby(dim2)[measure_col].sum().sort_values(ascending=False).head(10)
                    df_pie = series2.reset_index().rename(columns={measure_col: "value"})
                    fig_pie = px.pie(df_pie, names=dim2, values="value", title=f"Share by {dim2}", hole=0.4)
                    fig_pie.update_traces(
                        textposition='inside',
                        textinfo='percent+label',
                        insidetextorientation='radial',
                        textfont_size=12
                    )
                    fig_pie.update_layout(margin=dict(t=40,b=20,l=10,r=10), template="plotly_white")
                    plotly_figs.append((fig_pie, f"Share by {dim2}"))
                    fig_m, ax = plt.subplots(figsize=(8, 8))
                    wedges, texts, autotexts = ax.pie(
                        series2.values,
                        labels=series2.index.astype(str),
                        autopct=lambda pct: f"{pct:.1f}%",
                        startangle=90,
                        textprops={'fontsize': 10}
                    )
                    for text in texts:
                        text.set_rotation(30)
                    ax.set_title(f"Share by {dim2}", fontsize=14, fontweight='bold')
                    ax.axis('equal')
                    fig_m.tight_layout()
                    img_buf = BytesIO()
                    fig_m.savefig(img_buf, format="png", dpi=200, bbox_inches="tight")
                    img_buf.seek(0)
                    charts_buffers.append((img_buf, f"Share by {dim2}"))
                    plt.close(fig_m)
                except Exception as e:
                    st.warning(f"⚠️ Could not generate pie chart for {dim2}: {e}")
            else:
                if chosen_dim and measure_col:
                    try:
                        s = filtered.groupby(chosen_dim)[measure_col].sum().sort_values(ascending=False).head(8)
                        df_pie = s.reset_index().rename(columns={measure_col: "value"})
                        fig_pie = px.pie(df_pie, names=chosen_dim, values="value", title=f"Share by {chosen_dim}", hole=0.4)
                        fig_pie.update_traces(
                            textposition='inside',
                            textinfo='percent+label',
                            insidetextorientation='radial',
                            textfont_size=12
                        )
                        fig_pie.update_layout(margin=dict(t=40,b=20,l=10,r=10), template="plotly_white")
                        plotly_figs.append((fig_pie, f"Share by {chosen_dim}"))
                        fig_m, ax = plt.subplots(figsize=(8, 8))
                        wedges, texts, autotexts = ax.pie(
                            s.values,
                            labels=s.index.astype(str),
                            autopct=lambda pct: f"{pct:.1f}%",
                            startangle=90,
                            textprops={'fontsize': 10}
                        )
                        for text in texts:
                            text.set_rotation(30)
                        ax.set_title(f"Share by {chosen_dim}", fontsize=14, fontweight='bold')
                        ax.axis('equal')
                        fig_m.tight_layout()
                        img_buf = BytesIO()
                        fig_m.savefig(img_buf, format="png", dpi=200, bbox_inches="tight")
                        img_buf.seek(0)
                        charts_buffers.append((img_buf, f"Share by {chosen_dim}"))
                        plt.close(fig_m)
                    except Exception as e:
                        st.warning(f"⚠️ Could not generate fallback pie chart: {e}")
            # Chart C: Trend (Line) - ONLY if Month exists
            if "Month" in filtered.columns and measure_col and measure_col in filtered.columns:
                try:
                    ser = filtered.dropna(subset=["Month"])
                    if pd.api.types.is_datetime64_any_dtype(ser["Month"]):
                        ser["_yyyymm"] = ser["Month"].dt.to_period("M")
                        trend = ser.groupby("_yyyymm")[measure_col].sum().sort_index()
                        trend = trend.reset_index().rename(columns={measure_col: "value"})
                        trend["_yyyymm"] = trend["_yyyymm"].astype(str)
                        fig_line = px.line(trend, x="_yyyymm", y="value", markers=True, title="Trend by Month")
                        fig_line.update_traces(texttemplate='%{y:,.0f}', textposition='top center')
                        fig_line.update_layout(xaxis_title="Month", yaxis_title="Total", margin=dict(t=40,b=20,l=10,r=10), template="plotly_white")
                    else:
                        trend = ser.groupby("Month")[measure_col].sum().sort_index().reset_index().rename(columns={measure_col: "value"})
                        fig_line = px.line(trend, x="Month", y="value", markers=True, title="Trend by Month")
                        fig_line.update_traces(texttemplate='%{y:,.0f}', textposition='top center')
                        fig_line.update_layout(xaxis_title="Month", yaxis_title="Total", margin=dict(t=40,b=20,l=10,r=10), template="plotly_white")
                    plotly_figs.append((fig_line, "Trend by Month"))
                    fig_m, ax = plt.subplots(figsize=(10, 5))
                    ax.plot(trend.iloc[:, 0], trend["value"], marker='o')
                    ax.set_title("Trend by Month", fontsize=14, fontweight='bold')
                    ax.set_xlabel("Month")
                    ax.set_ylabel("Total")
                    ax.grid(True, alpha=0.3)
                    ax.yaxis.set_major_formatter(FuncFormatter(_format_millions))
                    for label in ax.get_xticklabels():
                        label.set_ha('right')
                    fig_m.tight_layout()
                    img_buf = BytesIO()
                    fig_m.savefig(img_buf, format="png", dpi=200, bbox_inches="tight")
                    img_buf.seek(0)
                    charts_buffers.append((img_buf, "Trend by Month"))
                    plt.close(fig_m)
                except Exception as e:
                    st.warning(f"⚠️ Could not generate trend chart: {e}")
            # Extra bars
            extra_dims = dims_for_charts[2:] if len(dims_for_charts) > 2 else []
            for ex_dim in extra_dims:
                if measure_col and ex_dim in filtered.columns:
                    try:
                        s = filtered.groupby(ex_dim)[measure_col].sum().sort_values(ascending=False).head(8)
                        dfe = s.reset_index().rename(columns={measure_col: "value"})
                        fig_extra = px.bar(dfe, x=ex_dim, y="value", title=f"By {ex_dim}", text="value")
                        fig_extra.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
                        fig_extra.update_layout(margin=dict(t=40,b=20,l=10,r=10), template="plotly_white")
                        plotly_figs.append((fig_extra, f"By {ex_dim}"))
                        fig_m, ax = plt.subplots(figsize=(9,4))
                        bars = ax.bar(s.index.astype(str), s.values)
                        ax.set_title(f"By {ex_dim}", fontsize=12, fontweight='bold')
                        ax.yaxis.set_major_formatter(FuncFormatter(_format_millions))
                        ax.tick_params(axis='x', rotation=45, labelsize=9)
                        for label in ax.get_xticklabels():
                            label.set_ha('right')
                        for b in bars:
                            h = b.get_height()
                            if pd.isna(h):
                                continue
                            ax.annotate(f"{h:,.0f}", xy=(b.get_x()+b.get_width()/2, h), xytext=(0,5), textcoords="offset points", ha='center', va='bottom', fontsize=8)
                        fig_m.tight_layout()
                        img_buf = BytesIO()
                        fig_m.savefig(img_buf, format="png", dpi=200, bbox_inches="tight")
                        img_buf.seek(0)
                        charts_buffers.append((img_buf, f"By {ex_dim}"))
                        plt.close(fig_m)
                    except Exception as e:
                        st.warning(f"⚠️ Could not generate chart for {ex_dim}: {e}")
            # Histogram - Only if we have numeric data AND no time dimension
            if measure_col and measure_col in filtered.columns and "Month" not in filtered.columns:
                try:
                    series_vals = filtered[measure_col].dropna().astype(float)
                    if len(series_vals) > 0:
                        df_hist = series_vals.to_frame(name="value")
                        fig_hist = px.histogram(df_hist, x="value", nbins=12, title="Distribution of Measure")
                        fig_hist.update_layout(margin=dict(t=40,b=20,l=10,r=10), template="plotly_white")
                        plotly_figs.append((fig_hist, "Distribution of Measure"))
                        fig_m, ax = plt.subplots(figsize=(8,3))
                        ax.hist(series_vals, bins=12, edgecolor='black')
                        ax.set_title("Distribution of Measure", fontsize=12, fontweight='bold')
                        ax.set_xlabel(measure_col)
                        ax.set_ylabel("Frequency")
                        ax.grid(True, alpha=0.3)
                        fig_m.tight_layout()
                        img_buf = BytesIO()
                        fig_m.savefig(img_buf, format="png", dpi=200, bbox_inches="tight")
                        img_buf.seek(0)
                        charts_buffers.append((img_buf, "Distribution of Measure"))
                        plt.close(fig_m)
                except Exception as e:
                    st.warning(f"⚠️ Could not generate distribution chart: {e}")
            # === Display Charts in Cards ===
            st.markdown("#### Dashboard — Charts (3 columns × up to 2 rows)")
            plotly_figs = plotly_figs[:6]
            while len(plotly_figs) < 6:
                plotly_figs.append((None, None))
            cols_row1 = st.columns(3)
            for i in range(3):
                fig, caption = plotly_figs[i]
                with cols_row1[i]:
                    if fig is not None:
                        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
                        st.plotly_chart(fig, use_container_width=True, theme="streamlit")
                        st.markdown(f'<div style="text-align:center; color:#FFD700; font-size:14px; margin-top:4px;">{caption}</div>', unsafe_allow_html=True)
                        st.markdown('</div>', unsafe_allow_html=True)
                    else:
                        st.write("")
            cols_row2 = st.columns(3)
            for i in range(3,6):
                fig, caption = plotly_figs[i]
                with cols_row2[i-3]:
                    if fig is not None:
                        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
                        st.plotly_chart(fig, use_container_width=True, theme="streamlit")
                        st.markdown(f'<div style="text-align:center; color:#FFD700; font-size:14px; margin-top:4px;">{caption}</div>', unsafe_allow_html=True)
                        st.markdown('</div>', unsafe_allow_html=True)
                    else:
                        st.write("")
            # === Export Section ===
            st.markdown("### 💾 Export Report / Data")
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                filtered.to_excel(writer, index=False, sheet_name='Filtered_Data')
            excel_data = excel_buffer.getvalue()
            st.download_button(
                label="⬇️ Download Filtered Data (Excel)",
                data=excel_data,
                file_name=f"{_safe_name(sheet_title)}_Filtered.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            if st.button("📥 Generate Dashboard PDF (charts only)"):
                with st.spinner("Generating Dashboard PDF (charts only)..."):
                    try:
                        pdf_buffer = build_pdf(sheet_title, charts_buffers=charts_buffers, include_table=False, filtered_df=None)
                        st.success("✅ Dashboard PDF جاهز.")
                        st.download_button(
                            label="⬇️ Download Dashboard PDF",
                            data=pdf_buffer,
                            file_name=f"{_safe_name(sheet_title)}_Dashboard.pdf",
                            mime="application/pdf"
                        )
                    except Exception as e:
                        st.error(f"❌ PDF generation failed: {e}")
            if st.checkbox("Include table in PDF report (optional)"):
                if st.button("📥 Generate Full PDF Report (charts + table)"):
                    with st.spinner("Generating full PDF..."):
                        try:
                            pdf_buffer = build_pdf(sheet_title, charts_buffers=charts_buffers, include_table=True, filtered_df=filtered)
                            st.success("✅ Full PDF جاهز.")
                            st.download_button(
                                label="⬇️ Download Full PDF (charts + table)",
                                data=pdf_buffer,
                                file_name=f"{_safe_name(sheet_title)}_FullReport.pdf",
                                mime="application/pdf"
                            )
                        except Exception as e:
                            st.error(f"❌ PDF generation failed: {e}")
            # === PowerPoint Export ===
            if st.button("📤 Export Dashboard to PowerPoint (PPTX)"):
                with st.spinner("Generating PowerPoint..."):
                    try:
                        pptx_buffer = build_pptx(sheet_title, charts_buffers)
                        st.success("✅ PowerPoint جاهز.")
                        st.download_button(
                            label="⬇️ Download Dashboard PowerPoint",
                            data=pptx_buffer.getvalue(),
                            file_name=f"{_safe_name(sheet_title)}_Dashboard.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                    except Exception as e:
                        st.error(f"❌ PowerPoint generation failed: {e}")
    except Exception as e:
        st.error(f"❌ Error generating dashboard: {e}")
# ------------------ قسم Info ------------------
st.markdown("<hr class='divider' id='info-section'>", unsafe_allow_html=True)
with st.expander("📖 How to Use - Click to view instructions"):
    st.markdown("""
    <div class='guide-title'>🎯 Welcome to a free tool provided by the company admin.!</div>
    هذه الأداة تقسم ودمج ملفات الإكسل <strong>بدقة وبدون فقدان التنسيق</strong>.
    ---
    ### 🔧 أولًا: التقسيم
    1. ارفع ملف Excel في خانة "Upload Excel File (Splitter/Merge)".
    2. اختر الشيت.
    3. اختر العمود اللي عاوز تقسّم عليه (مثل: "Area Manager").
    4. اضغط على **"Start Split"**.
    5. هيطلعلك ملف ZIP يحتوي على ملف منفصل لكل قيمة.
    ---
    ### 🔗 ثانيًا: الدمج
    - ارفع أكثر من ملف Excel في خانة "Upload Excel Files to Merge".
    - اضغط "Merge Files with Format".
    - طالع لك ملف واحد مدمج بالحفاظ على التنسيق.
    ---
    ### 📷 ثالثًا: تحويل الصور إلى PDF (جديد!)
    - ارفع صور JPG أو JPEG أو PNG في قسم "Convert Images to PDF".
    - اختر بين:
      - **"Create PDF (Original Quality)"** → PDF عادي.
      - **"Create PDF (CamScanner Style)"** → PDF مُحسّن (إذا كان `opencv` مثبتًا).
    - نزّل ملف PDF يحتوي على كل الصور كصفحات.
    - استخدم زر **"Clear All Images"** لمسح كل الصور دفعة واحدة.
    ---
    ### 📊 رابعًا: الـ Dashboard
    - ارفع ملف Excel في خانة "Upload Excel File for Dashboard (Auto)".
    - اختر الشيت.
    - استخدم Sidebar لاختيار "Primary Filter Column" (دروب ليست) ثم قيمه.
    - اختياريًا اختار أعمدة فلترة إضافية.
    - الداشبورد يبني رسومات أوتوماتيك ويعرضها في شبكة 3×2 داخل كروت.
    - اضغط **"Generate Dashboard PDF (charts only)"** لتنزيل PDF يحتوي على الرسومات فقط.
    - إذا حبيت تضيف جدول بالبيانات داخل الـ PDF فعّل الخيار "Include table in PDF report (optional)".
    - **جديد**: اضغط **"Export Dashboard to PowerPoint (PPTX)"** لتنزيل عرض تقديمي احترافي.
    ---
    🙋‍♂️ لأي استفسار: <a href="https://wa.me/201554694554" target="_blank">01554694554 (واتساب)</a>
    """, unsafe_allow_html=True)
