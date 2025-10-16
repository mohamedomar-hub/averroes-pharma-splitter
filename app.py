# -*- coding: utf-8 -*-
"""
Averroes Pharma File Splitter & Dashboard
Features added:
- Progress Bar for splitting
- Advanced Forecast: Moving Average + Trend Line (Linear Regression) + interactive Plotly chart
- Generate full PDF including KPIs, Charts and sample table
- Notification Toasts (uses streamlit-toast if available, else fallback to st.success)
"""
import streamlit as st
import pandas as pd
import numpy as np
import re
import os
from io import BytesIO
from zipfile import ZipFile
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl import load_workbook, Workbook

# Visualization / reports
import matplotlib.pyplot as plt
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image as RLImage, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
import plotly.express as px
import plotly.graph_objects as go

# Forecast / ML
from sklearn.linear_model import LinearRegression

# Image processing
from PIL import Image

# Try import toast notification lib, fallback to st.success
try:
    # pip name might be streamlit-toast or streamlit_toast depending on package; try both import styles
    try:
        from streamlit_toast import toast
    except Exception:
        from streamlit_toast import toast  # try again (keeps consistent)
    TOAST_AVAILABLE = True
except Exception:
    try:
        # some versions use streamlit_toast, alias check
        from streamlit_toast import toast
        TOAST_AVAILABLE = True
    except Exception:
        TOAST_AVAILABLE = False

# ---------------- Session state init ----------------
if 'clear_counter' not in st.session_state:
    st.session_state.clear_counter = 0

# ---------------- Page config ----------------
st.set_page_config(
    page_title="Averroes Pharma File Splitter & Dashboard",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ---------------- Styling ----------------
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
    label, .stSelectbox label, .stFileUploader label {
        color: #FFD700 !important;
        font-size: 16px !important;
        font-weight: bold !important;
    }
    .stButton>button, .stDownloadButton>button {
        background-color: #FFD700 !important;
        color: black !important;
        font-weight: bold !important;
        font-size: 16px !important;
        border-radius: 12px !important;
        padding: 10px 20px !important;
        border: none !important;
        box-shadow: 0 4px 8px rgba(0,0,0,0.3) !important;
        transition: all 0.3s ease !important;
        margin-top: 8px !important;
    }
    .kpi-card {
        padding: 12px;
        border-radius: 10px;
        color: white;
        text-align: center;
        box-shadow: 0 6px 18px rgba(0,0,0,0.3);
        font-weight: 700;
        margin: 8px;
    }
    .kpi-title { font-size: 13px; opacity: 0.9; }
    .kpi-value { font-size: 20px; margin-top:4px; }
    .chart-card {
        background-color: #00264d;
        border: 1px solid #FFD700;
        border-radius: 12px;
        padding: 12px;
        margin: 10px 0;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3);
    }
    hr.divider { border: 1px solid #FFD700; opacity: 0.6; margin:30px 0; }
    hr.divider-dashed { border: 1px dashed #FFD700; opacity: 0.7; margin:25px 0; }
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ---------------- Helpers ----------------
def show_toast(message, state="success", duration=3):
    """Show toast if available, else fallback to st.success/info/warning"""
    if TOAST_AVAILABLE:
        try:
            toast(message, state=state, duration=duration)
            return
        except Exception:
            pass
    # fallback
    if state == "success":
        st.success(message)
    elif state == "warning":
        st.warning(message)
    elif state == "error":
        st.error(message)
    else:
        st.info(message)

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

# existing build_pdf helper (enhanced to include KPIs & sample table)
def build_pdf(sheet_title, charts_buffers, include_table=False, filtered_df=None, kpis=None, max_table_rows=200):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=20, rightMargin=20, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    elements = []
    elements.append(Paragraph(f"<para align='center'><b>{sheet_title} — Full Dashboard Report</b></para>", styles['Title']))
    elements.append(Spacer(1,8))
    elements.append(Paragraph("<para align='center'>Averroes Pharma - Auto Generated Dashboard</para>", styles['Heading3']))
    elements.append(Spacer(1,12))

    # KPIs
    if kpis:
        elements.append(Paragraph("<b>Key Performance Indicators</b>", styles['Heading2']))
        data = [["Metric", "Value"]]
        for k, v in kpis.items():
            data.append([str(k), str(v)])
        tbl = Table(data, hAlign='CENTER')
        tbl.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#FFD700")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.black),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('GRID', (0,0), (-1,-1), 0.3, colors.grey),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ]))
        elements.append(tbl)
        elements.append(Spacer(1,12))

    # Charts
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

    # Table
    if include_table and filtered_df is not None:
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
            ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
        ]))
        elements.append(tbl)

    doc.build(elements)
    buf.seek(0)
    return buf

def build_pptx(sheet_title, charts_buffers):
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN

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

# ---------------- Navigation & Header ----------------
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

st.markdown("<h1 style='text-align:center; color:#FFD700;'>💊 Averroes Pharma File Splitter & Dashboard</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align:center; color:white;'>✂ Split, Merge, Image-to-PDF & Auto Dashboard Generator</h3>", unsafe_allow_html=True)

# ---------------- Tabs ----------------
tab1, tab2, tab3, tab4 = st.tabs([
    "📂 Split & Merge",
    "📷 Image to PDF",
    "📊 Auto Dashboard",
    "ℹ️ Info"
])

# ---------------- Tab 1: Split & Merge (with progress) ----------------
with tab1:
    st.markdown("### ✂ Split Excel/CSV File")
    uploaded_file = st.file_uploader(
        "📂 Upload Excel or CSV File (Splitter/Merge)",
        type=["xlsx", "csv"],
        accept_multiple_files=False,
        key=f"split_uploader_{st.session_state.clear_counter}"
    )
    if uploaded_file:
        # display file
        st.markdown(f"**Uploaded:** {uploaded_file.name} ({uploaded_file.size//1024} KB)")
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
            st.dataframe(df.head(200), use_container_width=True)

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

            # Add progress bar option
            include_progress = st.checkbox("Show Progress Bar during split", value=True)
            add_timestamp_to_filename = st.checkbox("Append date to generated filenames", value=True)

            if st.button("🚀 Start Split"):
                # start splitting with progress
                try:
                    if file_ext == "csv":
                        unique_values = df[col_to_split].dropna().unique()
                        zip_buffer = BytesIO()
                        with ZipFile(zip_buffer, "w") as zip_file:
                            if include_progress:
                                progress = st.progress(0)
                                status_text = st.empty()
                            for i, value in enumerate(unique_values):
                                filtered_df = df[df[col_to_split] == value]
                                csv_buffer = BytesIO()
                                filtered_df.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                                csv_buffer.seek(0)
                                fname = f"{_safe_name(value)}"
                                if add_timestamp_to_filename:
                                    fname = f"{fname}_{pd.Timestamp.now().strftime('%Y-%m-%d')}"
                                file_name = f"{fname}.csv"
                                zip_file.writestr(file_name, csv_buffer.read())
                                if include_progress:
                                    progress.progress((i+1)/len(unique_values))
                                    status_text.text(f"Created file: {file_name}")
                        zip_buffer.seek(0)
                        show_toast("🎉 Splitting completed successfully!", state="success")
                        st.download_button(
                            label="📥 Download Split Files (ZIP)",
                            data=zip_buffer.getvalue(),
                            file_name=f"Split_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip",
                            mime="application/zip"
                        )
                    else:
                        if split_option == "Split by Column Values":
                            ws = original_wb[selected_sheet]
                            col_idx = df.columns.get_loc(col_to_split) + 1
                            unique_values = df[col_to_split].dropna().unique()
                            zip_buffer = BytesIO()
                            with ZipFile(zip_buffer, "w") as zip_file:
                                if include_progress:
                                    progress = st.progress(0)
                                    status_text = st.empty()
                                for i, value in enumerate(unique_values):
                                    new_wb = Workbook()
                                    default_ws = new_wb.active
                                    new_wb.remove(default_ws)
                                    new_ws = new_wb.create_sheet(title=str(value)[:30] if value else "Sheet")
                                    # copy header and style from ws first row
                                    for cell in ws[1]:
                                        dst_cell = new_ws.cell(1, cell.column, cell.value)
                                        try:
                                            if cell.has_style:
                                                dst_cell.font = cell.font
                                                dst_cell.fill = cell.fill
                                                dst_cell.border = cell.border
                                                dst_cell.alignment = cell.alignment
                                                dst_cell.number_format = cell.number_format
                                        except Exception:
                                            pass
                                    row_idx = 2
                                    for row in ws.iter_rows(min_row=2):
                                        cell_in_col = row[col_idx - 1]
                                        if cell_in_col.value == value:
                                            for src_cell in row:
                                                dst_cell = new_ws.cell(row_idx, src_cell.column, src_cell.value)
                                                try:
                                                    if src_cell.has_style:
                                                        dst_cell.font = src_cell.font
                                                        dst_cell.fill = src_cell.fill
                                                        dst_cell.border = src_cell.border
                                                        dst_cell.alignment = src_cell.alignment
                                                        dst_cell.number_format = src_cell.number_format
                                                except Exception:
                                                    pass
                                            row_idx += 1
                                    # preserve column widths
                                    try:
                                        for col_letter in ws.column_dimensions:
                                            if ws.column_dimensions[col_letter].width:
                                                new_ws.column_dimensions[col_letter].width = ws.column_dimensions[col_letter].width
                                    except Exception:
                                        pass
                                    file_buffer = BytesIO()
                                    new_wb.save(file_buffer)
                                    file_buffer.seek(0)
                                    fname = f"{_safe_name(value)}"
                                    if add_timestamp_to_filename:
                                        fname = f"{fname}_{pd.Timestamp.now().strftime('%Y-%m-%d')}"
                                    file_name = f"{fname}.xlsx"
                                    zip_file.writestr(file_name, file_buffer.read())
                                    if include_progress:
                                        progress.progress((i+1)/len(unique_values))
                                        status_text.text(f"Created file: {file_name}")
                            zip_buffer.seek(0)
                            show_toast("🎉 Splitting completed successfully!", state="success")
                            st.download_button(
                                label="📥 Download Split Files (ZIP)",
                                data=zip_buffer.getvalue(),
                                file_name=f"Split_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip",
                                mime="application/zip"
                            )

                        elif split_option == "Split Each Sheet into Separate File":
                            zip_buffer = BytesIO()
                            sheet_names_local = original_wb.sheetnames
                            with ZipFile(zip_buffer, "w") as zip_file:
                                if include_progress:
                                    progress = st.progress(0)
                                    status_text = st.empty()
                                for i, sheet_name in enumerate(sheet_names_local):
                                    new_wb = Workbook()
                                    default_ws = new_wb.active
                                    new_wb.remove(default_ws)
                                    new_ws = new_wb.create_sheet(title=sheet_name)
                                    src_ws = original_wb[sheet_name]
                                    for row in src_ws.iter_rows():
                                        for src_cell in row:
                                            dst_cell = new_ws.cell(src_cell.row, src_cell.column, src_cell.value)
                                            try:
                                                if src_cell.has_style:
                                                    dst_cell.font = src_cell.font
                                                    dst_cell.fill = src_cell.fill
                                                    dst_cell.border = src_cell.border
                                                    dst_cell.alignment = src_cell.alignment
                                                    dst_cell.number_format = src_cell.number_format
                                            except Exception:
                                                pass
                                    # merged cells
                                    try:
                                        if src_ws.merged_cells.ranges:
                                            for merged_range in src_ws.merged_cells.ranges:
                                                new_ws.merge_cells(str(merged_range))
                                                top_left_cell = src_ws.cell(merged_range.min_row, merged_range.min_col)
                                                merged_value = top_left_cell.value
                                                new_ws.cell(merged_range.min_row, merged_range.min_col, merged_value)
                                    except Exception:
                                        pass
                                    try:
                                        for col_letter in src_ws.column_dimensions:
                                            if src_ws.column_dimensions[col_letter].width:
                                                new_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width
                                    except Exception:
                                        pass
                                    file_buffer = BytesIO()
                                    new_wb.save(file_buffer)
                                    file_buffer.seek(0)
                                    fname = f"{_safe_name(sheet_name)}"
                                    if add_timestamp_to_filename:
                                        fname = f"{fname}_{pd.Timestamp.now().strftime('%Y-%m-%d')}"
                                    file_name = f"{fname}.xlsx"
                                    zip_file.writestr(file_name, file_buffer.read())
                                    if include_progress:
                                        progress.progress((i+1)/len(sheet_names_local))
                                        status_text.text(f"Created file: {file_name}")
                            zip_buffer.seek(0)
                            show_toast("🎉 Splitting completed successfully!", state="success")
                            st.download_button(
                                label="📥 Download Split Files (ZIP)",
                                data=zip_buffer.getvalue(),
                                file_name=f"SplitBySheets_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip",
                                mime="application/zip"
                            )
                except Exception as e:
                    st.error(f"❌ Error during split: {e}")
                    show_toast("❌ Split failed", state="error")
        except Exception as e:
            st.error(f"❌ Error processing file: {e}")
    else:
        st.markdown("<p style='text-align:center; color:#FFD700;'>⚠️ No file uploaded yet for splitting.</p>", unsafe_allow_html=True)

    st.markdown("<hr class='divider-dashed'>", unsafe_allow_html=True)
    st.markdown("### 🔄 Merge Excel/CSV Files")
    merge_files = st.file_uploader(
        "📤 Upload Excel or CSV Files to Merge",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        key=f"merge_uploader_{st.session_state.clear_counter}"
    )
    if merge_files:
        st.markdown("#### Uploaded files:")
        for i, f in enumerate(merge_files):
            st.markdown(f"{i+1}. {f.name} ({f.size//1024} KB)")
        if st.button("🗑️ Clear All Merged Files", key="clear_merge"):
            st.session_state.clear_counter += 1
            st.rerun()
        if st.button("✨ Merge Files"):
            with st.spinner("Merging files..."):
                try:
                    all_dfs = []
                    for file in merge_files:
                        ext = file.name.split('.')[-1].lower()
                        if ext == "csv":
                            dfm = pd.read_csv(file)
                        else:
                            # if excel has multiple sheets, take first
                            try:
                                dfm = pd.read_excel(file)
                            except:
                                dfm = pd.read_excel(file, sheet_name=0)
                        dfm["Source_File"] = file.name
                        all_dfs.append(dfm)
                    merged_df = pd.concat(all_dfs, ignore_index=True)
                    output_buffer = BytesIO()
                    merged_df.to_excel(output_buffer, index=False, engine='openpyxl')
                    output_buffer.seek(0)
                    show_toast("✅ Merged successfully!", state="success")
                    st.download_button(
                        label="📥 Download Merged File (Excel)",
                        data=output_buffer.getvalue(),
                        file_name="Merged_Consolidated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"❌ Error during merge: {e}")

# ---------------- Tab 2: Image to PDF ----------------
with tab2:
    st.markdown("### 📷 Convert Images to PDF")
    uploaded_images = st.file_uploader(
        "📤 Upload JPG/JPEG/PNG Images to Convert to PDF",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key=f"image_uploader_{st.session_state.clear_counter}"
    )
    if uploaded_images:
        st.markdown("#### Uploaded images:")
        for i, imf in enumerate(uploaded_images):
            st.markdown(f"{i+1}. {imf.name} ({imf.size//1024} KB)")
        if st.button("🗑️ Clear All Images", key="clear_images"):
            st.session_state.clear_counter += 1
            st.rerun()
        try:
            import cv2
            import numpy as np
            def enhance_image_for_pdf(image_pil):
                image = np.array(image_pil)
                if image.ndim == 2:
                    image = cv2.cvtColor(image, cv2.COLOR_GRAY2BGR)
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
                        show_toast("✅ Enhanced PDF created successfully!", state="success")
                        st.download_button(
                            label="📥 Download Enhanced PDF",
                            data=pdf_buffer.getvalue(),
                            file_name="Enhanced_Images_CamScanner.pdf",
                            mime="application/pdf"
                        )
                    except Exception as e:
                        st.error(f"❌ Error creating enhanced PDF: {e}")
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
                        show_toast("✅ PDF created successfully!", state="success")
                        st.download_button(
                            label="📥 Download Original PDF",
                            data=pdf_buffer.getvalue(),
                            file_name="Images_Combined.pdf",
                            mime="application/pdf"
                        )
                    except Exception as e:
                        st.error(f"❌ Error creating PDF: {e}")
        except ImportError:
            st.warning("⚠️ CamScanner effect requires 'opencv-python'. Install it to enable this feature.")
    else:
        st.info("📤 Please upload one or more JPG/JPEG/PNG images to convert them into a single PDF file.")

# ---------------- Tab 3: Auto Dashboard (with Advanced Forecast) ----------------
with tab3:
    st.markdown("### 📊 Interactive Auto Dashboard Generator (with Moving Average + Trend)")
    dashboard_file = st.file_uploader(
        "📊 Upload Excel or CSV File for Dashboard (Auto)",
        type=["xlsx", "csv"],
        key=f"dashboard_uploader_{st.session_state.clear_counter}"
    )
    if dashboard_file:
        show_toast("File uploaded for Dashboard", state="success", duration=2)
        try:
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

            st.markdown("### 🔍 Data Preview (original)")
            st.dataframe(df0.head(), use_container_width=True)

            # Detect numeric columns and potential date/time
            numeric_cols = df0.select_dtypes(include='number').columns.tolist()
            possible_date_cols = [c for c in df0.columns if any(k in c.lower() for k in ["date", "month", "year"]) or pd.api.types.is_datetime64_any_dtype(df0[c])]
            # allow user to choose time column for forecasting
            st.markdown("### 🔎 Select columns for analysis")
            measure_col = None
            if numeric_cols:
                measure_col = st.selectbox("🎯 Select Sales/Value Column (for KPIs & Charts)", numeric_cols, index=0)
            else:
                st.warning("No numeric columns available for KPIs or forecasting.")

            date_col = None
            if possible_date_cols:
                date_col = st.selectbox("🗓️ Select Date/Time Column (for Forecasting & Trend)", ["-- None --"] + possible_date_cols)
                if date_col == "-- None --":
                    date_col = None

            # Offer moving average window selection and future periods
            st.markdown("### ⚙ Forecast Settings")
            ma_window = st.number_input("Moving Average window (periods)", min_value=2, max_value=52, value=3, step=1)
            future_periods = st.number_input("Forecast future periods (int)", min_value=1, max_value=12, value=3, step=1)
            apply_groupby = st.selectbox("Optional: Group by column (for group-specific KPIs/charts)", ["-- None --"] + [c for c in df0.columns if df0[c].dtype == object], index=0)

            # Create working df copy and handle month-name numeric mapping if necessary
            df_work = df0.copy()
            # Try parse date_col if exists
            if date_col:
                try:
                    df_work[date_col] = pd.to_datetime(df_work[date_col])
                except Exception:
                    # it might be strings like 'Jan', 'Feb' etc. try mapping
                    try:
                        df_work[date_col] = pd.to_datetime(df_work[date_col], errors='coerce')
                    except:
                        pass

            # Filter UI
            cat_cols = [c for c in df_work.columns if df_work[c].dtype == object or df_work[c].dtype.name.startswith('category')]
            st.sidebar.header("🔍 Dynamic Filters")
            primary_filter_col = None
            if cat_cols:
                primary_filter_col = st.sidebar.selectbox("Primary Filter Column", ["-- None --"] + cat_cols, index=0)
                if primary_filter_col == "-- None --":
                    primary_filter_col = None
            primary_values = None
            if primary_filter_col:
                vals = df_work[primary_filter_col].dropna().astype(str).unique().tolist()
                try:
                    vals = sorted(vals)
                except:
                    pass
                primary_values = st.sidebar.multiselect(f"Filter values for {primary_filter_col}", vals, default=vals)

            # Apply filters
            filtered = df_work.copy()
            if primary_filter_col and primary_values is not None and len(primary_values) > 0:
                filtered = filtered[filtered[primary_filter_col].astype(str).isin(primary_values)]

            st.markdown("### 📈 Filtered Data Preview")
            st.dataframe(filtered.head(200), use_container_width=True)

            # === Compute KPIs ===
            kpi_values = {}
            if measure_col and measure_col in filtered.columns:
                kpi_values['Total'] = filtered[measure_col].sum()
                kpi_values['Average'] = filtered[measure_col].mean()
                kpi_values['Count'] = filtered.shape[0]
            # unique dims
            for dim_alias, aliases in {"Area": ["area", "region"], "Branch": ["branch", "location"], "Rep": ["rep", "representative"]}.items():
                found = _find_col(filtered, aliases)
                if found:
                    kpi_values[f"Unique {dim_alias}"] = filtered[found].nunique()

            # Period comparison if two period columns exist (left as existing logic)
            # Display KPI cards
            st.markdown("### 🚀 KPIs")
            kpi_cards = []
            for k, v in kpi_values.items():
                kpi_cards.append({'title': k, 'value': f"{v:,.2f}" if isinstance(v, float) else f"{v}", 'color': 'linear-gradient(135deg, #28a745, #85e085)', 'icon': '📈'})

            cols_kpi = st.columns(min(6, max(1, len(kpi_cards))))
            for i, card in enumerate(kpi_cards[:6]):
                with cols_kpi[i]:
                    st.markdown(f"""
                    <div class='kpi-card' style='background:{card['color']};'>
                        <div class='kpi-title'>{card['icon']} {card['title']}</div>
                        <div class='kpi-value'>{card['value']}</div>
                    </div>
                    """, unsafe_allow_html=True)

            # === Advanced Forecast: Moving Average + Trend Line + future forecast ===
            charts_buffers = []  # keep (BytesIO_png, caption) for export
            plotly_figs = []

            if measure_col and measure_col in filtered.columns:
                # If grouping selected, compute grouped metric (sum) over date
                if date_col and date_col in filtered.columns:
                    df_for_forecast = filtered[[date_col, measure_col]].dropna().copy()
                    df_for_forecast = df_for_forecast.sort_values(date_col)
                    # aggregated by date
                    agg = df_for_forecast.groupby(date_col)[measure_col].sum().reset_index()
                    agg = agg.sort_values(date_col)
                    # compute moving average
                    agg['MA'] = agg[measure_col].rolling(window=ma_window, min_periods=1).mean()
                    # trend line via LinearRegression on ordinal X
                    X = np.arange(len(agg)).reshape(-1,1)
                    y = agg[measure_col].values
                    lr = LinearRegression().fit(X, y)
                    trend = lr.predict(X)
                    agg['Trend'] = trend
                    # future forecast using trend projection and optional extension
                    future_X = np.arange(len(agg), len(agg)+int(future_periods)).reshape(-1,1)
                    future_trend = lr.predict(future_X)
                    # Prepare Plotly figure
                    fig = go.Figure()
                    fig.add_trace(go.Scatter(x=agg[date_col], y=agg[measure_col], mode='lines+markers', name='Actual'))
                    fig.add_trace(go.Scatter(x=agg[date_col], y=agg['MA'], mode='lines', name=f'MA ({ma_window})'))
                    fig.add_trace(go.Scatter(x=agg[date_col], y=agg['Trend'], mode='lines', name='Trend Line', line=dict(dash='dash')))
                    # future x axis labels - use Period index if dates are monthly/daily
                    try:
                        last_date = agg[date_col].iloc[-1]
                        # try infer frequency (if monthly -> add months)
                        future_dates = []
                        freq = pd.infer_freq(agg[date_col])
                        if freq is not None:
                            future_dates = pd.date_range(start=last_date, periods=int(future_periods)+1, freq=freq, closed='right').tolist()
                        else:
                            # fallback to numeric sequential labels
                            future_dates = list(range(len(agg), len(agg)+int(future_periods)))
                        fig.add_trace(go.Scatter(x=future_dates, y=future_trend, mode='lines+markers', name='Forecast (trend projection)', line=dict(color='firebrick', dash='dot')))
                    except Exception:
                        # fallback numeric
                        fig.add_trace(go.Scatter(x=list(range(len(agg))) + list(range(len(agg), len(agg)+int(future_periods))), y=list(agg['Trend']) + list(future_trend), mode='lines+markers', name='Forecast'))

                    fig.update_layout(title=f"{measure_col} — Actual / MA({ma_window}) / Trend / Forecast", template="plotly_white", autosize=True)
                    st.markdown("### 🔮 Forecast & Trend")
                    st.plotly_chart(fig, use_container_width=True, theme="streamlit")
                    show_toast("📈 Forecast generated (MA + Trend)", state="success", duration=2)

                    # prepare a PNG buffer for PDF/PPT export
                    try:
                        img_bytes = fig.to_image(format="png", width=1400, height=700, scale=2)
                        buf = BytesIO(img_bytes)
                        charts_buffers.append((buf, "Forecast: Actual, Moving Average & Trend"))
                        plotly_figs.append((fig, "Forecast: Actual, Moving Average & Trend"))
                    except Exception:
                        # fallback: render matplotlib
                        plt.figure(figsize=(10,4))
                        plt.plot(agg[date_col], agg[measure_col], label='Actual')
                        plt.plot(agg[date_col], agg['MA'], label=f'MA ({ma_window})')
                        plt.plot(agg[date_col], agg['Trend'], label='Trend')
                        plt.legend()
                        plt.tight_layout()
                        buf = BytesIO()
                        plt.savefig(buf, format='png')
                        buf.seek(0)
                        charts_buffers.append((buf, "Forecast (matplotlib fallback)"))

                else:
                    st.info("For forecasting please select a valid Date/Time column that contains chronological data.")
            else:
                st.info("Please select a numeric measure column to compute forecasts/charts.")

            # === Additional auto charts (Top N) ===
            rep_col = _find_col(filtered, ["rep", "representative", "salesman", "employee", "name", "mr"])
            if rep_col and measure_col in filtered.columns:
                try:
                    rep_data = filtered.groupby(rep_col)[measure_col].sum().sort_values(ascending=False)
                    topN = st.selectbox("Top N for Employees chart", [5,10,15], index=1)
                    top_series = rep_data.head(topN).reset_index().rename(columns={measure_col: "value"})
                    fig_top = px.bar(top_series, x=rep_col, y="value", title=f"Top {topN} Employees", text="value")
                    fig_top.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
                    fig_top.update_layout(margin=dict(t=40,b=20,l=10,r=10), template="plotly_white")
                    st.plotly_chart(fig_top, use_container_width=True, theme="streamlit")
                    try:
                        img_bytes = fig_top.to_image(format="png", width=1200, height=600)
                        buf_top = BytesIO(img_bytes)
                        charts_buffers.append((buf_top, f"Top {topN} Employees"))
                        plotly_figs.append((fig_top, f"Top {topN} Employees"))
                    except Exception:
                        pass
                except Exception:
                    pass

            # === Export options: Excel, PDF, PPTX ===
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

            if st.button("📥 Generate Dashboard PDF (Full - KPIs + Charts + Table)"):
                with st.spinner("Generating Dashboard PDF..."):
                    try:
                        pdf_buffer = build_pdf(sheet_title, charts_buffers, include_table=True, filtered_df=filtered, kpis=kpi_values, max_table_rows=200)
                        st.success("✅ Dashboard PDF ready.")
                        st.download_button(
                            label="⬇️ Download Dashboard PDF",
                            data=pdf_buffer.getvalue(),
                            file_name=f"{_safe_name(sheet_title)}_Dashboard.pdf",
                            mime="application/pdf"
                        )
                        show_toast("📄 PDF Generated Successfully", state="success", duration=3)
                    except Exception as e:
                        st.error(f"❌ PDF generation failed: {e}")
                        show_toast("❌ PDF generation failed", state="error")

            if st.button("📥 Export Dashboard to PPTX (charts only)"):
                try:
                    ppt_buffer = build_pptx(sheet_title, charts_buffers)
                    st.download_button(
                        label="⬇️ Download PPTX",
                        data=ppt_buffer.getvalue(),
                        file_name=f"{_safe_name(sheet_title)}_Dashboard.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                    show_toast("📁 PPTX exported", state="success")
                except Exception as e:
                    st.error(f"❌ PPTX export failed: {e}")
                    show_toast("❌ PPTX export failed", state="error")

        except Exception as e:
            st.error(f"❌ Error generating dashboard: {e}")
            show_toast("❌ Dashboard generation error", state="error")
    else:
        st.info("Upload a file to generate dashboard (Excel/CSV).")

# ---------------- Tab 4: Info ----------------
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
                <li>Use Forecast Settings to compute Moving Average and Trend.</li>
                <li>Export PDF/PPTX of the dashboard.</li>
            </ul>
        </li>
    </ol>
    <br>
    <h3 style='color:#FFD700;'>💡 Tips</h3>
    <ul>
        <li>Forecasting works best with chronological Date/Time column.</li>
        <li>For moving averages choose a window appropriate to your data frequency (e.g., 3-6 for monthly, 7-14 for daily).</li>
    </ul>
    """, unsafe_allow_html=True)
