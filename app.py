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
import datetime

# ---------------- Session ----------------
if 'clear_counter' not in st.session_state:
    st.session_state.clear_counter = 0
if 'active_section' not in st.session_state:
    st.session_state.active_section = "Split Files"

# ---------------- Page Setup ----------------
st.set_page_config(
    page_title="Averroes Pharma File Splitter & Dashboard",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------- Styling: fixed sidebar + theme ----------------
# note: Streamlit's internal classnames can change; included multiple selectors for compatibility
fixed_sidebar_css = """
<style>
/* Make Streamlit sidebar fixed and styled */
[data-testid="stSidebar"] > div:first-child {
  position: fixed;
  left: 0;
  top: 0;
  height: 100vh;
  width: 260px;
  padding-top: 28px;
  background: linear-gradient(180deg,#001a33,#00264d);
  border-right: 1px solid rgba(255,215,0,0.08);
  z-index: 9999;
  color: #FFD700;
}

/* Adjust main content to avoid overlap with fixed sidebar */
section[data-testid="stAppViewContainer"] > div > section {
  margin-left: 280px !important;
  padding-top: 18px !important;
}

/* Sidebar inner texts/buttons style */
[data-testid="stSidebar"] .css-1d391kg, 
[data-testid="stSidebar"] .css-1lcbmhc {
  color: #FFD700;
}

/* Buttons and download buttons global style (keeps streamlit behavior) */
.stButton>button, .stDownloadButton>button {
    background-color: #FFD700 !important;
    color: black !important;
    font-weight: bold !important;
    border-radius: 10px !important;
}

/* KPI card styling */
.kpi-card { padding: 12px; border-radius: 10px; color: white; text-align:center; margin:6px; box-shadow:0 6px 18px rgba(0,0,0,0.3); }
.kpi-title { font-size:13px; opacity:0.9; }
.kpi-value { font-size:18px; margin-top:4px; }

/* Chart card */
.chart-card { background-color:#00264d; border:1px solid #FFD700; border-radius:12px; padding:12px; margin:10px 0; }

/* Small helper to ensure sidebar items spacing */
[data-testid="stSidebar"] .stSelectbox, [data-testid="stSidebar"] .stRadio {
    margin-top: 10px;
    margin-bottom: 10px;
}
</style>
"""
st.markdown(fixed_sidebar_css, unsafe_allow_html=True)

# ---------------- Helper functions ----------------
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

def display_uploaded_files(file_list, file_type="Excel/CSV"):
    if file_list:
        st.markdown("### 📁 Uploaded Files:")
        for i, f in enumerate(file_list):
            try:
                st.markdown(f"<div style='background:#003366; color:white; padding:6px 8px; border-radius:6px; margin:4px 0; display:block;'>"
                            f"{i+1}. {f.name} ({f.size//1024} KB)</div>", unsafe_allow_html=True)
            except Exception:
                st.write(f"{i+1}. {f.name}")

# Lightweight PDF builder using reportlab
def build_pdf(sheet_title, charts_buffers, include_table=False, filtered_df=None, max_table_rows=200):
    buf = BytesIO()
    try:
        doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=20, rightMargin=20, topMargin=20, bottomMargin=20)
        styles = getSampleStyleSheet()
        elements = []
        elements.append(Paragraph(f"<para align='center'><b>{sheet_title} Report</b></para>", styles['Title']))
        elements.append(Spacer(1,12))
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
            ]))
            elements.append(tbl)
        doc.build(elements)
        buf.seek(0)
        return buf
    except Exception as e:
        st.warning(f"PDF build failed: {e}")
        return None

def build_pptx(sheet_title, charts_buffers):
    try:
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
            except Exception:
                pass
        pptx_buffer = BytesIO()
        prs.save(pptx_buffer)
        pptx_buffer.seek(0)
        return pptx_buffer
    except Exception as e:
        st.warning(f"PPTX build failed: {e}")
        return None

# ---------------- Sidebar Navigation ----------------
NAV_OPTIONS = ["Split Files", "Merge Files", "Convert Images to PDF", "Dashboard"]
if 'active_section' not in st.session_state or st.session_state.active_section not in NAV_OPTIONS:
    st.session_state.active_section = NAV_OPTIONS[0]

section = st.sidebar.selectbox("🧭 Choose Section", NAV_OPTIONS, index=NAV_OPTIONS.index(st.session_state.active_section))
st.session_state.active_section = section

# top placeholder for progress bars (we show progress bars using this placeholder)
progress_placeholder = st.empty()

# ---------------- Split Files ----------------
if st.session_state.active_section == "Split Files":
    st.markdown("### ✂ Split Excel/CSV File")
    uploaded_file = st.file_uploader("📂 Upload Excel or CSV File (Splitter)", type=["xlsx", "csv"], key=f"split_uploader_{st.session_state.clear_counter}")
    if uploaded_file:
        display_uploaded_files([uploaded_file])
        if st.button("🗑️ Clear Uploaded File", key="clear_split"):
            st.session_state.clear_counter += 1
            st.rerun()
        try:
            file_ext = uploaded_file.name.split('.')[-1].lower()
            if file_ext == "csv":
                df = pd.read_csv(uploaded_file)
                sheet_names = ["Sheet1"]
                selected_sheet = "Sheet1"
                st.success("✅ CSV uploaded")
            else:
                input_bytes = uploaded_file.getvalue()
                original_wb = load_workbook(filename=BytesIO(input_bytes), data_only=False)
                sheet_names = original_wb.sheetnames
                st.success(f"✅ Excel uploaded: {len(sheet_names)} sheets")
                selected_sheet = st.selectbox("Select Sheet (for Split)", sheet_names)
                df = pd.read_excel(BytesIO(input_bytes), sheet_name=selected_sheet)
            st.dataframe(df.head(200), use_container_width=True)
            col_to_split = st.selectbox("Split by Column", df.columns)
            split_option = st.radio("Choose split method:", ["Split by Column Values", "Split Each Sheet into Separate File"], index=0)
            add_timestamp = st.checkbox("Append date to filenames", value=True)
            show_progress = st.checkbox("Show progress bar", value=True)
            if st.button("🚀 Start Split"):
                try:
                    if show_progress:
                        prog = progress_placeholder.progress(0)
                        status = st.empty()
                    if file_ext == "csv":
                        unique_vals = df[col_to_split].dropna().unique().tolist()
                        zip_buffer = BytesIO()
                        with ZipFile(zip_buffer, "w") as zf:
                            for i, val in enumerate(unique_vals):
                                subset = df[df[col_to_split] == val]
                                buf = BytesIO()
                                subset.to_excel(buf, index=False, engine='openpyxl')
                                buf.seek(0)
                                safe = _safe_name(val)
                                if add_timestamp:
                                    safe = f"{safe}_{datetime.date.today().isoformat()}"
                                fname = f"{safe}.xlsx"
                                zf.writestr(fname, buf.read())
                                if show_progress:
                                    prog.progress((i+1)/len(unique_vals))
                                    status.text(f"Created {fname}")
                        zip_buffer.seek(0)
                        st.success("🎉 Split completed")
                        st.download_button("📥 Download Split ZIP", data=zip_buffer.getvalue(), file_name=f"Split_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip", mime="application/zip")
                    else:
                        if split_option == "Split by Column Values":
                            ws = original_wb[selected_sheet]
                            unique_vals = df[col_to_split].dropna().unique().tolist()
                            zip_buffer = BytesIO()
                            with ZipFile(zip_buffer, "w") as zf:
                                for i, val in enumerate(unique_vals):
                                    new_wb = Workbook()
                                    new_ws = new_wb.active
                                    # copy header
                                    for j, header in enumerate(df.columns, start=1):
                                        new_ws.cell(1, j, header)
                                    row_idx = 2
                                    for _, row in df[df[col_to_split] == val].iterrows():
                                        for j, header in enumerate(df.columns, start=1):
                                            new_ws.cell(row_idx, j, row[header])
                                        row_idx += 1
                                    buf = BytesIO()
                                    new_wb.save(buf)
                                    buf.seek(0)
                                    safe = _safe_name(val)
                                    if add_timestamp:
                                        safe = f"{safe}_{datetime.date.today().isoformat()}"
                                    fname = f"{safe}.xlsx"
                                    zf.writestr(fname, buf.read())
                                    if show_progress:
                                        prog.progress((i+1)/len(unique_vals))
                                        status.text(f"Created {fname}")
                            zip_buffer.seek(0)
                            st.success("🎉 Split completed")
                            st.download_button("📥 Download Split ZIP", data=zip_buffer.getvalue(), file_name=f"Split_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip", mime="application/zip")
                        else:
                            # split each sheet
                            zip_buffer = BytesIO()
                            with ZipFile(zip_buffer, "w") as zf:
                                for i, sname in enumerate(original_wb.sheetnames):
                                    sheet_df = pd.read_excel(BytesIO(input_bytes), sheet_name=sname)
                                    buf = BytesIO()
                                    sheet_df.to_excel(buf, index=False, engine='openpyxl')
                                    buf.seek(0)
                                    fname = f"{_safe_name(sname)}.xlsx"
                                    zf.writestr(fname, buf.read())
                                    if show_progress:
                                        prog.progress((i+1)/len(original_wb.sheetnames))
                                        status.text(f"Created {fname}")
                            zip_buffer.seek(0)
                            st.success("🎉 Split completed (sheets)")
                            st.download_button("📥 Download SplitBySheets ZIP", data=zip_buffer.getvalue(), file_name=f"SplitBySheets_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip", mime="application/zip")
                except Exception as e:
                    st.error(f"Error during split: {e}")
                finally:
                    progress_placeholder.empty()
        except Exception as e:
            st.error(f"❌ Error processing file: {e}")
            progress_placeholder.empty()
    else:
        st.info("Upload a file to split.")

# ---------------- Merge Files ----------------
elif st.session_state.active_section == "Merge Files":
    st.markdown("### 🔀 Merge Excel/CSV Files")
    merge_files = st.file_uploader("📤 Upload Excel or CSV Files to Merge", type=["xlsx","csv"], accept_multiple_files=True, key=f"merge_uploader_{st.session_state.clear_counter}")
    if merge_files:
        display_uploaded_files(merge_files)
        if st.button("🗑️ Clear All Merged Files", key="clear_merge"):
            st.session_state.clear_counter += 1
            st.rerun()
        add_source = st.checkbox("Add Source_File column", value=True)
        show_progress = st.checkbox("Show progress bar", value=True)
        if st.button("✨ Start Merge"):
            try:
                if show_progress:
                    prog = progress_placeholder.progress(0)
                    status = st.empty()
                dfs = []
                for i, f in enumerate(merge_files):
                    ext = f.name.split('.')[-1].lower()
                    try:
                        if ext == "csv":
                            d = pd.read_csv(f)
                        else:
                            d = pd.read_excel(f)
                        if add_source:
                            d["Source_File"] = f.name
                        dfs.append(d)
                    except Exception as e:
                        st.write(f"Could not read {f.name}: {e}")
                    if show_progress:
                        prog.progress((i+1)/len(merge_files))
                        status.text(f"Processed {i+1}/{len(merge_files)}")
                if len(dfs) == 0:
                    st.warning("No files readable to merge.")
                else:
                    merged = pd.concat(dfs, ignore_index=True)
                    out_buf = BytesIO()
                    merged.to_excel(out_buf, index=False, engine='openpyxl')
                    out_buf.seek(0)
                    st.success("✅ Merged successfully")
                    st.download_button("📥 Download Merged Excel", data=out_buf.getvalue(), file_name="Merged_Consolidated.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"❌ Error during merge: {e}")
            finally:
                progress_placeholder.empty()
    else:
        st.info("Upload files to merge.")

# ---------------- Images to PDF ----------------
elif st.session_state.active_section == "Convert Images to PDF":
    st.markdown("### 📷 Convert Images to PDF")
    uploaded_images = st.file_uploader("📤 Upload JPG/JPEG/PNG Images to Convert to PDF", type=["jpg","jpeg","png"], accept_multiple_files=True, key=f"image_uploader_{st.session_state.clear_counter}")
    if uploaded_images:
        display_uploaded_files(uploaded_images, "Images")
        if st.button("🗑️ Clear All Images", key="clear_images"):
            st.session_state.clear_counter += 1
            st.rerun()
        show_progress = st.checkbox("Show progress bar", value=True)
        quality = st.selectbox("Quality", ["High", "Medium", "Low"])
        if st.button("🖨️ Create PDF (Original)"):
            try:
                if show_progress:
                    prog = progress_placeholder.progress(0)
                    status = st.empty()
                imgs = []
                for i, f in enumerate(uploaded_images):
                    try:
                        img = Image.open(f).convert("RGB")
                        if quality == "Medium":
                            img = img.resize((int(img.width*0.8), int(img.height*0.8)))
                        elif quality == "Low":
                            img = img.resize((int(img.width*0.6), int(img.height*0.6)))
                        imgs.append(img)
                    except Exception as e:
                        st.write(f"Skip {f.name}: {e}")
                    if show_progress:
                        prog.progress((i+1)/len(uploaded_images))
                        status.text(f"Prepared {i+1}/{len(uploaded_images)}")
                if len(imgs) == 0:
                    st.warning("No valid images.")
                else:
                    buf = BytesIO()
                    imgs[0].save(buf, format="PDF", save_all=True, append_images=imgs[1:])
                    buf.seek(0)
                    st.success("✅ PDF ready")
                    st.download_button("📥 Download PDF", data=buf.getvalue(), file_name="Images_Combined.pdf", mime="application/pdf")
            except Exception as e:
                st.error(f"❌ Error creating PDF: {e}")
            finally:
                progress_placeholder.empty()
    else:
        st.info("Upload images to convert.")

# ---------------- Dashboard ----------------
elif st.session_state.active_section == "Dashboard":
    st.markdown("### 📊 Interactive Auto Dashboard")
    dashboard_file = st.file_uploader("📊 Upload Excel or CSV File for Dashboard", type=["xlsx","csv"], key=f"dashboard_uploader_{st.session_state.clear_counter}")
    if dashboard_file:
        display_uploaded_files([dashboard_file])
        if st.button("🗑️ Clear Dashboard File", key="clear_dashboard"):
            st.session_state.clear_counter += 1
            st.rerun()
        try:
            if dashboard_file.name.lower().endswith(".csv"):
                df0 = pd.read_csv(dashboard_file)
                sheet_title = "CSV Data"
            else:
                df_dict = pd.read_excel(dashboard_file, sheet_name=None)
                sheet_names = list(df_dict.keys())
                sel = st.selectbox("Select sheet", sheet_names, key="sheet_dash")
                df0 = df_dict[sel].copy()
                sheet_title = sel
            st.dataframe(df0.head(), use_container_width=True)
            numeric_cols = df0.select_dtypes(include='number').columns.tolist()
            if not numeric_cols:
                st.warning("No numeric columns found.")
            else:
                measure_col = st.selectbox("Select value column", numeric_cols)
                # filters
                cat_cols = [c for c in df0.columns if df0[c].dtype == object or df0[c].dtype.name.startswith('category')]
                st.sidebar.header("Filters")
                filter_col = None
                if cat_cols:
                    filter_col = st.sidebar.selectbox("Filter column", ["-- None --"] + cat_cols)
                    if filter_col == "-- None --":
                        filter_col = None
                filter_vals = None
                if filter_col:
                    vals = sorted(df0[filter_col].dropna().astype(str).unique().tolist())
                    filter_vals = st.sidebar.multiselect(f"Filter values for {filter_col}", vals, default=vals)
                filtered = df0.copy()
                if filter_col and filter_vals:
                    filtered = filtered[filtered[filter_col].astype(str).isin(filter_vals)]
                st.dataframe(filtered.head(200), use_container_width=True)
                # KPIs
                total = filtered[measure_col].sum()
                avg = filtered[measure_col].mean()
                cnt = filtered.shape[0]
                c1, c2, c3 = st.columns(3)
                c1.metric("Total", f"{total:,.2f}")
                c2.metric("Average", f"{avg:,.2f}")
                c3.metric("Records", f"{cnt}")
                # simple chart
                group_col = st.selectbox("Optional: group by", ["-- None --"] + [c for c in df0.columns if df0[c].dtype == object])
                chart_buf = None
                if group_col and group_col != "-- None --":
                    try:
                        grp = filtered.groupby(group_col)[measure_col].sum().sort_values(ascending=False).head(10)
                        fig, ax = plt.subplots(figsize=(8,4))
                        ax.bar(grp.index.astype(str), grp.values)
                        ax.set_xticklabels(grp.index.astype(str), rotation=45, ha="right")
                        plt.tight_layout()
                        buf = BytesIO()
                        fig.savefig(buf, format="png")
                        buf.seek(0)
                        st.image(buf)
                        chart_buf = buf
                    except Exception as e:
                        st.warning(f"Chart failed: {e}")
                # export simple PDF
                if st.button("Generate simple report (PDF)"):
                    try:
                        fig, ax = plt.subplots(figsize=(8.27, 11.69))
                        ax.axis("off")
                        text = f"Averroes Pharma Report\n\nFile: {dashboard_file.name}\n\nTotal: {total:,.2f}\nAverage: {avg:,.2f}\nRecords: {cnt}"
                        ax.text(0.01, 0.95, text, fontsize=12, va="top")
                        if chart_buf:
                            img = Image.open(chart_buf)
                            ax_im = fig.add_axes([0.1, 0.2, 0.8, 0.6])
                            ax_im.imshow(img)
                            ax_im.axis("off")
                        pdf_bytes = BytesIO()
                        fig.savefig(pdf_bytes, format="pdf")
                        pdf_bytes.seek(0)
                        st.success("Report ready")
                        st.download_button("📥 Download Report PDF", data=pdf_bytes.getvalue(), file_name="report.pdf", mime="application/pdf")
                    except Exception as e:
                        st.error(f"Could not generate report: {e}")
        except Exception as e:
            st.error(f"❌ Dashboard error: {e}")
        finally:
            progress_placeholder.empty()
    else:
        st.info("Upload a file to generate dashboard.")

# ---------------- Sidebar footer info ----------------
st.sidebar.markdown("---")
st.sidebar.markdown("### ℹ️ Info")
st.sidebar.markdown("By **Mohamed Abd ELGhany**")
st.sidebar.markdown("[Contact via WhatsApp](https://wa.me/201554694554)")
