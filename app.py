# Writing a new simplified, stable Streamlit app to /mnt/data/app.py per the user's request.
code = r'''# -*- coding: utf-8 -*-
"""
Averroes Pharma - Stable Light Version (app.py)
Features:
- Sidebar navigation (Split, Merge, Images->PDF, Dashboard)
- Progress bars in each operation (clear and simple)
- Defensive try/except with banners to avoid crashing on Streamlit Cloud
- Minimal external dependencies for stability: pandas, openpyxl, pillow, matplotlib (optional)
Instructions:
- Upload this file as app.py to your Streamlit repo and deploy.
- Recommended requirements.txt: streamlit, pandas, openpyxl, pillow, matplotlib
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
from PIL import Image
import os
import matplotlib.pyplot as plt
import datetime

st.set_page_config(page_title="Averroes Pharma - Stable", page_icon="💊", layout="wide")

# Simple styling using Streamlit native components - keep it light and clear.
st.title("💊 Averroes Pharma — Stable Light App")
st.markdown("A lightweight, stable version — Sidebar navigation + clear progress bars.")

# Sidebar navigation
st.sidebar.title("Navigation")
section = st.sidebar.radio("Choose section", ("Split Files", "Merge Files", "Images → PDF", "Dashboard"))

# Helper functions
def banner(msg, level="info"):
    if level == "success":
        st.success(msg)
    elif level == "warning":
        st.warning(msg)
    elif level == "error":
        st.error(msg)
    else:
        st.info(msg)

def safe_read_excel(file):
    try:
        return pd.read_excel(file)
    except Exception:
        try:
            return pd.read_excel(file, engine="openpyxl")
        except Exception as e:
            banner(f"Could not read Excel file: {e}", level="error")
            return pd.DataFrame()

def safe_read_csv(file):
    try:
        return pd.read_csv(file)
    except Exception:
        try:
            return pd.read_csv(file, encoding="utf-8", errors="replace")
        except Exception as e:
            banner(f"Could not read CSV file: {e}", level="error")
            return pd.DataFrame()

def safe_save_bytesio_excel(df):
    buf = BytesIO()
    try:
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        buf.seek(0)
    except Exception as e:
        banner(f"Could not write Excel buffer: {e}", level="error")
        return None
    return buf

# -------------------- Split Files --------------------
if section == "Split Files":
    st.header("✂ Split Excel/CSV file by column")
    uploaded = st.file_uploader("Upload an Excel (.xlsx) or CSV file to split", type=["xlsx","csv"], accept_multiple_files=False)
    if uploaded is not None:
        try:
            ext = uploaded.name.split('.')[-1].lower()
            if ext == "csv":
                df = safe_read_csv(uploaded)
            else:
                df = safe_read_excel(uploaded)
            if df.empty:
                banner("Uploaded file contains no data or could not be read.", level="error")
            else:
                st.write("Preview:")
                st.dataframe(df.head(20))
                col = st.selectbox("Select column to split by", options=df.columns.tolist())
                add_timestamp = st.checkbox("Append date to output filenames", value=True)
                show_progress = st.checkbox("Show progress bar", value=True)
                if st.button("Start Split"):
                    unique_vals = df[col].dropna().unique().tolist()
                    if len(unique_vals) == 0:
                        banner("No values to split on in selected column.", level="warning")
                    else:
                        zip_buf = BytesIO()
                        with ZipFile(zip_buf, "w") as zf:
                            if show_progress:
                                prog = st.progress(0)
                                status = st.empty()
                            for i, val in enumerate(unique_vals):
                                try:
                                    subset = df[df[col] == val]
                                    safe_name = "".join(c if c.isalnum() else "_" for c in str(val))[:50] or "value"
                                    if add_timestamp:
                                        safe_name = f"{safe_name}_{datetime.date.today().isoformat()}"
                                    file_name = f"{safe_name}.xlsx"
                                    buf = safe_save_bytesio_excel(subset)
                                    if buf:
                                        zf.writestr(file_name, buf.read())
                                    if show_progress:
                                        prog.progress((i+1)/len(unique_vals))
                                        status.text(f"Created {file_name} ({len(subset)} rows)")
                                except Exception as e:
                                    st.write(f"Skipping value {val} due to error: {e}")
                        zip_buf.seek(0)
                        banner("Splitting completed ✅", level="success")
                        st.download_button("Download split files (zip)", data=zip_buf.getvalue(), file_name=f"split_{uploaded.name.rsplit('.',1)[0]}.zip", mime="application/zip")
        except Exception as e:
            banner(f"Error during split: {e}", level="error")

# -------------------- Merge Files --------------------
elif section == "Merge Files":
    st.header("🔀 Merge multiple Excel/CSV files")
    files = st.file_uploader("Upload multiple Excel/CSV files to merge", type=["xlsx","csv"], accept_multiple_files=True)
    if files:
        st.write("Uploaded files:")
        for f in files:
            st.write(f"- {f.name} ({f.size//1024} KB)")
        show_progress = st.checkbox("Show progress bar", value=True)
        add_source_col = st.checkbox("Add Source_File column to merged data", value=True)
        if st.button("Start Merge"):
            try:
                dfs = []
                total = len(files)
                if show_progress:
                    prog = st.progress(0)
                    status = st.empty()
                for i, f in enumerate(files):
                    try:
                        ext = f.name.split('.')[-1].lower()
                        if ext == "csv":
                            d = safe_read_csv(f)
                        else:
                            d = safe_read_excel(f)
                        if add_source_col:
                            d["Source_File"] = f.name
                        dfs.append(d)
                    except Exception as e:
                        st.write(f"Could not read {f.name}: {e}")
                    if show_progress:
                        prog.progress((i+1)/total)
                        status.text(f"Processed {i+1}/{total}")
                if len(dfs) == 0:
                    banner("No files could be read to merge.", level="warning")
                else:
                    merged = pd.concat(dfs, ignore_index=True)
                    buf = safe_save_bytesio_excel(merged)
                    if buf:
                        banner("Merge completed ✅", level="success")
                        st.download_button("Download merged file", data=buf.getvalue(), file_name="merged_consolidated.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                banner(f"Error during merge: {e}", level="error")

# -------------------- Images to PDF --------------------
elif section == "Images → PDF":
    st.header("📷 Convert Images to single PDF")
    imgs = st.file_uploader("Upload images (jpg/jpeg/png) - multiple allowed", type=["jpg","jpeg","png"], accept_multiple_files=True)
    if imgs:
        st.write(f"{len(imgs)} files uploaded")
        show_progress = st.checkbox("Show progress bar", value=True)
        quality = st.selectbox("PDF image quality (affects size)", ("High","Medium","Low"))
        if st.button("Create PDF"):
            try:
                pil_images = []
                total = len(imgs)
                if show_progress:
                    prog = st.progress(0)
                    status = st.empty()
                for i, f in enumerate(imgs):
                    try:
                        img = Image.open(f).convert("RGB")
                        if quality == "Medium":
                            img = img.resize((int(img.width*0.8), int(img.height*0.8)))
                        elif quality == "Low":
                            img = img.resize((int(img.width*0.6), int(img.height*0.6)))
                        pil_images.append(img)
                    except Exception as e:
                        st.write(f"Could not open {f.name}: {e}")
                    if show_progress:
                        prog.progress((i+1)/total)
                        status.text(f"Prepared {i+1}/{total}")
                if len(pil_images) == 0:
                    banner("No valid images to convert.", level="warning")
                else:
                    pdf_buf = BytesIO()
                    pil_images[0].save(pdf_buf, format="PDF", save_all=True, append_images=pil_images[1:])
                    pdf_buf.seek(0)
                    banner("PDF created ✅", level="success")
                    st.download_button("Download combined PDF", data=pdf_buf.getvalue(), file_name="images_combined.pdf", mime="application/pdf")
            except Exception as e:
                banner(f"Error creating PDF: {e}", level="error")

# -------------------- Dashboard --------------------
elif section == "Dashboard":
    st.header("📊 Simple Auto Dashboard (safe)")
    data_file = st.file_uploader("Upload Excel or CSV for dashboard", type=["xlsx","csv"], accept_multiple_files=False)
    if data_file is not None:
        try:
            ext = data_file.name.split('.')[-1].lower()
            if ext == "csv":
                df = safe_read_csv(data_file)
            else:
                df = safe_read_excel(data_file)
            if df.empty:
                banner("No data in uploaded file.", level="warning")
            else:
                st.write("Preview:")
                st.dataframe(df.head(50))
                numeric_cols = df.select_dtypes(include='number').columns.tolist()
                if len(numeric_cols) == 0:
                    banner("No numeric columns found for KPIs/charts.", level="warning")
                else:
                    value_col = st.selectbox("Select value column for KPIs/charts", numeric_cols)
                    # Simple KPIs
                    total = df[value_col].sum()
                    avg = df[value_col].mean()
                    count = df.shape[0]
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Total", f"{total:,.2f}")
                    col2.metric("Average", f"{avg:,.2f}")
                    col3.metric("Records", f"{count}")
                    # group by optional
                    group_col = st.selectbox("Optional: Group by column", ["-- None --"] + [c for c in df.columns if df[c].dtype == object])
                    chart_buf = None
                    if group_col != "-- None --":
                        try:
                            grp = df.groupby(group_col)[value_col].sum().sort_values(ascending=False).head(10)
                            # plot via matplotlib for stability
                            fig, ax = plt.subplots(figsize=(8,4))
                            ax.bar(grp.index.astype(str), grp.values)
                            ax.set_xticklabels(grp.index.astype(str), rotation=45, ha="right")
                            ax.set_title(f"Top by {group_col}")
                            plt.tight_layout()
                            buf = BytesIO()
                            fig.savefig(buf, format="png")
                            buf.seek(0)
                            st.image(buf)
                            chart_buf = buf
                        except Exception as e:
                            banner(f"Chart failed: {e}", level="warning")
                    # Generate simple PDF report with chart + KPIs if user wants (basic using PIL & simple layout)
                    if st.button("Generate simple report (PDF)"):
                        try:
                            # Build a very simple one-page PDF using matplotlib figure saved as PDF
                            # We'll create a figure that contains KPIs and optional chart
                            fig, ax = plt.subplots(figsize=(8.27, 11.69))  # A4 portrait
                            ax.axis("off")
                            text = f"Averroes Pharma Report\n\nFile: {data_file.name}\n\nTotal: {total:,.2f}\nAverage: {avg:,.2f}\nRecords: {count}"
                            ax.text(0.01, 0.95, text, fontsize=12, va="top")
                            if chart_buf is not None:
                                try:
                                    img = Image.open(chart_buf)
                                    # place image below text
                                    ax_im = fig.add_axes([0.1, 0.2, 0.8, 0.6])
                                    ax_im.imshow(img)
                                    ax_im.axis("off")
                                except Exception:
                                    pass
                            pdf_bytes = BytesIO()
                            fig.savefig(pdf_bytes, format="pdf")
                            pdf_bytes.seek(0)
                            banner("Report generated ✅", level="success")
                            st.download_button("Download report PDF", data=pdf_bytes.getvalue(), file_name="report.pdf", mime="application/pdf")
                        except Exception as e:
                            banner(f"Could not generate report: {e}", level="error")
        except Exception as e:
            banner(f"Error reading dashboard file: {e}", level="error")

# Footer
st.markdown("---")
st.markdown("Built for Averroes Pharma — Stable Light Version. If any section fails, take a screenshot and send it so I can fix the specific line.")'''

with open('/mnt/data/app.py', 'w', encoding='utf-8') as f:
    f.write(code)

"/mnt/data/app.py written"
