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
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet

# ------------------ ربط بخط عربي جميل (Cairo) ------------------
st.markdown(
    '<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap" rel="stylesheet">',
    unsafe_allow_html=True
)

# ------------------ إعدادات الصفحة ------------------
st.set_page_config(
    page_title="Averroes Pharma Splitter + Dashboard",
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
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

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
st.markdown("<h3 style='text-align:center; color:white;'>✂ Split, Merge and Auto Dashboard Generator</h3>", unsafe_allow_html=True)

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
        return f"{x/1_000:.1f}K"
    return f"{x:.0f}"

def make_bar(ax, series, title, ylabel):
    bars = ax.bar(series.index.astype(str), series.values)
    ax.set_title(title)
    ax.set_ylabel(ylabel)
    ax.yaxis.set_major_formatter(FuncFormatter(_format_millions))
    ax.tick_params(axis='x', rotation=45)
    for b in bars:
        h = b.get_height()
        ax.annotate(f"{h:,.0f}", xy=(b.get_x()+b.get_width()/2, h),
                    xytext=(0, 5), textcoords="offset points", ha='center', va='bottom', fontsize=8)

def make_pie(ax, series, title):
    total = series.sum()
    autopct = lambda pct: f"{pct:.1f}%\n({(pct/100.0)*total:,.0f})"
    ax.pie(series.values, labels=series.index.astype(str), autopct=autopct, startangle=90)
    ax.set_title(title)
    ax.axis('equal')

def make_line(ax, series, title, ylabel):
    ax.plot(series.index.astype(str), series.values, marker='o')
    ax.set_title(title)
    ax.set_ylabel(ylabel)
    ax.grid(True, linestyle='--', alpha=0.4)
    for x, y in zip(range(len(series.index)), series.values):
        ax.annotate(f"{y:,.0f}", xy=(x, y), xytext=(0, 6), textcoords="offset points", ha='center', va='bottom', fontsize=8)

def build_pdf(sheet_title, filtered_df, charts_buffers, max_table_rows=200):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=20, rightMargin=20, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    elements = []

    # Cover
    elements.append(Paragraph(f"<para align='center'><b>{sheet_title} Report</b></para>", styles['Title']))
    elements.append(Spacer(1,12))
    elements.append(Paragraph("<para align='center'>Averroes Pharma - Auto Generated Dashboard</para>", styles['Heading3']))
    elements.append(Spacer(1,18))

    # Charts
    for img_buf, caption in charts_buffers:
        try:
            img = Image(img_buf, width=760, height=360)
            elements.append(img)
            elements.append(Spacer(1,6))
            elements.append(Paragraph(f"<para align='center'>{caption}</para>", styles['Normal']))
            elements.append(Spacer(1,12))
        except Exception:
            pass

    # Table - limit rows to avoid huge PDF
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

# ------------------ Upload for Splitter ------------------
uploaded_file = st.file_uploader("📂 Upload Excel File (Splitter/Merge) — Use this to Split or Merge as before", type=["xlsx"], accept_multiple_files=False)

if uploaded_file:
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

            # --- Split button ---
            if st.button("🚀 Start Split"):
                with st.spinner("Splitting process in progress while preserving original format..."):

                    def clean_name(name):
                        name = str(name).strip()
                        invalid_chars = r'[\\/*?:\[\]|<>"]'
                        cleaned = re.sub(invalid_chars, '_', name)
                        return cleaned[:30] if cleaned else "Sheet"

                    ws = original_wb[selected_sheet]
                    col_idx = df.columns.get_loc(col_to_split) + 1  # 1-based
                    unique_values = df[col_to_split].dropna().unique()

                    zip_buffer = BytesIO()
                    with ZipFile(zip_buffer, "w") as zip_file:
                        for value in unique_values:
                            # create new workbook, preserve header formatting & styles where possible
                            new_wb = Workbook()
                            default_ws = new_wb.active
                            new_wb.remove(default_ws)
                            new_ws = new_wb.create_sheet(title=clean_name(value))

                            # copy header
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

                            # copy rows for this value
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

                            # copy column widths
                            try:
                                for col_letter in ws.column_dimensions:
                                    new_ws.column_dimensions[col_letter].width = ws.column_dimensions[col_letter].width
                            except Exception:
                                pass

                            # save to buffer
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

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.markdown("<p style='text-align:center; color:#FFD700;'>⚠️ No file uploaded yet for splitting.</p>", unsafe_allow_html=True)

# -----------------------------------------------
# Merge area (unchanged)
# -----------------------------------------------
st.markdown("<hr class='divider-dashed'>", unsafe_allow_html=True)
st.markdown("### 🔄 Merge Excel Files (Keep Original Format)")
merge_files = st.file_uploader(
    "📤 Upload Excel Files to Merge",
    type=["xlsx"],
    accept_multiple_files=True,
    key="merge_uploader"
)

if merge_files:
    if st.button("✨ Merge Files with Format"):
        with st.spinner("Merging files while preserving formatting..."):
            try:
                combined_wb = Workbook()
                combined_ws = combined_wb.active
                combined_ws.title = "Consolidated"

                first_file = merge_files[0]
                temp_wb = load_workbook(filename=BytesIO(first_file.getvalue()), data_only=False)
                temp_ws = temp_wb.active

                # copy header with styles
                for cell in temp_ws[1]:
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

                # column widths
                try:
                    for col_letter in temp_ws.column_dimensions:
                        combined_ws.column_dimensions[col_letter].width = temp_ws.column_dimensions[col_letter].width
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
# 📊 Dashboard Generator (Auto unpivot + Dynamic Filters + Auto Charts + PDF/Excel Export)
# ====================================================================================
st.markdown("<hr class='divider' id='dashboard-section'>", unsafe_allow_html=True)
st.markdown("### 📊 Interactive Auto Dashboard Generator")

dashboard_file = st.file_uploader("📊 Upload Excel File for Dashboard (Auto)", type=["xlsx"], key="dashboard_uploader")

if dashboard_file:
    try:
        # read all sheets
        df_dict = pd.read_excel(dashboard_file, sheet_name=None)
        sheet_names = list(df_dict.keys())
        selected_sheet_dash = st.selectbox("Select Sheet for Dashboard", sheet_names, key="sheet_dash")

        if selected_sheet_dash:
            sheet_title = selected_sheet_dash
            df0 = df_dict[selected_sheet_dash].copy()

            st.markdown("### 🔍 Data Preview (original)")
            st.dataframe(df0.head(), use_container_width=True)

            # ---- Detect categorical and numeric columns ----
            # Consider month-like columns (Jan..Dec) and numeric columns as measures
            # Heuristic for month columns:
            month_names = ["jan","feb","mar","apr","may","jun","jul","aug","sep","sept","oct","nov","dec"]
            cols_lower = [c.strip().lower() for c in df0.columns]
            potential_months = [c for c in df0.columns if c.strip().lower() in month_names]

            # numeric columns (measures)
            numeric_cols = df0.select_dtypes(include='number').columns.tolist()

            # if month-like columns exist, we'll melt them
            if potential_months:
                id_vars = [c for c in df0.columns if c not in potential_months]
                value_vars = potential_months
                # unpivot: month-name -> Month, value -> Sales
                df_long = df0.melt(id_vars=id_vars, value_vars=value_vars, var_name="Month", value_name="Value")
                # normalize month strings
                df_long["Month"] = df_long["Month"].astype(str)
                measure_col = "Value"
            else:
                # If no month columns, if there is exactly one numeric column use it as measure.
                # Otherwise, create an aggregated measure by summing numeric columns per row (auto sales)
                if len(numeric_cols) == 1:
                    measure_col = numeric_cols[0]
                    df_long = df0.copy()
                elif len(numeric_cols) > 1:
                    df_long = df0.copy()
                    df_long["__auto_sales__"] = df_long[numeric_cols].sum(axis=1, numeric_only=True)
                    measure_col = "__auto_sales__"
                else:
                    # fallback: try to coerce any column that looks like amounts
                    measure_col = None
                    df_long = df0.copy()

            # Detect categorical columns (object / low-cardinality)
            cat_cols = [c for c in df_long.columns if df_long[c].dtype == "object" or df_long[c].dtype.name.startswith("category")]
            # Also include any columns with low unique values (<=100) that are not numeric
            for c in df_long.columns:
                if c not in cat_cols and df_long[c].nunique(dropna=True) <= 100 and df_long[c].dtype != "float64" and df_long[c].dtype != "int64":
                    cat_cols.append(c)
            # Make sure index column not in list
            cat_cols = [c for c in cat_cols if c is not None]

            # Sidebar dynamic filters: create multiselects for a subset of categorical columns
            st.sidebar.header("🔍 Dynamic Filters")
            filter_cols = st.sidebar.multiselect("Choose filter columns (the dashboard will update automatically)", cat_cols, default=cat_cols[:3] if len(cat_cols)>0 else [])
            active_filters = {}
            for fc in filter_cols:
                opts = df_long[fc].dropna().astype(str).unique().tolist()
                # sort options
                try:
                    opts = sorted(opts)
                except Exception:
                    pass
                selected = st.sidebar.multiselect(f"Filter: {fc}", opts, default=opts)
                active_filters[fc] = selected

            # Apply filters
            filtered = df_long.copy()
            for fc, sel in active_filters.items():
                if sel is not None and len(sel) > 0:
                    filtered = filtered[filtered[fc].astype(str).isin(sel)]

            st.markdown("### 📈 Filtered Data Preview")
            st.dataframe(filtered.head(200), use_container_width=True)

            # Auto KPIs
            st.markdown("### 🚀 KPIs")
            if measure_col and measure_col in filtered.columns:
                total_sales = filtered[measure_col].sum()
                avg_sales = filtered[measure_col].mean()
                count_rows = len(filtered)
            else:
                total_sales = None
                avg_sales = None
                count_rows = len(filtered)

            c1, c2, c3 = st.columns(3)
            c1.metric("Total (Measure)", f"{total_sales:,.0f}" if total_sales is not None else "-")
            c2.metric("Average (Measure)", f"{avg_sales:,.0f}" if avg_sales is not None else "-")
            c3.metric("Rows (filtered)", f"{count_rows}")

            # Auto charts:
            st.markdown("### 📊 Auto Charts (built from data)")

            charts_buffers = []

            # 1) If there is a categorical column besides Month, choose the top categorical for breakdown
            # Find best categorical for breakdown: the one with most non-null values and reasonable cardinality
            possible_dims = [c for c in filtered.columns if c not in [measure_col, "Month"]]
            # prefer columns that look like "Item", "Product", "Area", "Branch", "Manager"
            prefer_order = ["item","product","area","branch","manager","rep","representative","salesman","brick"]
            chosen_dim = None
            for p in prefer_order:
                for c in possible_dims:
                    if p in c.lower():
                        chosen_dim = c
                        break
                if chosen_dim:
                    break
            if not chosen_dim and len(possible_dims):
                # pick the categorical with smallest cardinality >1
                lens = [(c, filtered[c].nunique(dropna=True)) for c in possible_dims]
                lens = sorted([x for x in lens if x[1] > 1], key=lambda x: x[1])
                if lens:
                    chosen_dim = lens[0][0]

            # Chart A: Top N by chosen dimension
            if chosen_dim and measure_col and chosen_dim in filtered.columns:
                try:
                    topN = 10
                    series = filtered.groupby(chosen_dim)[measure_col].sum().sort_values(ascending=False).head(topN)
                    fig, ax = plt.subplots(figsize=(9,4))
                    make_bar(ax, series, f"Top {len(series)} by {chosen_dim}", "Total")
                    fig.tight_layout()
                    img_buf = BytesIO()
                    fig.savefig(img_buf, format="png", dpi=150, bbox_inches="tight")
                    img_buf.seek(0)
                    charts_buffers.append((img_buf, f"Top {len(series)} by {chosen_dim}"))
                    st.pyplot(fig)
                    plt.close(fig)

                    # pie of same series
                    fig, ax = plt.subplots(figsize=(7,4))
                    make_pie(ax, series, f"Share by {chosen_dim}")
                    fig.tight_layout()
                    img_buf = BytesIO()
                    fig.savefig(img_buf, format="png", dpi=150, bbox_inches="tight")
                    img_buf.seek(0)
                    charts_buffers.append((img_buf, f"Share by {chosen_dim}"))
                    st.pyplot(fig)
                    plt.close(fig)
                except Exception:
                    pass

            # Chart B: Trend by Month if available
            if "Month" in filtered.columns and measure_col and measure_col in filtered.columns:
                try:
                    ser = filtered.dropna(subset=["Month"])
                    # If Month is period or datetime, convert to string for ordering
                    if pd.api.types.is_datetime64_any_dtype(ser["Month"]):
                        ser["_yyyymm"] = ser["Month"].dt.to_period("M")
                        trend = ser.groupby("_yyyymm")[measure_col].sum().sort_index()
                        trend.index = trend.index.astype(str)
                    else:
                        trend = ser.groupby("Month")[measure_col].sum().sort_index()
                    if len(trend):
                        fig, ax = plt.subplots(figsize=(9,4))
                        make_line(ax, trend, "Trend by Month", "Total")
                        fig.tight_layout()
                        img_buf = BytesIO()
                        fig.savefig(img_buf, format="png", dpi=150, bbox_inches="tight")
                        img_buf.seek(0)
                        charts_buffers.append((img_buf, "Trend by Month"))
                        st.pyplot(fig)
                        plt.close(fig)
                except Exception:
                    pass

            # Chart C: Overall distribution by another categorical if available
            # pick second categorical
            second_dim = None
            for c in possible_dims:
                if c != chosen_dim:
                    second_dim = c
                    break
            if second_dim and measure_col and second_dim in filtered.columns:
                try:
                    series2 = filtered.groupby(second_dim)[measure_col].sum().sort_values(ascending=False).head(10)
                    fig, ax = plt.subplots(figsize=(9,4))
                    make_bar(ax, series2, f"By {second_dim}", "Total")
                    fig.tight_layout()
                    img_buf = BytesIO()
                    fig.savefig(img_buf, format="png", dpi=150, bbox_inches="tight")
                    img_buf.seek(0)
                    charts_buffers.append((img_buf, f"By {second_dim}"))
                    st.pyplot(fig)
                    plt.close(fig)
                except Exception:
                    pass

            # === PDF / Excel Export ===
            st.markdown("### 💾 Export Report / Data")

            # Excel of filtered data
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

            # PDF generation (includes charts captured above and first rows of table)
            if st.button("📥 Generate PDF Report"):
                with st.spinner("Generating PDF..."):
                    pdf_buffer = build_pdf(sheet_title, filtered, charts_buffers)
                    st.download_button(
                        label="⬇️ Download Dashboard PDF",
                        data=pdf_buffer,
                        file_name=f"{_safe_name(sheet_title)}.pdf",
                        mime="application/pdf"
                    )

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
    1. ارفع ملف Excel.
    2. اختر الشيت.
    3. اختر العمود اللي عاوز تقسّم عليه (مثل: "Area Manager").
    4. اضغط على **"ابدأ التقسيم"**.
    5. هيطلعلك **ملف ZIP يحتوي على ملف منفصل لكل قيمة**.

    ✅ كل ملف يحتوي على البيانات الخاصة بهذه القيمة فقط.

    ---

    ### 🔗 ثانيًا: الدمج
    - ارفع أكثر من ملف Excel.
    - اضغط "ادمج الملفات".
    - هتلاقي زر لتحميل ملف واحد فيه كل البيانات مع **الحفاظ على التنسيق الأصلي**.

    ---

    ### 📊 ثالثًا: الـ Dashboard
    - ارفع ملف Excel.
    - اختر الشيت (اسم الشيت هيكون عنوان الـ PDF).
    - استخدم Sidebar لاختيار أعمدة الفلترة الديناميكية (Drop-lists).
    - الداشبورد يتبنى أوتوماتيك الرسومات من الداتا.
    - حمّل **PDF** فيه الرسومات بالأرقام وجداول البيانات، وكمان **Excel** بالبيانات المفلترة.

    ---

    🙋‍♂️ لأي استفسار: <a href="https://wa.me/201554694554" target="_blank">01554694554 (واتساب)</a>
    """, unsafe_allow_html=True)
