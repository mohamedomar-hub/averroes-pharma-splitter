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
    page_title="Averroes Pharma Splitter",
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
        <a href="#" onclick="window.location.reload()">Home</a>
        <a href="https://wa.me/201554694554" target="_blank">Contact</a>
        <a href="#info-section">Info</a>
    </div>
    """,
    unsafe_allow_html=True
)

# ------------------ عرض اللوجو (مع حماية من الخطأ) ------------------
logo_path = "logo.png"
if os.path.exists(logo_path):
    st.markdown('<div style="text-align:center; margin:20px 0;">', unsafe_allow_html=True)
    st.image(logo_path, width=200)
    st.markdown('</div>', unsafe_allow_html=True)
else:
    # إذا لم يكن اللوجو موجودًا، نعرض نصًا بديلًا
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
st.markdown("<h1 style='text-align:center; color:#FFD700;'>💊 Averroes Pharma File Splitter</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align:center; color:white;'>✂ Divide your files easily and accurately.</h3>", unsafe_allow_html=True)

# ------------------ رفع الملف (التقسيم) ------------------
uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"], accept_multiple_files=False)

if uploaded_file:
    try:
        input_bytes = uploaded_file.getvalue()
        original_wb = load_workbook(filename=BytesIO(input_bytes), data_only=False)
        st.success(f"✅ The file has been uploaded successfully. Number of sheets: {len(original_wb.sheetnames)}")

        selected_sheet = st.selectbox("Select Sheet", original_wb.sheetnames)

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

            # --- زر التقسيم ---
            if st.button("🚀 Start Split"):
                with st.spinner("Splitting process in progress while preserving original format..."):

                    def clean_name(name):
                        name = str(name).strip()
                        invalid_chars = r'[\\/*?:\[\]|<>"]'
                        cleaned = re.sub(invalid_chars, '_', name)
                        return cleaned[:30] if cleaned else "Sheet"

                    ws = original_wb[selected_sheet]
                    col_idx = df.columns.get_loc(col_to_split) + 1  # رقم العمود (1-based)
                    unique_values = df[col_to_split].dropna().unique()

                    zip_buffer = BytesIO()
                    with ZipFile(zip_buffer, "w") as zip_file:
                        for value in unique_values:
                            # --- إنشاء ملف جديد ---
                            new_wb = Workbook()
                            default_ws = new_wb.active
                            new_wb.remove(default_ws)
                            new_ws = new_wb.create_sheet(title=clean_name(value))

                            # --- نسخ الرأس ---
                            for cell in ws[1]:
                                dst_cell = new_ws.cell(1, cell.column, cell.value)
                                if cell.has_style:
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

                            # --- نسخ الصفوف ---
                            row_idx = 2
                            for row in ws.iter_rows(min_row=2):
                                cell_in_col = row[col_idx - 1]
                                if cell_in_col.value == value:
                                    for src_cell in row:
                                        dst_cell = new_ws.cell(row_idx, src_cell.column, src_cell.value)
                                        if src_cell.has_style:
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
                                    row_idx += 1

                            # --- نسخ عرض الأعمدة ---
                            for col_letter in ws.column_dimensions:
                                new_ws.column_dimensions[col_letter].width = ws.column_dimensions[col_letter].width

                            # --- حفظ الملف ---
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
                        file_name=f"Split_{re.sub(r'[^A-Za-z0-9_-]+','_', uploaded_file.name.rsplit('.',1)[0])}.zip",
                        mime="application/zip"
                    )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.markdown("<p style='text-align:center; color:#FFD700;'>⚠️ No file uploaded yet.</p>", unsafe_allow_html=True)

# -----------------------------------------------
# 🔄 دمج ملفات Excel - مستقل ومحفوظ التنسيق
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
                # إنشاء ملف مدمج جديد
                combined_wb = Workbook()
                combined_ws = combined_wb.active
                combined_ws.title = "Consolidated"

                # قراءة أول ملف لنسخ تنسيق الرأس وعرض الأعمدة
                first_file = merge_files[0]
                temp_wb = load_workbook(filename=BytesIO(first_file.getvalue()), data_only=False)
                temp_ws = temp_wb.active

                # نسخ صف الرأس مع التنسيق الكامل
                for cell in temp_ws[1]:
                    new_cell = combined_ws.cell(1, cell.column, cell.value)
                    if cell.has_style:
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

                # نسخ عرض الأعمدة من أول ملف
                for col_letter in temp_ws.column_dimensions:
                    combined_ws.column_dimensions[col_letter].width = temp_ws.column_dimensions[col_letter].width

                # بدء من الصف الثاني
                row_idx = 2
                for file in merge_files:
                    wb = load_workbook(filename=BytesIO(file.getvalue()), data_only=True)
                    ws = wb.active
                    for row in ws.iter_rows(min_row=2):
                        for cell in row:
                            if cell.value is not None:
                                new_cell = combined_ws.cell(row_idx, cell.column, cell.value)
                                # نسخ التنسيق إن وُجد
                                if cell.has_style:
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
                        row_idx += 1

                # حفظ الملف المدمج
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
# 📊 قسم جديد: Interactive Dashboard + PDF Report (عنوان مطابق لاسم الشيت)
# ====================================================================================
st.markdown("<hr class='divider' id='dashboard-section'>", unsafe_allow_html=True)
st.markdown("### 📊 Interactive Dashboard Generator")

dashboard_file = st.file_uploader("📊 Upload Excel File for Dashboard", type=["xlsx"], key="dashboard_uploader")

def _find_col(df, aliases):
    lowered = {c.lower(): c for c in df.columns}
    for a in aliases:
        if a.lower() in lowered:
            return lowered[a.lower()]
    # محاولات تقريبية
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

def make_bar(fig_ax, series, title, ylabel):
    ax = fig_ax
    bars = ax.bar(series.index.astype(str), series.values)
    ax.set_title(title)
    ax.set_ylabel(ylabel)
    ax.yaxis.set_major_formatter(FuncFormatter(_format_millions))
    ax.tick_params(axis='x', rotation=0)
    for b in bars:
        h = b.get_height()
        ax.annotate(f"{h:,.0f}", xy=(b.get_x()+b.get_width()/2, h),
                    xytext=(0, 5), textcoords="offset points", ha='center', va='bottom', fontsize=9)

def make_pie(fig_ax, series, title):
    ax = fig_ax
    wedges, texts, autotexts = ax.plot([],[]) ,[],[]  # placeholder
    total = series.sum()
    autopct = lambda pct: f"{pct:.1f}%\n({(pct/100.0)*total:,.0f})"
    ax.clear()
    ax.pie(series.values, labels=series.index.astype(str), autopct=autopct, startangle=90)
    ax.set_title(title)
    ax.axis('equal')

def make_line(fig_ax, series, title, ylabel):
    ax = fig_ax
    ax.plot(series.index.astype(str), series.values, marker='o')
    ax.set_title(title)
    ax.set_ylabel(ylabel)
    ax.grid(True, linestyle='--', alpha=0.4)
    for x, y in zip(range(len(series.index)), series.values):
        ax.annotate(f"{y:,.0f}", xy=(x, y), xytext=(0, 6), textcoords="offset points", ha='center', va='bottom', fontsize=9)

def build_pdf(sheet_title, filtered_df, charts_buffers):
    buf = BytesIO()
    # عرض أفقي عشان الرسومات تكون واضحة
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24)
    styles = getSampleStyleSheet()
    elements = []

    # غلاف أنيق بألوان متناسقة
    elements.append(Paragraph(f"<para align='center'><b>{sheet_title}</b></para>", styles['Title']))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph("<para align='center' color='#001F3F'>Averroes Pharma – Interactive Dashboard</para>", styles['Heading3']))
    elements.append(Spacer(1, 12))

    # إدراج الرسومات
    for img_buf, caption in charts_buffers:
        img = Image(img_buf, width=760, height=360)  # تقريباً عرض الصفحة الأفقية
        elements.append(img)
        elements.append(Spacer(1, 6))
        elements.append(Paragraph(f"<para align='center'><font color='#6c757d'>{caption}</font></para>", styles['Normal']))
        elements.append(Spacer(1, 18))

    # جدول البيانات (نقصّم لو كبير)
    table_data = [filtered_df.columns.tolist()] + filtered_df.astype(object).astype(str).values.tolist()
    # لعدم ثقل الملف، نقسم على صفحات كل 25 صف
    chunk = 25
    for i in range(0, len(table_data), chunk):
        part = table_data[i:i+chunk]
        tbl = Table(part, hAlign='CENTER')
        tbl.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#FFD700")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.black),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.whitesmoke, colors.HexColor("#F7F7F7")])
        ]))
        elements.append(tbl)
        if i + chunk < len(table_data):
            elements.append(PageBreak())

    doc.build(elements)
    buf.seek(0)
    return buf

if dashboard_file:
    try:
        df_all = pd.read_excel(dashboard_file, sheet_name=None)
        sheet_names = list(df_all.keys())
        selected_sheet_dash = st.selectbox("Select Sheet for Dashboard", sheet_names, key="sheet_dash")

        if selected_sheet_dash:
            sheet_title = selected_sheet_dash  # العنوان في PDF مطابق لاسم الشيت
            df_dash = df_all[selected_sheet_dash].copy()

            st.markdown("### 🔍 Data Preview")
            st.dataframe(df_dash, use_container_width=True)

            # كشف أعمدة مهمة (شهر/مندوب/مبيعات) تدعم عربي/إنجليزي
            month_col = _find_col(df_dash, ["Month", "الشهر", "month", "MONTH"])
            rep_col   = _find_col(df_dash, ["Rep", "Sales Rep", "المندوب", "مندوب", "representative"])
            sales_col = _find_col(df_dash, ["Sales", "المبيعات", "value", "amount", "NET", "Total"])

            # محاولات تحويل الشهر لتاريخ
            if month_col:
                try:
                    df_dash[month_col] = pd.to_datetime(df_dash[month_col], errors='coerce')
                except Exception:
                    pass

            # سايدبار فلاتر: شهر + مندوب
            st.sidebar.header("🔍 Filters")
            filtered = df_dash.copy()

            # فلتر الشهر (لو عمود الشهر متاح)
            if month_col and pd.api.types.is_datetime64_any_dtype(filtered[month_col]):
                min_d, max_d = filtered[month_col].min(), filtered[month_col].max()
                d_range = st.sidebar.date_input("📅 Date Range", [min_d.date() if pd.notna(min_d) else None,
                                                                  max_d.date() if pd.notna(max_d) else None])
                if isinstance(d_range, list) and len(d_range) == 2 and all(d is not None for d in d_range):
                    start_d = pd.to_datetime(d_range[0])
                    end_d = pd.to_datetime(d_range[1])
                    filtered = filtered[(filtered[month_col] >= start_d) & (filtered[month_col] <= end_d)]
            elif month_col:
                # شهر نصّي: نعرض قيم ونفلتر
                month_vals = filtered[month_col].dropna().astype(str).unique().tolist()
                selected_months = st.sidebar.multiselect("📅 Months", month_vals, default=month_vals)
                filtered = filtered[filtered[month_col].astype(str).isin(selected_months)]

            # فلتر المندوب
            if rep_col:
                reps = filtered[rep_col].dropna().astype(str).unique().tolist()
                selected_reps = st.sidebar.multiselect("🧑‍💼 Representatives", reps, default=reps)
                filtered = filtered[filtered[rep_col].astype(str).isin(selected_reps)]

            st.markdown("### 📈 Filtered Data")
            st.dataframe(filtered, use_container_width=True)

            # تحضير عمود المبيعات (إن لم يوجد، نجمع كل الأعمدة الرقمية كـ إجمالي)
            if sales_col is None:
                num_cols = filtered.select_dtypes(include='number').columns.tolist()
                if len(num_cols):
                    filtered["__auto_sales__"] = filtered[num_cols].sum(axis=1, numeric_only=True)
                    sales_col = "__auto_sales__"

            charts_buffers = []

            if sales_col is not None:
                # 1) Bar: المبيعات حسب المندوب
                if rep_col:
                    sales_by_rep = filtered.groupby(rep_col)[sales_col].sum().sort_values(ascending=False)
                    if len(sales_by_rep):
                        fig, ax = plt.subplots(figsize=(9, 4))
                        make_bar(ax, sales_by_rep, "Sales by Representative", "Total Sales")
                        fig.tight_layout()
                        img_buf = BytesIO()
                        fig.savefig(img_buf, format="png", dpi=200, bbox_inches="tight")
                        img_buf.seek(0)
                        charts_buffers.append((img_buf, "Sales by Representative"))
                        st.pyplot(fig)
                        plt.close(fig)

                        # Pie
                        fig, ax = plt.subplots(figsize=(7, 4))
                        make_pie(ax, sales_by_rep, "Sales Share by Representative")
                        fig.tight_layout()
                        img_buf = BytesIO()
                        fig.savefig(img_buf, format="png", dpi=200, bbox_inches="tight")
                        img_buf.seek(0)
                        charts_buffers.append((img_buf, "Sales Share by Representative"))
                        st.pyplot(fig)
                        plt.close(fig)

                # 2) Line: المبيعات حسب الشهر
                if month_col:
                    mser = filtered.dropna(subset=[month_col])
                    if pd.api.types.is_datetime64_any_dtype(mser[month_col]):
                        mser["_yyyymm"] = mser[month_col].dt.to_period("M")
                        sales_by_month = mser.groupby("_yyyymm")[sales_col].sum().sort_index()
                        sales_by_month.index = sales_by_month.index.astype(str)
                    else:
                        sales_by_month = filtered.groupby(month_col)[sales_col].sum()
                    if len(sales_by_month):
                        fig, ax = plt.subplots(figsize=(9, 4))
                        make_line(ax, sales_by_month, "Sales Trend by Month", "Total Sales")
                        fig.tight_layout()
                        img_buf = BytesIO()
                        fig.savefig(img_buf, format="png", dpi=200, bbox_inches="tight")
                        img_buf.seek(0)
                        charts_buffers.append((img_buf, "Sales Trend by Month"))
                        st.pyplot(fig)
                        plt.close(fig)

            # === تحميل كـ PDF (العنوان مطابق لاسم الشيت + الأرقام داخل الرسوم) ===
            st.markdown("### 💾 Download PDF Report")
            if st.button("📥 Generate PDF Report"):
                with st.spinner("Generating PDF..."):
                    pdf_buffer = build_pdf(sheet_title, filtered.fillna(""), charts_buffers)
                    st.download_button(
                        label="⬇️ Download Dashboard PDF",
                        data=pdf_buffer,
                        file_name=f"{re.sub(r'[^A-Za-z0-9_-]+','_', sheet_title)}.pdf",
                        mime="application/pdf"
                    )

            # --- تحميل كـ Excel للبيانات المفلترة ---
            st.markdown("### 💾 Download Filtered Data (Excel)")
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                filtered.to_excel(writer, index=False, sheet_name='Filtered Data')
            st.download_button(
                label="⬇️ Download Filtered Data.xlsx",
                data=excel_buffer.getvalue(),
                file_name=f"{re.sub(r'[^A-Za-z0-9_-]+','_', sheet_title)}_Filtered.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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
    - اختر شيت (اسم الشيت هيكون عنوان الـ PDF).
    - فلتر بالشهر والمندوب.
    - تشوف الرسومات التفاعلية في التطبيق.
    - حمّل **PDF** فيه الرسومات بالأرقام وجداول البيانات، وكمان **Excel** بالبيانات المفلترة.

    ---

    🙋‍♂️ لأي استفسار: <a href="https://wa.me/201554694554" target="_blank">01554694554 (واتساب)</a>
    """, unsafe_allow_html=True)
