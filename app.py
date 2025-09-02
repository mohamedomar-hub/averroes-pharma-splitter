import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
import re
import os
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl import load_workbook, Workbook

# ------------------ إضافات جديدة للداش بورد ------------------
import plotly.express as px
import plotly.graph_objects as go
from fpdf2 import FPDF  # تثبيت: pip install fpdf2
import matplotlib.pyplot as plt

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
        <a href="  https://wa.me/201554694554  " target="_blank">Contact</a>
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
        <a href="https://wa.me/201554694554  " target="_blank" style="color:#FFD700; text-decoration:none;">
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
                        file_name=f"Split_{clean_name(uploaded_file.name.rsplit('.',1)[0])}.zip",
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
# 📊 قسم جديد: Interactive Dashboard Generator
# ====================================================================================
st.markdown("<hr class='divider' id='dashboard-section'>", unsafe_allow_html=True)
st.markdown("### 📊 Interactive Dashboard Generator")

dashboard_file = st.file_uploader("📊 Upload Excel File for Dashboard", type=["xlsx"], key="dashboard_uploader")

if dashboard_file:
    try:
        df_dash = pd.read_excel(dashboard_file, sheet_name=None)
        sheet_names = list(df_dash.keys())
        selected_sheet_dash = st.selectbox("Select Sheet for Dashboard", sheet_names, key="sheet_dash")

        if selected_sheet_dash:
            df = df_dash[selected_sheet_dash].copy()

            # عرض البيانات
            st.markdown("### 🔍 Data Preview")
            st.dataframe(df, use_container_width=True)

            # --- تحويل الأعمدة الزمنية ---
            date_columns = df.select_dtypes(include='datetime').columns.tolist()
            for col in df.columns:
                if col not in date_columns:
                    # محاولة تحويل إلى تاريخ
                    try:
                        if pd.to_datetime(df[col], errors='raise').dtype == 'datetime64[ns]':
                            df[col] = pd.to_datetime(df[col])
                            date_columns.append(col)
                    except:
                        pass

            # --- الفلاتر في الـ sidebar ---
            st.sidebar.header("🔍 Filters")
            filters = {}

            for col in df.columns:
                if col in date_columns:
                    min_date = df[col].min().date()
                    max_date = df[col].max().date()
                    start, end = st.sidebar.date_input(f"Date Range: {col}", [min_date, max_date])
                    filters[col] = (pd.to_datetime(start), pd.to_datetime(end))
                elif df[col].nunique() < 50:  # فئات صغيرة (مثل موظفين، مديرين)
                    options = df[col].dropna().unique().tolist()
                    selected = st.sidebar.multiselect(f"Filter by: {col}", options, default=options)
                    filters[col] = selected

            # تطبيق الفلاتر
            filtered_df = df.copy()
            for col, filt in filters.items():
                if col in date_columns:
                    start, end = filt
                    filtered_df = filtered_df[(filtered_df[col] >= start) & (filtered_df[col] <= end)]
                else:
                    filtered_df = filtered_df[filtered_df[col].isin(filt)]

            st.markdown("### 📈 Filtered Data")
            st.dataframe(filtered_df, use_container_width=True)

            # --- رسم بياني تفاعلي ---
            st.markdown("### 📊 Interactive Chart")
            numeric_cols = filtered_df.select_dtypes(include='number').columns.tolist()
            categorical_cols = filtered_df.select_dtypes(exclude='number').columns.tolist()

            if len(numeric_cols) > 0 and len(categorical_cols) > 0:
                x_col = st.selectbox("X-Axis (Categories)", categorical_cols)
                y_col = st.selectbox("Y-Axis (Values)", numeric_cols)
                fig = px.bar(filtered_df, x=x_col, y=y_col, title=f"{y_col} by {x_col}")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Not enough columns to generate a chart.")

            # --- تحميل البيانات ---
            st.markdown("### 💾 Download Filtered Data")

            # 1. تنزيل كـ Excel
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                filtered_df.to_excel(writer, index=False, sheet_name='Filtered Data')
            excel_data = excel_buffer.getvalue()

            st.download_button(
                label="📥 Download as Excel",
                data=excel_data,
                file_name="Filtered_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # 2. تنزيل كـ PDF
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.set_fill_color(255, 215, 0)  # ذهبي فاتح
            pdf.set_text_color(0, 31, 63)  # أزرق داكن

            # عنوان
            pdf.cell(0, 10, "Filtered Data Report", ln=True, align='C')
            pdf.ln(5)

            # جدول
            headers = filtered_df.columns.tolist()
            rows = filtered_df.values.tolist()

            # رأس الجدول
            col_width = 190 / max(len(headers), 1)
            for h in headers:
                pdf.cell(col_width, 10, str(h), border=1, fill=True)
            pdf.ln(10)

            # الصفوف
            pdf.set_font("Arial", size=10)
            for row in rows[:100]:  # فقط أول 100 صف لتجنب التوقف
                for item in row:
                    pdf.cell(col_width, 10, str(item), border=1)
                pdf.ln(10)

            if len(rows) > 100:
                pdf.cell(0, 10, f"... and {len(rows) - 100} more rows", ln=True)

            pdf_data = pdf.output(dest='S').encode('latin1')

            st.download_button(
                label="📥 Download as PDF",
                data=pdf_data,
                file_name="Filtered_Data_Report.pdf",
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
    - اختر شيت.
    - استخدم الفلاتر في الشريط الجانبي (التاريخ، الموظفين، المديرين...).
    - شاهد البيانات والرسم البياني.
    - حمل البيانات كـ Excel أو PDF.

    ---

    🙋‍♂️ لأي استفسار: <a href="https://wa.me/201554694554  " target="_blank">01554694554 (واتساب)</a>
    """, unsafe_allow_html=True)
