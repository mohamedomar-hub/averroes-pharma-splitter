import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
import re
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# ------------------ ربط خط عربي ------------------
st.markdown('<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap" rel="stylesheet">', unsafe_allow_html=True)

# ------------------ إعدادات الصفحة ------------------
st.set_page_config(page_title="Averroes Pharma Splitter", page_icon="💊", layout="wide", initial_sidebar_state="collapsed")

# ------------------ إخفاء عناصر Streamlit ------------------
st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
.stApp { background-color: #001f3f; color: white; font-family: 'Cairo', sans-serif; }
</style>
""", unsafe_allow_html=True)

# ------------------ عرض اللوجو ------------------
logo_path = "logo.png"
if os.path.exists(logo_path):
    st.image(logo_path, width=200, use_column_width="center")

# ------------------ معلومات المطور ------------------
st.markdown("""
<div style="text-align:center; font-size:18px; color:#FFD700;">
    By <strong>Mohamed Abd ELGhany</strong> – 
    <a href="https://wa.me/201554694554" target="_blank" style="color:#FFD700;">01554694554 (WhatsApp)</a>
</div>
""", unsafe_allow_html=True)

# ------------------ العنوان ------------------
st.markdown("<h1 style='text-align:center; color:#FFD700;'>💊 Averroes Pharma File Splitter</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align:center; color:white;'>✂ Format your files accurately and without losing formatting.</h3>", unsafe_allow_html=True)

# ------------------ رفع الملف ------------------
uploaded_file = st.file_uploader("📂 Upload file Excel", type=["xlsx"])

if uploaded_file:
    try:
        # قراءة الملف
        input_bytes = uploaded_file.getvalue()
        original_wb = load_workbook(filename=BytesIO(input_bytes), data_only=False)
        sheet_names = original_wb.sheetnames

        selected_sheet = st.selectbox("Select Sheet", sheet_names)

        # قراءة البيانات كـ DataFrame
        df = pd.read_excel(BytesIO(input_bytes), sheet_name=selected_sheet)
        st.dataframe(df, use_container_width=True)

        # اختيار العمود
        col_to_split = st.selectbox("Select the column you will split.", df.columns)

        if st.button("🚀 Start Spilit"):
            with st.spinner("جاري التقسيم بدقة..."):

                # --- تنظيف الأسماء ---
                def clean_name(name):
                    name = str(name).strip()
                    return re.sub(r'[\\/*?:\[\]|<>"]', '_', name)[:30] or "Sheet"

                base_filename = clean_name(uploaded_file.name.rsplit('.', 1)[0])
                zip_buffer = BytesIO()
                with ZipFile(zip_buffer, "w") as zip_file:
                    ws = original_wb[selected_sheet]
                    col_index = df.columns.get_loc(col_to_split)  # 0-based

                    # --- جمع القيم الفريدة من العمود المختار ---
                    unique_values = df[col_to_split].dropna().unique()

                    for value in unique_values:
                        # --- إنشاء ملف جديد ---
                        new_wb = load_workbook(filename=BytesIO(input_bytes))
                        new_ws = new_wb.active
                        new_ws.title = clean_name(value)

                        # --- نسخ الرأس (الصف الأول) ---
                        for cell in ws[1]:
                            dst_cell = new_ws.cell(1, cell.column, cell.value)
                            if cell.has_style:
                                if cell.font:
                                    dst_cell.font = Font(
                                        name=cell.font.name, size=cell.font.size,
                                        bold=cell.font.bold, italic=cell.font.italic,
                                        color=cell.font.color
                                    )
                                if cell.fill and cell.fill.fill_type:
                                    dst_cell.fill = cell.fill
                                if cell.border:
                                    dst_cell.border = cell.border
                                if cell.alignment:
                                    dst_cell.alignment = cell.alignment
                                dst_cell.number_format = cell.number_format

                        # --- نسخ الصفوف اللي فيها القيمة المطلوبة ---
                        row_idx_new = 2
                        for row in ws.iter_rows(min_row=2):
                            cell_in_col = row[col_index]  # العمود المختار (0-based)
                            if cell_in_col.value == value:
                                for src_cell in row:
                                    dst_cell = new_ws.cell(row_idx_new, src_cell.column, src_cell.value)
                                    if src_cell.has_style:
                                        if src_cell.font:
                                            dst_cell.font = Font(
                                                name=src_cell.font.name, size=src_cell.font.size,
                                                bold=src_cell.font.bold, italic=src_cell.font.italic,
                                                color=src_cell.font.color
                                            )
                                        if src_cell.fill and src_cell.fill.fill_type:
                                            dst_cell.fill = src_cell.fill
                                        if src_cell.border:
                                            dst_cell.border = src_cell.border
                                        if src_cell.alignment:
                                            dst_cell.alignment = src_cell.alignment
                                        dst_cell.number_format = src_cell.number_format
                                row_idx_new += 1

                        # --- نسخ عرض الأعمدة ---
                        for col_letter in ws.column_dimensions:
                            new_ws.column_dimensions[col_letter].width = ws.column_dimensions[col_letter].width

                        # --- حفظ الملف ---
                        file_buffer = BytesIO()
                        new_wb.save(file_buffer)
                        file_buffer.seek(0)
                        file_name = f"{base_filename}_{clean_name(value)}.xlsx"
                        zip_file.writestr(file_name, file_buffer.read())
                        st.write(f"✅ تم إنشاء ملف لـ: **{value}**")

                zip_buffer.seek(0)
                st.success("🎉 Done Spilit successfully.!")
                st.download_button(
                    label="📥 Download files spilit (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"{base_filename}_Split.zip",
                    mime="application/zip"
                )

    except Exception as e:
        st.error(f"❌ خطأ: {str(e)}")
else:
    st.info("📂 Upload file excel to spilit.")
