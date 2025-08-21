import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
import re
import os
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl import load_workbook, Workbook

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

# ------------------ عرض اللوجو ------------------
logo_path = "logo.png"
if os.path.exists(logo_path):
    st.markdown('<div style="text-align:center; margin:20px 0;">', unsafe_allow_html=True)
    st.image(logo_path, width=200)
    st.markdown('</div>', unsafe_allow_html=True)
else:
    st.warning("⚠️ لم يتم العثور على ملف اللوجو 'logo.png'.")

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
st.markdown("<h3 style='text-align:center; color:white;'>✂ DivisionDivide your files with ease and accuracy.</h3>", unsafe_allow_html=True)

# ------------------ رفع الملف ------------------
uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"], accept_multiple_files=False)

if uploaded_file:
    try:
        input_bytes = uploaded_file.getvalue()
        original_wb = load_workbook(filename=BytesIO(input_bytes), data_only=False)
        st.success(f"✅The files have been uploaded successfully. Count Sheets: {len(original_wb.sheetnames)}")

        selected_sheet = st.selectbox("Select Sheet", original_wb.sheetnames)

        if selected_sheet:
            df = pd.read_excel(BytesIO(input_bytes), sheet_name=selected_sheet)
            st.markdown(f"### 📊 Data View – {selected_sheet}")
            st.dataframe(df, use_container_width=True)

            st.markdown("### ✂ Select Coulmn to spilit it")
            col_to_split = st.selectbox(
                "Split by Column",
                df.columns,
                help="Select Coulmn to spilit it, Like 'Brick' Or 'Area Manager'"
            )

            # --- زر التقسيم ---
            if st.button("🚀 Start Spilit"):
                with st.spinner("The splitting process is ongoing while the original format is preserved...."):

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

                            # --- نسخ الصفوف اللي فيها القيمة ---
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
                            st.write(f"📁Complete create file: `{value}`")

                    zip_buffer.seek(0)
                    st.success("🎉 The division was successful.!")
                    st.download_button(
                        label="📥 Upload Division files (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name=f"Split_{clean_name(uploaded_file.name.rsplit('.',1)[0])}.zip",
                        mime="application/zip"
                    )

        # -----------------------------------------------
        # 🔄 دمج ملفات Excel
        # -----------------------------------------------
        st.markdown("<hr class='divider-dashed'>", unsafe_allow_html=True)
        st.markdown("### 🔄 Merge files Excel Multiple")
        merge_files = st.file_uploader("📤 ارفع ملفات Excel للدمج", type=["xlsx"], accept_multiple_files=True)

        if merge_files:
            if st.button("✨ Merge files"):
                with st.spinner("Merging is currently in progress...."):
                    combined_df = pd.DataFrame()
                    for file in merge_files:
                        df_temp = pd.read_excel(file)
                        df_temp["Source File"] = file.name
                        combined_df = pd.concat([combined_df, df_temp], ignore_index=True)

                    combined_buffer = BytesIO()
                    with pd.ExcelWriter(combined_buffer, engine="openpyxl") as writer:
                        combined_df.to_excel(writer, index=False, sheet_name="Consolidated")
                    combined_buffer.seek(0)

                    st.success("✅ Done Merge Successfully!")
                    st.download_button(
                        label="📥 Upload File Merge",
                        data=combined_buffer.getvalue(),
                        file_name="Merged_Consolidated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    except Exception as e:
        st.error(f"❌ خطأ في معالجة الملف: {e}")
else:
    st.markdown("<p style='text-align:center; color:#FFD700;'>⚠️ لم يتم رفع ملف بعد.</p>", unsafe_allow_html=True)

# ------------------ قسم Info ------------------
st.markdown("<hr class='divider' id='info-section'>", unsafe_allow_html=True)
with st.expander("📖 How to use - Click to view instructions "):
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
    - هتلاقي زر لتحميل ملف واحد فيه كل البيانات.

    ---

    🙋‍♂️ لأي استفسار: <a href="https://wa.me/201554694554" target="_blank">01554694554 (واتساب)</a>
    """, unsafe_allow_html=True)






