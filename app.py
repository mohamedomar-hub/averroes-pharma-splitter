import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
import re
import os
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl import load_workbook

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
st.markdown("<h3 style='text-align:center; color:white;'>✂ Split & Merge Excel Files with Full Formatting</h3>", unsafe_allow_html=True)

# ------------------ رفع الملفات ------------------
uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"], accept_multiple_files=False)

if uploaded_file:
    try:
        input_bytes = uploaded_file.getvalue()
        original_wb = load_workbook(filename=BytesIO(input_bytes), data_only=False)

        st.success(f"✅ File uploaded successfully. Sheets found: {len(original_wb.sheetnames)}")

        selected_sheet = st.selectbox("📑 Select Sheet", original_wb.sheetnames)

        if selected_sheet:
            df = pd.read_excel(BytesIO(input_bytes), sheet_name=selected_sheet)
            df = df.fillna(method="ffill", axis=0).fillna(method="ffill", axis=1)

            st.markdown(f"### 📊 Data View – {selected_sheet}")
            st.dataframe(df, use_container_width=True)

            st.markdown("### ✂ Select the column to split by")
            col_to_split = st.selectbox(
                "Split by Column",
                df.columns,
                help="اختر العمود اللي هتقسّم عليه، مثل 'الفرع' أو 'المنطقة'"
            )

            if st.button("🚀 Start Split with Original Format"):
                with st.spinner("Splitting files while preserving full formatting and blank rows..."):
                    zip_buffer = BytesIO()
                    with ZipFile(zip_buffer, "w") as zip_file:
                        original_ws = original_wb[selected_sheet]
                        col_idx = df.columns.get_loc(col_to_split)

                        # جمع القيم المختلفة من العمود المختار
                        values = set()
                        for row in original_ws.iter_rows(min_row=2, max_row=original_ws.max_row):
                            cell = row[col_idx - 1]  # -1 لأن الصف الأول هو الرأس
                            if cell.value is not None:
                                values.add(cell.value)

                        # دالة تنظيف اسم الشيت (تجنب الرموز الممنوعة)
                        def clean_sheet_name(name):
                            name = str(name).strip()
                            invalid_chars = r'[\\/*?:\[\]|<>]'
                            cleaned = re.sub(invalid_chars, '-', name)
                            if not cleaned or cleaned in ['.', '..']:
                                cleaned = "Sheet"
                            return cleaned[:30]

                        # دالة تنظيف اسم الملف
                        def clean_filename(name):
                            name = str(name).strip()
                            invalid_chars = r'[\\/*?:\[\]|<>]'
                            cleaned = re.sub(invalid_chars, '_', name)
                            return cleaned[:250]

                        # اسم الملف الأصلي (بدون расширение)
                        base_filename = clean_filename(uploaded_file.name.split('.')[0])

                        # إنشاء ملف لكل قيمة
                        for value in values:
                            output_buffer = BytesIO()
                            new_wb = load_workbook(filename=BytesIO(input_bytes))
                            new_ws = new_wb.active
                            new_ws.title = clean_sheet_name(value)

                            # نسخ الصف الأول (الرأس)
                            for col_letter in original_ws[1]:
                                src_cell = col_letter
                                dst_cell = new_ws.cell(1, col_letter.column)
                                dst_cell.value = src_cell.value
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
                                        border = src_cell.border
                                        dst_cell.border = Border(
                                            left=border.left,
                                            right=border.right,
                                            top=border.top,
                                            bottom=border.bottom,
                                            diagonal=border.diagonal,
                                            diagonal_direction=border.diagonal_direction,
                                            outline=border.outline,
                                            vertical=border.vertical,
                                            horizontal=border.horizontal
                                        )
                                    if src_cell.alignment:
                                        dst_cell.alignment = Alignment(
                                            horizontal=src_cell.alignment.horizontal,
                                            vertical=src_cell.alignment.vertical,
                                            text_rotation=src_cell.alignment.text_rotation,
                                            wrap_text=src_cell.alignment.wrap_text,
                                            shrink_to_fit=src_cell.alignment.shrink_to_fit,
                                            indent=src_cell.alignment.indent
                                        )
                                    dst_cell.number_format = src_cell.number_format

                            # نسخ الصفوف (بما في ذلك الفراغات)
                            for row_idx, row in enumerate(original_ws.iter_rows(min_row=2), 2):
                                src_cell = row[col_idx - 1]
                                if src_cell.value == value:
                                    for col_idx, cell in enumerate(row, 1):
                                        dst_cell = new_ws.cell(row=row_idx, column=col_idx)
                                        dst_cell.value = cell.value
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
                                                border = cell.border
                                                dst_cell.border = Border(
                                                    left=border.left,
                                                    right=border.right,
                                                    top=border.top,
                                                    bottom=border.bottom,
                                                    diagonal=border.diagonal,
                                                    diagonal_direction=border.diagonal_direction,
                                                    outline=border.outline,
                                                    vertical=border.vertical,
                                                    horizontal=border.horizontal
                                                )
                                            if cell.alignment:
                                                dst_cell.alignment = Alignment(
                                                    horizontal=cell.alignment.horizontal,
                                                    vertical=cell.alignment.vertical,
                                                    text_rotation=cell.alignment.text_rotation,
                                                    wrap_text=cell.alignment.wrap_text,
                                                    shrink_to_fit=cell.alignment.shrink_to_fit,
                                                    indent=cell.alignment.indent
                                                )
                                            dst_cell.number_format = cell.number_format

                            # نسخ عرض الأعمدة
                            for col_letter in original_ws.column_dimensions:
                                new_ws.column_dimensions[col_letter].width = original_ws.column_dimensions[col_letter].width

                            # حفظ الملف
                            new_wb.save(output_buffer)
                            output_buffer.seek(0)

                            # اسم الملف الناتج: اسم الملف الأصلي + القيمة
                            clean_value = clean_filename(str(value))
                            file_name = f"{base_filename}_{clean_value}.xlsx"

                            zip_file.writestr(file_name, output_buffer.read())

                    zip_buffer.seek(0)
                    st.success("✅ Files split successfully with original formatting!")
                    st.download_button(
                        label="📥 Download Split Files (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name=f"Split_{base_filename}.zip",
                        mime="application/zip"
                    )

        # -----------------------------------------------
        # 🔄 دمج ملفات Excel
        # -----------------------------------------------
        st.markdown("<hr class='divider-dashed'>", unsafe_allow_html=True)
        st.markdown("### 🔄 Merge Multiple Excel Files (Preserve Data)")
        merge_files = st.file_uploader("📤 Upload Excel Files to Merge", type=["xlsx"], accept_multiple_files=True)

        if merge_files:
            if st.button("✨ Merge Selected Files"):
                with st.spinner("Merging Excel files..."):
                    combined_df = pd.DataFrame()
                    for file in merge_files:
                        df_temp = pd.read_excel(file)
                        df_temp["Source File"] = file.name
                        combined_df = pd.concat([combined_df, df_temp], ignore_index=True)

                    combined_buffer = BytesIO()
                    with pd.ExcelWriter(combined_buffer, engine="openpyxl") as writer:
                        combined_df.to_excel(writer, index=False, sheet_name="Consolidated")
                    combined_buffer.seek(0)

                    st.success("✅ Files merged successfully!")
                    st.download_button(
                        label="📥 Download Merged File",
                        data=combined_buffer.getvalue(),
                        file_name="Merged_Consolidated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        # -----------------------------------------------
        # 💾 تحميل كل الشيتات نظيفة
        # -----------------------------------------------
        st.markdown("<hr class='divider'>", unsafe_allow_html=True)
        st.markdown("### 📥 Download Full Cleaned File (All Sheets, Original Format)")
        cleaned_buffer = BytesIO()
        with ZipFile(cleaned_buffer, "w") as zip_out:
            for sheet_name in original_wb.sheetnames:
                df_sheet = pd.read_excel(BytesIO(input_bytes), sheet_name=sheet_name)
                df_sheet = df_sheet.fillna(method="ffill", axis=0).fillna(method="ffill", axis=1)

                temp_buffer = BytesIO()
                with pd.ExcelWriter(temp_buffer, engine="openpyxl") as writer:
                    df_sheet.to_excel(writer, index=False, sheet_name=sheet_name)
                    wb_temp = writer.book
                    ws_temp = writer.sheets[sheet_name]
                    orig_ws = original_wb[sheet_name]

                    for col_letter in orig_ws.column_dimensions:
                        ws_temp.column_dimensions[col_letter].width = orig_ws.column_dimensions[col_letter].width

                temp_buffer.seek(0)
                zip_out.writestr(f"Cleaned_{sheet_name}.xlsx", temp_buffer.read())

        cleaned_buffer.seek(0)
        st.download_button(
            label="⬇️ Download All Cleaned Sheets (ZIP)",
            data=cleaned_buffer.getvalue(),
            file_name="All_Cleaned_Sheets.zip",
            mime="application/zip"
        )

    except Exception as e:
        st.error(f"❌ Error while processing the file: {e}")
else:
    st.markdown("<p style='text-align:center; color:#FFD700;'>⚠️ No file uploaded yet.</p>", unsafe_allow_html=True)

# ------------------ قسم Info ------------------
st.markdown("<hr class='divider' id='info-section'>", unsafe_allow_html=True)
with st.expander("📖 طريقة الاستخدام - اضغط لعرض التعليمات"):
    st.markdown("""
    <div class='guide-title'>🎯 مرحبًا بك في أداة Averroes Pharma!</div>
    هذه الأداة تقسم ودمج ملفات الإكسل <strong>مع الحفاظ الكامل على التنسيق الأصلي</strong> (الألوان، الخطوط، الأحجام، والحدود).

    ---

    ### 🔧 أولًا: التقسيم مع الحفاظ على التنسيق
    1. ارفع ملف Excel.
    2. اختر الشيت.
    3. اختر عمود التقسيم (مثل: "الفرع").
    4. اضغط على **"Start Split with Original Format"**.
    5. هيتم إنشاء ملفات منفصلة لكل قيمة، <strong>بنفس التنسيق، الألوان، وعرض الأعمدة</strong>.

    ✅ الناتج: كل ملف يشبه تمامًا الجزء الأصلي من الشيت.

    ---

    ### 🔗 ثانيًا: الدمج
    - ارفع أكثر من ملف.
    - اضغط "Merge" لتحصل على ملف واحد.

    ---

    ### 💾 تنظيف وتحميل
    - خيار لتحميل كل الشيتات بعد التنظيف (بدون فقدان التنسيق).

    ---

    🙋‍♂️ لأي استفسار: <a href="https://wa.me/201554694554" target="_blank">01554694554 (واتساب)</a>
    """, unsafe_allow_html=True)
