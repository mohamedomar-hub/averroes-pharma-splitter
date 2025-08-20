import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
import re
import os

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

# ------------------ ستايل مخصص (محسّن بالكامل) ------------------
custom_css = """
    <style>
    /* تطبيق الخط والخلفية */
    .stApp {
        background-color: #001f3f;
        color: white;
        font-family: 'Cairo', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }

    /* تأثير دخول تدريجي */
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(10px); }
        to { opacity: 1; transform: translateY(0); }
    }
    .fade-in {
        animation: fadeIn 1.5s ease-in;
    }

    /* تحسين المحتوى العام */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        padding-left: 4rem;
        padding-right: 4rem;
    }

    /* عناوين التحديد والرفع */
    label, .stSelectbox label, .stFileUploader label {
        color: #FFD700 !important;
        font-size: 18px !important;
        font-weight: bold !important;
    }

    /* أزرار احترافية */
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
        transform: scale(1.08) !important;
        box-shadow: 0 6px 12px rgba(0,0,0,0.4) !important;
    }

    /* مربع المعلومات (الاسم ورقم الواتساب) */
    .info-box {
        text-align: center;
        font-size: 18px;
        color: #FFD700;
        margin-top: 10px;
        line-height: 1.8;
    }
    .info-box a {
        color: #FFD700;
        text-decoration: none;
    }

    /* حاوية اللوجو في المنتصف */
    .logo-container {
        text-align: center;
        margin: 25px 0;
    }
    .logo-container img {
        max-width: 220px;
        max-height: 160px;
        border-radius: 14px;
        box-shadow: 0 6px 12px rgba(0,0,0,0.3);
        object-fit: contain;
    }

    /* فواصل أنيقة بين الأقسام */
    hr.divider {
        border: 1px solid #FFD700;
        opacity: 0.6;
        margin: 30px 0;
        border-radius: 1px;
    }
    hr.divider-dashed {
        border: 1px dashed #FFD700;
        opacity: 0.7;
        margin: 25px 0;
    }

    /* تحسين مظهر الجداول */
    .stDataFrame {
        box-shadow: 0 4px 10px rgba(0,0,0,0.2);
        border-radius: 12px;
        overflow: hidden;
        margin: 10px 0;
    }

    /* تحسين زر الرفع */
    .stFileUploader {
        border: 2px dashed #FFD700;
        border-radius: 10px;
        padding: 15px;
        background-color: rgba(255, 215, 0, 0.1);
    }
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ------------------ بدء التأثير التدريجي ------------------
st.markdown('<div class="fade-in">', unsafe_allow_html=True)

# ------------------ عرض اللوجو في المنتصف ------------------
logo_path = "logo.png"

if os.path.exists(logo_path):
    try:
        st.markdown('<div class="logo-container">', unsafe_allow_html=True)
        st.image(logo_path, width=200)
        st.markdown('</div>', unsafe_allow_html=True)
    except Exception as e:
        st.warning("⚠️ تعذر تحميل اللوجو.")
else:
    st.warning("⚠️ لم يتم العثور على ملف اللوجو 'logo.png'. تأكد من وجوده في نفس مجلد الكود.")

# ------------------ معلومات المطور (تحت اللوجو) ------------------
st.markdown(
    """
    <div class="info-box">
        <strong>Mohamed Abd ELGhany</strong><br>
        💬 
        <a href="https://wa.me/201554694554" target="_blank">
            01554694554 (WhatsApp)
        </a><br>
        📍 Head Office - 5 Settelment
    </div>
    """,
    unsafe_allow_html=True
)

# ------------------ العنوان ------------------
st.markdown("<h1 style='text-align:center; color:#FFD700;'>💊 Averroes Pharma File Splitter</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align:center; color:white;'>✂ Split & Merge Excel Files Fast & Easily</h3>", unsafe_allow_html=True)

# ------------------ رفع الملفات ------------------
uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"], accept_multiple_files=False)

# ✅ الـ if والـ else لازم يكونوا متتاليين بدون انقطاع
if uploaded_file:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        st.success(f"✅ File uploaded successfully. Sheets found: {len(excel_file.sheet_names)}")

        selected_sheet = st.selectbox("📑 Select Sheet", excel_file.sheet_names)

        if selected_sheet:
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
            df = df.fillna(method="ffill", axis=0).fillna(method="ffill", axis=1)

            st.markdown(f"### 📊 Data View – {selected_sheet}")
            st.dataframe(df, use_container_width=True)

            st.markdown("### ✂ Select the column to split by")
            col_to_split = st.selectbox(
                "Split by Column",
                df.columns,
                help="اختر العمود اللي هتقسّم عليه، مثل 'الفرع' أو 'المنطقة'"
            )

            if st.button("🚀 Start Split"):
                with st.spinner("Splitting files..."):
                    zip_buffer = BytesIO()
                    with ZipFile(zip_buffer, "w") as zip_file:
                        for value in df[col_to_split].dropna().unique():
                            sub_df = df[df[col_to_split] == value]
                            row_count = len(sub_df)
                            st.write(f"📁 **{value}**: {row_count} rows")

                            file_buffer = BytesIO()
                            with pd.ExcelWriter(file_buffer, engine="openpyxl") as writer:
                                sub_df.to_excel(writer, index=False, sheet_name=str(value)[:30])
                            file_buffer.seek(0)
                            safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1F]', '_', str(value))
                            zip_file.writestr(f"{safe_name}.xlsx", file_buffer.read())

                    zip_buffer.seek(0)

                    if zip_buffer.getvalue():
                        st.success("✅ Files split successfully!")
                        st.download_button(
                            label="📥 Download Split Files (ZIP)",
                            data=zip_buffer.getvalue(),
                            file_name=f"Split_{selected_sheet}.zip",
                            mime="application/zip"
                        )
                    else:
                        st.error("❌ Failed to generate zip file.")

        # -----------------------------------------------
        # ✅ فاصل أنيق قبل قسم الدمج
        # -----------------------------------------------
        st.markdown("<hr class='divider-dashed'>", unsafe_allow_html=True)

        # -----------------------------------------------
        # ✅ دمج ملفات Excel متعددة (ليس صفحات)
        # -----------------------------------------------
        st.markdown("### 🔄 Merge Multiple Excel Files into One")
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
                        file_name="Merged_Excel_Files.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        # -----------------------------------------------
        # ✅ فاصل قبل التحميل
        # -----------------------------------------------
        st.markdown("<hr class='divider'>", unsafe_allow_html=True)

        # -----------------------------------------------
        # ✅ تحميل كل الشيتات كما هي (مصفاة فقط)
        # -----------------------------------------------
        st.markdown("### 📥 Download Full Cleaned File (All Sheets)")
        all_sheets_output = BytesIO()
        with pd.ExcelWriter(all_sheets_output, engine="openpyxl") as writer:
            for sheet_name in excel_file.sheet_names:
                df_sheet = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                df_sheet = df_sheet.fillna(method="ffill", axis=0).fillna(method="ffill", axis=1)
                df_sheet.to_excel(writer, index=False, sheet_name=sheet_name)
        all_sheets_output.seek(0)

        st.download_button(
            label="⬇️ Download All Sheets (Cleaned)",
            data=all_sheets_output.getvalue(),
            file_name="All_Sheets_Cleaned.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error while processing the file: {e}")

    # ✅ إغلاق الـ fade-in هنا، بعد انتهاء كود if
    st.markdown('</div>', unsafe_allow_html=True)

else:
    # ✅ الـ else يجب أن يكون مباشرة بعد الـ try-except داخل if
    st.markdown('<p style="text-align:center; color:#FFD700;">⚠️ No file uploaded yet.</p></div>', unsafe_allow_html=True)
