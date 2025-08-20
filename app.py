import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
import re
import os

# ------------------ إعدادات الصفحة (يجب أن تكون أول سطر بعد الاستيراد) ------------------
st.set_page_config(
    page_title="Averroes Pharma Splitter",
    page_icon="💊",
    layout="wide",  # ← يضمن عرض Sidebar
    initial_sidebar_state="expanded"  # ← Sidebar مفتوح من البداية
)

# ------------------ ربط بخط عربي جميل (Cairo) ------------------
st.markdown(
    '<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap" rel="stylesheet">',
    unsafe_allow_html=True
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
    .stButton>button {
        background-color: #FFD700 !important;
        color: black !important;
        font-weight: bold !important;
        border-radius: 12px !important;
        padding: 12px 24px !important;
        box-shadow: 0 4px 8px rgba(0,0,0,0.3) !important;
        transition: all 0.3s ease !important;
    }
    .stButton>button:hover {
        background-color: #FFC107 !important;
        transform: scale(1.08);
    }
    hr.divider {
        border: 1px solid #FFD700;
        opacity: 0.6;
        margin: 20px 0;
    }
    .stDataFrame {
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 4px 10px rgba(0,0,0,0.2);
    }
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ------------------ Sidebar (يجب أن يكون بعد set_page_config مباشرة) ------------------
with st.sidebar:
    st.markdown("<h3 style='color:#FFD700;'>🔐 Averroes Pharma</h3>", unsafe_allow_html=True)

    # عرض اللوجو
    logo_path = "logo.png"
    if os.path.exists(logo_path):
        try:
            st.image(logo_path, width=140)
        except Exception as e:
            st.caption("Logo not found")

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)

    # بيانات المطور
    st.markdown("### Created by")
    st.markdown("**Mohamed Abd ELGhany**")
    st.markdown("[💬 WhatsApp: 01554694554](https://wa.me/201554694554)")
    st.markdown("📍 Head Office - 5 Settelment")

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)

    # شرح الاستخدام
    with st.expander("ℹ️ طريقة الاستخدام"):
        st.markdown("""
        <div style="font-size:16px; line-height:1.7;">
        <strong>🎯 التقسيم:</strong><br>
        1. ارفع ملف Excel.<br>
        2. اختر العمود اللي عاوز تقسم عليه.<br>
        3. اضغط <strong>Start Split</strong>.<br>
        4. حمل الملفات المقسمة من الزر.

        <br>
        <strong>🔗 الدمج:</strong><br>
        1. في قسم الدمج، ارفع أكثر من ملف.<br>
        2. اضغط <strong>Merge Selected Files</strong>.<br>
        3. حمل الملف الموحد.

        <br>
        <strong>⚠️ ملاحظات:</strong><br>
        • لا يتم حفظ بياناتك.<br>
        • كل العمليات على جهازك.<br>
        • يدعم فقط ملفات .xlsx
        </div>
        """, unsafe_allow_html=True)

# ------------------ المحتوى الرئيسي ------------------
st.markdown("<h1 style='text-align:center; color:#FFD700;'>💊 Averroes Pharma File Splitter</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align:center; color:white;'>✂ Split & Merge Excel Files Fast & Easily</h3>", unsafe_allow_html=True)
st.markdown("<hr class='divider'>", unsafe_allow_html=True)

# ------------------ رفع الملفات ------------------
uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"], accept_multiple_files=False)

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
                help="اختر العمود اللي هتقسّم عليه"
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
        # ✅ دمج ملفات Excel متعددة
        # -----------------------------------------------
        st.markdown("<hr class='divider'>", unsafe_allow_html=True)
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
        # ✅ تحميل كل الشيتات
        # -----------------------------------------------
        st.markdown("<hr class='divider'>", unsafe_allow_html=True)
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
else:
    st.markdown("<p style='text-align:center; color:#FFD700;'>⚠️ No file uploaded yet.</p>", unsafe_allow_html=True)
