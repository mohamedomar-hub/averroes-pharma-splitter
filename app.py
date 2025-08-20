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
    initial_sidebar_state="expanded"  # Sidebar مفتوح من البداية
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
    /* تطبيق الخط والخلفية */
    .stApp {
        background-color: #001f3f;
        color: white;
        font-family: 'Cairo', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
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

    /* مربع المعلومات في الـ Sidebar */
    .info-box {
        font-size: 17px;
        color: #FFD700;
        line-height: 1.8;
        text-align: center;
    }
    .info-box a {
        color: #FFD700;
        text-decoration: none;
    }

    /* فواصل أنيقة */
    hr.divider {
        border: 1px solid #FFD700;
        opacity: 0.6;
        margin: 20px 0;
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

    /* عنوان الشرح */
    .guide-title {
        color: #FFD700;
        font-weight: bold;
        font-size: 20px;
    }
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ------------------ Sidebar (الشريط الجانبي) ------------------
with st.sidebar:
    # لوجو في الأعلى
    logo_path = "logo.png"
    if os.path.exists(logo_path):
        try:
            st.image(logo_path, width=140)
        except:
            pass

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)

    # --- بيانات المطور ---
    st.markdown("<p class='info-box'>", unsafe_allow_html=True)
    st.markdown("**Created by**")
    st.markdown("""
    <strong>Mohamed Abd ELGhany</strong><br>
    💬 <a href="https://wa.me/201554694554" target="_blank">01554694554 (WhatsApp)</a><br>
    📍 Head Office - 5 Settelment
    """, unsafe_allow_html=True)
    st.markdown("</p>", unsafe_allow_html=True)

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)

    # --- شرح الاستخدام ---
    with st.expander("ℹ️ Info - طريقة الاستخدام"):
        st.markdown("""
        <div class='guide-title'>🎯 مرحبًا بك في أداة Averroes Pharma!</div>
        
        هذه الأداة تساعدك على **تقسيم ودمج ملفات الإكسل بسرعة ودقة** بدون برامج إضافية.

        ---

        ### 🔧 أولًا: تقسيم ملف Excel
        1. **ارفع ملف الإكسل** من زر "Upload Excel File".
        2. اختر **الشيت اللي عاوزه** من القائمة.
        3. اختر **العمود اللي عاوز تقسم عليه** (مثل: "الفرع"، "المنطقة"، "المندوب").
        4. اضغط على **🚀 Start Split**.
        5. هتلاقي زر لتحميل ملف ZIP يحتوي على كل جزء منفصل.

        ✅ مثال: لو قسمت على "الفرع"، هيكون عندك: `القاهرة.xlsx`, `الإسكندرية.xlsx`, إلخ.

        ---

        ### 🔗 ثانيًا: دمج ملفات Excel (منفصلة)
        1. في الأسفل، اضغط على **"Upload Excel Files to Merge"**.
        2. ارفع **أكثر من ملف Excel** (مثلاً: `يناير.xlsx`, `فبراير.xlsx`).
        3. اضغط على **✨ Merge Selected Files**.
        4. هتلاقي زر لتحميل ملف واحد يحتوي على كل البيانات.

        ✅ ملاحظة: كل صف هيكون فيه عمود "Source File" يوضح منين جاي.

        ---

        ### 💾 ثالثًا: تحميل الملفات
        - **📥 Download Split Files (ZIP)**: الملفات المقسمة.
        - **📥 Download Merged File**: الملفات المدموجة.
        - **⬇️ Download All Sheets (Cleaned)**: نفس الملف اللي رفعته، بس تم تنظيف الخلايا الفارغة.

        ---

        ### ❓ أسئلة شائعة
        - **هل يتم تعديل البيانات؟**  
          لا، فقط يتم "ملء" الخلايا الفارغة بالقيمة السابقة (لتحسين العرض).
        - **هل يدعم CSV؟**  
          لا، حاليًا يدعم فقط `.xlsx`.
        - **هل البيانات تُحفظ على سيرفر؟**  
          لا، كل شيء يتم على جهازك، وما يُرفع يُمسح بعد التحديث.

        ---

        🙋‍♂️ لو واجهتك أي مشكلة، ابعتلي علي الواتساب.
        """, unsafe_allow_html=True)

# ------------------ العنوان الرئيسي ------------------
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
else:
    st.markdown("<p style='text-align:center; color:#FFD700;'>⚠️ No file uploaded yet.</p>", unsafe_allow_html=True)
