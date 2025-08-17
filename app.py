import streamlit as st
import pandas as pd
from io import BytesIO
import base64

# ------------------ إعدادات الصفحة ------------------
st.set_page_config(page_title="Averroes Pharma Splitter", page_icon="💊", layout="wide")

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
        background-color: #001f3f; /* كحلي */
        color: white;
        font-size: 18px;
        font-family: 'Cairo', sans-serif;
    }
    .header-container {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 10px 20px;
    }
    .logo {
        max-height: 120px;
    }
    .admin-text {
        font-size: 22px;
        font-weight: bold;
        color: #FFD700; /* ذهبي */
    }
    .title {
        text-align: center;
        color: #FFD700;
        font-size: 42px;
        font-weight: bold;
        margin-bottom: 5px;
    }
    .subtitle {
        text-align: center;
        color: white;
        font-size: 22px;
        margin-bottom: 30px;
    }
    .stButton>button {
        background-color: #FFD700;
        color: black;
        border-radius: 10px;
        padding: 10px 20px;
        font-size: 18px;
        border: none;
        cursor: pointer;
    }
    .stButton>button:hover {
        background-color: #daa520;
    }
    .stFileUploader label {
        color: white !important;
        font-size: 20px !important;
        font-weight: bold !important;
        text-align: center;
    }
    .stFileUploader div div button {
        background-color: #FFD700 !important;
        color: black !important;
        font-weight: bold !important;
        font-size: 18px !important;
        border-radius: 8px !important;
        padding: 10px 20px !important;
        border: 2px solid #FFA500 !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2) !important;
        transition: all 0.3s ease !important;
    }
    .stFileUploader div div button:hover {
        background-color: #FFC107 !important;
        color: #1a1a1a !important;
        transform: scale(1.05);
        border-color: #FF8C00 !important;
    }
    .stFileUploader div div button:active {
        background-color: #FFB300 !important;
        color: #000 !important;
    }
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ------------------ عرض اللوجو ------------------
def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

logo_path = "logo.png"
try:
    logo_base64 = get_base64_of_bin_file(logo_path)
except FileNotFoundError:
    st.error("❌ لم يتم العثور على ملف اللوجو 'logo.png'. تأكد من وجوده في نفس مجلد الكود.")
    logo_base64 = ""

if logo_base64:
    st.markdown(
        f"""
        <div class="header-container">
            <img src="data:image/png;base64,{logo_base64}" class="logo">
            <div class="admin-text">
                By Admin Mohamed Abd ELGhany – 01554694554
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
else:
    st.markdown(
        """
        <div class="admin-text" style="text-align: center; margin-bottom: 20px;">
            By Admin Mohamed Abd ELGhany – 01554694554
        </div>
        """,
        unsafe_allow_html=True
    )

# ------------------ العنوان ------------------
st.markdown("<div class='title'>Averroes Pharma File Splitter</div>", unsafe_allow_html=True)
st.markdown("<div class='subtitle'>✂ تقسيم ملفات Excel بسهولة وسرعة</div>", unsafe_allow_html=True)

# ------------------ رفع الملف ------------------
uploaded_file = st.file_uploader("📂 ارفع ملف Excel", type=["xlsx"])

if uploaded_file:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        st.success(f"✅ تم تحميل الملف وفيه {len(excel_file.sheet_names)} شيت.")

        # لتجميع كل الشيتات في ملف واحد
        all_sheets_output = BytesIO()
        with pd.ExcelWriter(all_sheets_output, engine="openpyxl") as writer:
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

                # معالجة merge cells
                df = df.fillna(method="ffill", axis=0).fillna(method="ffill", axis=1)

                # عرض بيانات الشيت بالكامل
                with st.expander(f"📊 بيانات شيت {sheet_name}"):
                    st.dataframe(df)

                # تجهيز ملف Excel لكل شيت منفصل
                output = BytesIO()
                df.to_excel(output, index=False, sheet_name=sheet_name)
                output.seek(0)

                st.download_button(
                    label=f"⬇ تحميل {sheet_name}.xlsx",
                    data=output.getvalue(),
                    file_name=f"{sheet_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # كتابة كل الشيتات في ملف واحد
                df.to_excel(writer, index=False, sheet_name=sheet_name)

        all_sheets_output.seek(0)

        # زر تحميل الكل
        st.markdown(
            """
            <style>
            .big-download button {
                background-color: #28a745 !important;
                color: white !important;
                font-size: 20px !important;
                font-weight: bold !important;
                border-radius: 12px !important;
                padding: 15px 30px !important;
                border: 3px solid #1e7e34 !important;
                box-shadow: 0px 4px 8px rgba(0,0,0,0.3);
            }
            .big-download button:hover {
                background-color: #218838 !important;
                border-color: #18632a !important;
            }
            </style>
            """,
            unsafe_allow_html=True
        )

        st.download_button(
            label="⬇⬇ تحميل كل الشيتات مرة واحدة ⬇⬇",
            data=all_sheets_output.getvalue(),
            file_name="All_Sheets.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="all_sheets",
            help="تحميل ملف Excel يحتوي على كل الشيتات"
        )

    except Exception as e:
        st.error(f"❌ حدث خطأ أثناء قراءة الملف: {e}")
