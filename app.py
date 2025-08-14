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

    /* تعديل شكل نص رفع الملف (label) */
    .stFileUploader label {
        color: white !important;
        font-size: 20px !important;
        font-weight: bold !important;
        text-align: center;
    }

    /* تعديل زر رفع الملف (Browse files) ليكون أكثر وضوحًا */
    .stFileUploader div div button {
        background-color: #FFD700 !important;
        color: black !important;           /* لون نص واضح */
        font-weight: bold !important;
        font-size: 18px !important;
        border-radius: 8px !important;
        padding: 10px 20px !important;
        border: 2px solid #FFA500 !important; /* حدود ذهبية فاتحة لزيادة الوضوح */
        box-shadow: 0 2px 4px rgba(0,0,0,0.2) !important;
        transition: all 0.3s ease !important;
    }

    /* تأثير عند التحويم على زر الرفع */
    .stFileUploader div div button:hover {
        background-color: #FFC107 !important;
        color: #1a1a1a !important;
        transform: scale(1.05);
        border-color: #FF8C00 !important;
    }

    /* تأكيد وضوح النص داخل الزر حتى في الحالات النشطة */
    .stFileUploader div div button:active {
        background-color: #FFB300 !important;
        color: #000 !important;
    }
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ------------------ عرض اللوجو من فولدر المشروع ------------------
def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

logo_path = "logo.png"  # تأكد أن الصورة موجودة في نفس الفولدر
try:
    logo_base64 = get_base64_of_bin_file(logo_path)
except FileNotFoundError:
    st.error("❌ لم يتم العثور على ملف اللوجو 'logo.png'. تأكد من وجوده في نفس مجلد الكود.")
    logo_base64 = ""  # في حالة عدم وجود اللوجو

# عرض الهيدر فقط إذا كان اللوجو متاحًا
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
        df = pd.read_excel(uploaded_file)
        st.success("✅ تم تحميل الملف بنجاح!")
        st.dataframe(df)

        col = st.selectbox("📌 اختر العمود للتقسيم", df.columns)

        if st.button("🚀 تقسيم الملف"):
            for value, group in df.groupby(col):
                output = BytesIO()
                group.to_excel(output, index=False)
                output.seek(0)  # تأكد من إعادة المؤشر للبداية
                st.download_button(
                    label=f"⬇ تحميل {value}.xlsx",
                    data=output.getvalue(),
                    file_name=f"{value}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"❌ حدث خطأ أثناء قراءة الملف: {e}")
