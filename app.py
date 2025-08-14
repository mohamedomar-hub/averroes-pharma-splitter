import streamlit as st
import pandas as pd
from io import BytesIO

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
        max-height: 100px;
    }
    .admin-text {
        font-size: 20px;
        font-weight: bold;
        color: #FFD700; /* ذهبي */
    }
    .title {
        text-align: center;
        color: #FFD700;
        font-size: 40px;
        font-weight: bold;
        margin-bottom: 5px;
    }
    .subtitle {
        text-align: center;
        color: white;
        font-size: 20px;
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
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ------------------ الهيدر ------------------
st.markdown(
    f"""
    <div class="header-container">
        <img src="logo.png" class="logo">
        <div class="admin-text">
            By Admin Mohamed Abd ELGhany – 01554694554
        </div>
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
    df = pd.read_excel(uploaded_file)
    st.success("✅ تم تحميل الملف بنجاح!")
    st.dataframe(df)

    col = st.selectbox("📌 اختر العمود للتقسيم", df.columns)

    if st.button("🚀 تقسيم الملف"):
        for value, group in df.groupby(col):
            output = BytesIO()
            group.to_excel(output, index=False)
            st.download_button(
                label=f"⬇ تحميل {value}.xlsx",
                data=output.getvalue(),
                file_name=f"{value}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
