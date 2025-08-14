import streamlit as st
import pandas as pd
from io import BytesIO

# ------------------ إعدادات الصفحة ------------------
st.set_page_config(page_title="Averroes Pharma Splitter", page_icon="💊", layout="centered")

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
    body {
        background-color: #f4f6f9;
        font-family: 'Cairo', sans-serif;
    }
    .title {
        text-align: center;
        color: #2e86de;
        font-size: 38px;
        font-weight: bold;
        margin-bottom: 5px;
    }
    .subtitle {
        text-align: center;
        color: #555;
        font-size: 18px;
        margin-bottom: 30px;
    }
    .stButton>button {
        background-color: #2e86de;
        color: white;
        border-radius: 10px;
        padding: 10px 20px;
        font-size: 18px;
        border: none;
        cursor: pointer;
    }
    .stButton>button:hover {
        background-color: #1b4f72;
    }
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ------------------ اللوجو والعنوان ------------------
st.image("logo.png", width=170)
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
