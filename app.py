import streamlit as st
import pandas as pd
from io import BytesIO
import base64
import os

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
    label, .stSelectbox label, .stFileUploader label {
        color: #FFD700 !important;
        font-size: 18px !important;
        font-weight: bold !important;
    }
    .stButton>button {
        background-color: #FFD700 !important;
        color: black !important;
        font-weight: bold !important;
        font-size: 18px !important;
        border-radius: 8px !important;
        padding: 10px 20px !important;
        border: none !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2) !important;
        transition: all 0.3s ease !important;
    }
    .stButton>button:hover {
        background-color: #FFC107 !important;
        transform: scale(1.05);
    }
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ------------------ زر تحميل مخصص ------------------
def get_download_link(data, filename, label):
    b64 = base64.b64encode(data).decode()
    href = f"""
        <a href="data:application/octet-stream;base64,{b64}" download="{filename}" 
        style="
            background-color: #FFD700;
            color: black;
            font-weight: bold;
            font-size: 18px;
            border-radius: 8px;
            padding: 10px 20px;
            text-decoration: none;
            display: inline-block;
            box-shadow: 0 2px 4px rgba(0,0,0,0.2);
            transition: all 0.3s ease;
            margin-top: 10px;
        "
        onmouseover="this.style.backgroundColor='#FFC107'; this.style.transform='scale(1.05)'"
        onmouseout="this.style.backgroundColor='#FFD700'; this.style.transform='scale(1)'"
        >
        📥 {label}
        </a>
    """
    return href

# ------------------ عرض اللوجو والمعلومات ------------------
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
        <div style='display:flex; justify-content:space-between; align-items:center; padding:10px 20px;'>
            <img src="data:image/png;base64,{logo_base64}" style='max-height:120px;'>
            <div style='font-size:22px; font-weight:bold; color:#FFD700;'>
                By Admin Mohamed Abd ELGhany – 📱 
                <a href="https://wa.me/201554694554" target="_blank" style="color:#FFD700; text-decoration:none;">
                    01554694554 (WhatsApp)
                </a>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
else:
    st.markdown(
        """
        <div style='text-align:center; margin-bottom:20px; font-size:22px; font-weight:bold; color:#FFD700;'>
            By Admin Mohamed Abd ELGhany – 📱 
            <a href="https://wa.me/201554694554" target="_blank" style="color:#FFD700; text-decoration:none;">
                01554694554 (WhatsApp)
            </a>
        </div>
        """,
        unsafe_allow_html=True
    )

# ------------------ العنوان ------------------
st.markdown("<h1 style='text-align:center; color:#FFD700;'>Averroes Pharma File Splitter</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align:center; color:white;'>✂ Split & Merge Excel Files Fast & Easily</h3>", unsafe_allow_html=True)

# ------------------ رفع الملف ------------------
uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

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
            col_to_split = st.selectbox("Split by Column", df.columns)

            if st.button("🚀 Start Split"):
                with st.spinner("Splitting files..."):
                    zip_buffer = BytesIO()
                    from zipfile import ZipFile

                    with ZipFile(zip_buffer, "a") as zip_file:
                        for value in df[col_to_split].dropna().unique():
                            sub_df = df[df[col_to_split] == value]
                            file_buffer = BytesIO()
                            with pd.ExcelWriter(file_buffer, engine="openpyxl") as writer:
                                sub_df.to_excel(writer, index=False, sheet_name=str(value)[:30])
                            file_buffer.seek(0)
                            zip_file.writestr(f"{str(value)}.xlsx", file_buffer.read())

                    zip_buffer.seek(0)
                    st.success("✅ Files split successfully!")
                    download_link = get_download_link(zip_buffer.getvalue(), f"Split_{selected_sheet}.zip", "Download Split Files (ZIP)")
                    st.markdown(download_link, unsafe_allow_html=True)

        # تحميل كل الشيتات مع بعض
        all_sheets_output = BytesIO()
        with pd.ExcelWriter(all_sheets_output, engine="openpyxl") as writer:
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                df = df.fillna(method="ffill", axis=0).fillna(method="ffill", axis=1)
                df.to_excel(writer, index=False, sheet_name=sheet_name)
        all_sheets_output.seek(0)
        download_all_link = get_download_link(all_sheets_output.getvalue(), "All_Sheets_Cleaned.xlsx", "Download All Sheets (Cleaned)")
        st.markdown(download_all_link, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"❌ Error while processing the file: {e}")
else:
    st.warning("⚠️ No file uploaded yet.")
