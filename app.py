import streamlit as st
import pandas as pd
from io import BytesIO
import base64

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
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

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
            <div style='font-size:22px; font-weight:bold; color:#FFD700;'>By Admin Mohamed Abd ELGhany – 01554694554</div>
        </div>
        """,
        unsafe_allow_html=True
    )
else:
    st.markdown(
        """
        <div style='text-align:center; margin-bottom:20px; font-size:22px; font-weight:bold; color:#FFD700;'>
            By Admin Mohamed Abd ELGhany – 01554694554
        </div>
        """,
        unsafe_allow_html=True
    )

# ------------------ العنوان ------------------
st.markdown("<h1 style='text-align:center; color:#FFD700;'>Averroes Pharma File Splitter</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align:center; color:white;'>✂ Split & Merge Excel Files Fast & Easily</h3>", unsafe_allow_html=True)

# ------------------ رفع الملف ------------------
uploaded_file = st.file_uploader("📂 Upload  File Excel", type=["xlsx"])

if uploaded_file:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        st.success(f"✅ تم تحميل الملف وفيه {len(excel_file.sheet_names)} شيت.")

        selected_sheet = st.selectbox("📑 Select Sheet:", excel_file.sheet_names)

        if selected_sheet:
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
            df = df.fillna(method="ffill", axis=0).fillna(method="ffill", axis=1)

            st.markdown(f"### 📊 Data View – {selected_sheet}")
            st.dataframe(df, use_container_width=True)

            # ✅ Dropdown لاختيار العمود بعد عرض البيانات
            st.markdown("### ✂ Select the column you want to divide based on:")
            col_to_split = st.selectbox("Select Coulmn:", df.columns)

            if st.button("🚀 Start Spilit"):
                with st.spinner("جاري التقسيم..."):
                    split_dfs = {str(value): df[df[col_to_split] == value] for value in df[col_to_split].unique()}
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        for key, sub_df in split_dfs.items():
                            sheet_name = str(key)[:30]
                            sub_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    output.seek(0)
                    st.success("✅ The files have been successfully divided.!")
                    st.download_button(
                        label="📥Dpwnload File Spilit",
                        data=output.getvalue(),
                        file_name=f"Split_{selected_sheet}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        # تحميل كل الشيتات مع بعض
        all_sheets_output = BytesIO()
        with pd.ExcelWriter(all_sheets_output, engine="openpyxl") as writer:
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                df = df.fillna(method="ffill", axis=0).fillna(method="ffill", axis=1)
                df.to_excel(writer, index=False, sheet_name=sheet_name)
        all_sheets_output.seek(0)
        st.download_button(
            label="⬇️ Download All Files",
            data=all_sheets_output.getvalue(),
            file_name="All_Sheets_Cleaned.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ حدث خطأ أثناء قراءة الملف: {e}")
else:
    st.warning("⚠️ لم يتم رفع أي ملف بعد.")

