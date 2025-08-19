import streamlit as st
import pandas as pd
from io import BytesIO
import base64

# ------------------ إعدادات الصفحة ------------------
st.set_page_config(
    page_title="Averroes Pharma Splitter",
    page_icon="💊",
    layout="wide",  # يمكن تركه wide
    initial_sidebar_state="expanded"  # ← هذا مهم: يفتح Sidebar تلقائيًا
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

    /* جعل Sidebar واضحًا وبارزًا */
    [data-testid="stSidebar"] {
        background-color: #003366 !important;
        color: white !important;
        border-right: 4px solid #FFD700 !important;
        width: 300px !important; /* عرض ثابت */
        min-height: 100vh;
        box-shadow: 2px 0 5px rgba(0,0,0,0.3);
    }

    [data-testid="stSidebar"] .css-1d391kg {
        color: #FFD700 !important;
        font-size: 22px !important;
        font-weight: bold;
        text-align: center;
        margin-bottom: 15px;
    }

    [data-testid="stSidebar"] .css-1v3fvvy,
    [data-testid="stSidebar"] .css-1l02zno {
        color: white !important;
        font-size: 18px !important;
    }

    [data-testid="stSidebar"] .stButton>button {
        background-color: #FFD700 !important;
        color: black !important;
        border-radius: 10px;
        font-weight: bold;
        font-size: 18px;
        padding: 12px 20px;
        border: none;
        width: 100%;
        margin: 10px 0;
    }
    [data-testid="stSidebar"] .stButton>button:hover {
        background-color: #daa520 !important;
        transform: scale(1.03);
    }

    [data-testid="stSidebar"] .stSelectbox label {
        color: #FFD700 !important;
        font-weight: bold;
        font-size: 18px;
    }

    [data-testid="stSidebar"] .stSelectbox div[data-baseweb="select"] {
        background-color: #002b4d !important;
        color: white !important;
        border: 2px solid #FFD700 !important;
        border-radius: 8px;
    }

    /* تنسيق الصفحة الرئيسية */
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
        color: #FFD700;
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
st.markdown("<div class='subtitle'>✂ Spilit & Merge Files Excel Fast & Easily</div>", unsafe_allow_html=True)

# ------------------ رفع الملف ------------------
uploaded_file = st.file_uploader("📂 ارفع ملف Excel", type=["xlsx"])

if uploaded_file:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        st.success(f"✅ تم تحميل الملف وفيه {len(excel_file.sheet_names)} شيت.")

        # اختيار الشيت من سايدبار
        st.sidebar.markdown("<h3 style='color:#FFD700; text-align:center;'>📑 اختيار الورقة</h3>", unsafe_allow_html=True)
        selected_sheet = st.sidebar.selectbox(
            "اختر الورقة المراد معالجتها:",
            excel_file.sheet_names,
            key="select_sheet"
        )

        if selected_sheet:
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

            # معالجة merge cells
            df = df.fillna(method="ffill", axis=0).fillna(method="ffill", axis=1)

            # عرض بيانات الشيت بالكامل
            with st.expander(f"📊 عرض البيانات - {selected_sheet}"):
                st.dataframe(df, use_container_width=True)

            # ================= Sidebar Options ==================
            st.sidebar.markdown("---")
            st.sidebar.markdown("<h3 style='color:#FFD700; text-align:center;'>⚙️ إعدادات التقسيم</h3>", unsafe_allow_html=True)

            col_to_split = st.sidebar.selectbox(
                "اختر العمود المراد التقسيم بناءً عليه:",
                df.columns,
                key="select_col"
            )

            if st.sidebar.button("🚀 بدء التقسيم"):
                with st.spinner("جاري التقسيم..."):
                    split_dfs = {str(value): df[df[col_to_split] == value] for value in df[col_to_split].unique()}

                    output = BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        for key, sub_df in split_dfs.items():
                            sheet_name = str(key)[:30]
                            sub_df.to_excel(writer, sheet_name=sheet_name, index=False)

                    output.seek(0)
                    st.success("✅ تم التقسيم بنجاح!")

                    st.download_button(
                        label="📥 تنزيل الملف المنقسم",
                        data=output.getvalue(),
                        file_name=f"Split_{selected_sheet}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        # ---------------- تحميل كل الشيتات مع بعض ----------------
        all_sheets_output = BytesIO()
        with pd.ExcelWriter(all_sheets_output, engine="openpyxl") as writer:
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                df = df.fillna(method="ffill", axis=0).fillna(method="ffill", axis=1)
                df.to_excel(writer, index=False, sheet_name=sheet_name)

        all_sheets_output.seek(0)

        st.download_button(
            label="⬇️ تنزيل جميع الأوراق (نسخة نظيفة)",
            data=all_sheets_output.getvalue(),
            file_name="All_Sheets_Cleaned.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="all_sheets"
        )

    except Exception as e:
        st.error(f"❌ حدث خطأ أثناء قراءة الملف: {e}")
