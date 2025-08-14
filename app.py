import streamlit as st
import pandas as pd
import zipfile
import io

st.set_page_config(page_title="Averroes Pharma", layout="wide")
st.title("Averroes Pharma - Excel Splitter")

uploaded_file = st.file_uploader("ارفع ملف Excel", type=["xlsx"])

if uploaded_file:
    # قراءة الملف
    df = pd.read_excel(uploaded_file)
    st.subheader("📄 محتوى الملف:")
    st.dataframe(df)

    # اختيار العمود للتقطيع
    column_to_split = st.selectbox("اختر العمود للتقطيع", df.columns)

    if st.button("✂️ تقطيع الملف"):
        unique_values = df[column_to_split].dropna().unique()

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for value in unique_values:
                part_df = df[df[column_to_split] == value]
                file_buffer = io.BytesIO()
                part_df.to_excel(file_buffer, index=False)
                zip_file.writestr(f"{value}.xlsx", file_buffer.getvalue())

        zip_buffer.seek(0)
        st.download_button(
            label="⬇️ تحميل الملفات كـ ZIP",
            data=zip_buffer,
            file_name="split_files.zip",
            mime="application/zip"
        )
