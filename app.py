# -*- coding: utf-8 -*-
"""
Streamlit App — Simplified per user request
- Keep: Split & Merge, Image→PDF (Original)
- Add: Excel Processor (rename/delete/reorder columns) from Edit 1
- Add: Folder Images → single PDF (Edit 2 logic adapted)
- Remove: Dashboard section & CamScanner enhancement
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
import re
import os
import base64

from openpyxl.styles import Font, PatternFill, Border, Alignment
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from PIL import Image

# ====== Lottie animation (optional; safe if requests not available) ======
try:
    from streamlit_lottie import st_lottie  # type: ignore
    import requests  # type: ignore
except Exception:  # pragma: no cover
    st_lottie = None
    requests = None


def load_lottie_url(url: str):
    if not requests:
        return None
    try:
        r = requests.get(url, timeout=8)
        if r.status_code == 200:
            return r.json()
    except Exception:
        return None
    return None

# ------------------ Page Setup ------------------
st.set_page_config(
    page_title="Tricks For Excel — Split/Merge & PDF Tools",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

LOTTIE_SPLIT = load_lottie_url("https://assets9.lottiefiles.com/packages/lf20_wx9z5gxb.json")
LOTTIE_MERGE = load_lottie_url("https://assets10.lottiefiles.com/packages/lf20_cg3rwjul.json")
LOTTIE_PDF   = load_lottie_url("https://assets1.lottiefiles.com/packages/lf20_zyu0ct3i.json")

# ------------------ Helper Functions ------------------

def display_uploaded_files(file_list, file_type="Excel/CSV"):
    if file_list:
        st.markdown("### 📁 الملفات المرفوعة:")
        for i, f in enumerate(file_list):
            st.markdown(f"**{i+1}.** {f.name} — {f.size//1024} KB")


def _safe_name(s):
    return re.sub(r"[^A-Za-z0-9_-]+", "_", str(s))


def get_image_as_base64(image_path):
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except Exception:
        return None


# ------------------ Header / Branding ------------------
logo_b64 = get_image_as_base64("logo.png")
if logo_b64:
    st.markdown(
        f'<div style="display:flex;align-items:center;gap:10px">\n'
        f'<img src="data:image/png;base64,{logo_b64}" width="40"/>\n'
        f'<h2 style="margin:0">Tricks For Excel</h2></div>',
        unsafe_allow_html=True,
    )
else:
    st.markdown("## Tricks For Excel")

st.markdown(
    """
    **الأدوات المتاحة:**
    - ✂️ تقسيم ودمج ملفات Excel/CSV
    - 🧰 معالج Excel (إعادة تسمية/حذف/ترتيب أعمدة)
    - 🖼️ تحويل الصور إلى PDF (جودة أصلية) + خيار مجلد محلي → PDF
    """
)

# Keep a counter to allow clearing uploaders
if 'clear_counter' not in st.session_state:
    st.session_state.clear_counter = 0

# ===================== Section: Split =====================
st.markdown("---")
st.markdown("### ✂️ Split Excel/CSV File")

uploaded_file = st.file_uploader(
    "📂 ارفع ملف Excel أو CSV (للتقسيم)",
    type=["xlsx", "csv"],
    accept_multiple_files=False,
    key=f"split_uploader_{st.session_state.clear_counter}",
)

if uploaded_file:
    display_uploaded_files([uploaded_file], "Excel/CSV")
    if st.button("🧹 مسح ملفات التقسيم", key="clear_split"):
        st.session_state.clear_counter += 1
        st.rerun()

    try:
        file_ext = uploaded_file.name.split(".")[-1].lower()
        if file_ext == "csv":
            df = pd.read_csv(uploaded_file)
            sheet_names = ["Sheet1"]
            selected_sheet = "Sheet1"
            st.success("✅ تم رفع ملف CSV بنجاح")
        else:
            input_bytes = uploaded_file.getvalue()
            original_wb = load_workbook(filename=BytesIO(input_bytes), data_only=False)
            sheet_names = original_wb.sheetnames
            st.success(f"✅ تم رفع ملف Excel بنجاح — عدد الشيتات: {len(sheet_names)}")
            selected_sheet = st.selectbox("اختر الشيت للتقسيم", sheet_names)
            df = pd.read_excel(BytesIO(input_bytes), sheet_name=selected_sheet)

        st.markdown(f"#### 👀 معاينة البيانات — {selected_sheet}")
        st.dataframe(df, use_container_width=True)

        col_to_split = st.selectbox("اختر العمود للتقسيم", df.columns)

        split_option = st.radio(
            "طريقة التقسيم:",
            ["Split by Column Values", "Split Each Sheet into Separate File"],
            index=0,
        )

        if st.button("🚀 ابدأ التقسيم"):
            with st.spinner("جاري التقسيم..."):
                if st_lottie and LOTTIE_SPLIT:
                    st_lottie(LOTTIE_SPLIT, height=140, key="lottie_split_main")

                def clean_name(name: str) -> str:
                    name = str(name).strip()
                    invalid_chars = r'[\\/*?:\[\]\n<>"\']'
                    cleaned = re.sub(invalid_chars, "_", name)
                    return cleaned[:30] if cleaned else "Sheet"

                if file_ext == "csv":
                    unique_values = df[col_to_split].dropna().unique()
                    zip_buffer = BytesIO()
                    with ZipFile(zip_buffer, "w") as zip_file:
                        for value in unique_values:
                            filtered_df = df[df[col_to_split] == value]
                            csv_buffer = BytesIO()
                            filtered_df.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                            csv_buffer.seek(0)
                            file_name = f"{clean_name(value)}.csv"
                            zip_file.writestr(file_name, csv_buffer.read())
                    zip_buffer.seek(0)
                    st.success("🎉 تم التقسيم بنجاح!")
                    st.download_button(
                        label="⬇️ تحميل الملفات (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name=f"Split_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip",
                        mime="application/zip",
                    )
                else:
                    if split_option == "Split by Column Values":
                        ws = original_wb[selected_sheet]
                        col_idx = df.columns.get_loc(col_to_split) + 1
                        unique_values = df[col_to_split].dropna().unique()
                        zip_buffer = BytesIO()
                        with ZipFile(zip_buffer, "w") as zip_file:
                            for value in unique_values:
                                new_wb = Workbook()
                                default_ws = new_wb.active
                                new_wb.remove(default_ws)
                                new_ws = new_wb.create_sheet(title=clean_name(value))

                                # copy header row w/ simple styling
                                for cell in ws[1]:
                                    dst = new_ws.cell(1, cell.column, cell.value)
                                    if cell.has_style:
                                        try:
                                            if cell.font:
                                                dst.font = Font(name=cell.font.name, size=cell.font.size,
                                                                bold=cell.font.bold, italic=cell.font.italic,
                                                                color=cell.font.color)
                                            if cell.fill and cell.fill.fill_type:
                                                dst.fill = PatternFill(fill_type=cell.fill.fill_type,
                                                                       start_color=cell.fill.start_color,
                                                                       end_color=cell.fill.end_color)
                                            if cell.border:
                                                dst.border = Border(left=cell.border.left, right=cell.border.right,
                                                                    top=cell.border.top, bottom=cell.border.bottom)
                                            if cell.alignment:
                                                dst.alignment = Alignment(horizontal=cell.alignment.horizontal,
                                                                          vertical=cell.alignment.vertical,
                                                                          wrap_text=cell.alignment.wrap_text)
                                            dst.number_format = cell.number_format
                                        except Exception:
                                            pass

                                # copy matching rows
                                row_out = 2
                                for row in ws.iter_rows(min_row=2):
                                    if row[col_idx - 1].value == value:
                                        for src in row:
                                            dst = new_ws.cell(row_out, src.column, src.value)
                                            if src.has_style:
                                                try:
                                                    if src.font:
                                                        dst.font = Font(name=src.font.name, size=src.font.size,
                                                                        bold=src.font.bold, italic=src.font.italic,
                                                                        color=src.font.color)
                                                    if src.fill and src.fill.fill_type:
                                                        dst.fill = PatternFill(fill_type=src.fill.fill_type,
                                                                               start_color=src.fill.start_color,
                                                                               end_color=src.fill.end_color)
                                                    if src.border:
                                                        dst.border = Border(left=src.border.left, right=src.border.right,
                                                                            top=src.border.top, bottom=src.border.bottom)
                                                    if src.alignment:
                                                        dst.alignment = Alignment(horizontal=src.alignment.horizontal,
                                                                                  vertical=src.alignment.vertical,
                                                                                  wrap_text=src.alignment.wrap_text)
                                                    dst.number_format = src.number_format
                                                except Exception:
                                                    pass
                                        row_out += 1

                                # copy column widths
                                try:
                                    for col_letter in ws.column_dimensions:
                                        width = ws.column_dimensions[col_letter].width
                                        if width:
                                            new_ws.column_dimensions[col_letter].width = width
                                except Exception:
                                    pass

                                file_buffer = BytesIO()
                                new_wb.save(file_buffer)
                                file_buffer.seek(0)
                                file_name = f"{clean_name(value)}.xlsx"
                                zip_file.writestr(file_name, file_buffer.read())
                        zip_buffer.seek(0)
                        st.success("🎉 تم التقسيم بنجاح!")
                        st.download_button(
                            label="⬇️ تحميل الملفات (ZIP)",
                            data=zip_buffer.getvalue(),
                            file_name=f"Split_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip",
                            mime="application/zip",
                        )
                    else:  # Split Each Sheet into Separate File
                        zip_buffer = BytesIO()
                        with ZipFile(zip_buffer, "w") as zip_file:
                            for sheet_name in original_wb.sheetnames:
                                new_wb = Workbook()
                                default_ws = new_wb.active
                                new_wb.remove(default_ws)
                                new_ws = new_wb.create_sheet(title=sheet_name)
                                src_ws = original_wb[sheet_name]
                                for row in src_ws.iter_rows():
                                    for src_cell in row:
                                        dst = new_ws.cell(src_cell.row, src_cell.column, src_cell.value)
                                        if src_cell.has_style:
                                            try:
                                                if src_cell.font:
                                                    dst.font = Font(name=src_cell.font.name, size=src_cell.font.size,
                                                                    bold=src_cell.font.bold, italic=src_cell.font.italic,
                                                                    color=src_cell.font.color)
                                                if src_cell.fill and src_cell.fill.fill_type:
                                                    dst.fill = PatternFill(fill_type=src_cell.fill.fill_type,
                                                                           start_color=src_cell.fill.start_color,
                                                                           end_color=src_cell.fill.end_color)
                                                if src_cell.border:
                                                    dst.border = Border(left=src_cell.border.left, right=src_cell.border.right,
                                                                        top=src_cell.border.top, bottom=src_cell.border.bottom)
                                                if src_cell.alignment:
                                                    dst.alignment = Alignment(horizontal=src_cell.alignment.horizontal,
                                                                              vertical=src_cell.alignment.vertical,
                                                                              wrap_text=src_cell.alignment.wrap_text)
                                                dst.number_format = src_cell.number_format
                                            except Exception:
                                                pass
                                try:
                                    for col_letter in src_ws.column_dimensions:
                                        if src_ws.column_dimensions[col_letter].width:
                                            new_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width
                                except Exception:
                                    pass
                                file_buffer = BytesIO()
                                new_wb.save(file_buffer)
                                file_buffer.seek(0)
                                file_name = f"{_safe_name(sheet_name)}.xlsx"
                                zip_file.writestr(file_name, file_buffer.read())
                        zip_buffer.seek(0)
                        st.success("🎉 تم التقسيم بنجاح!")
                        st.download_button(
                            label="⬇️ تحميل الملفات (ZIP)",
                            data=zip_buffer.getvalue(),
                            file_name=f"SplitBySheets_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip",
                            mime="application/zip",
                        )
    except Exception as e:
        st.error(f"❌ حدث خطأ أثناء التقسيم: {e}")
else:
    st.info("⬆️ ارفع ملفًا للبدء في التقسيم")


# ===================== Section: Merge =====================
st.markdown("---")
st.markdown("### 🔁 Merge Excel/CSV Files")

merge_files = st.file_uploader(
    "📂 ارفع ملفات Excel/CSV للدمج",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    key=f"merge_uploader_{st.session_state.clear_counter}",
)

if merge_files:
    display_uploaded_files(merge_files, "Excel/CSV")
    if st.button("🧹 مسح ملفات الدمج", key="clear_merge"):
        st.session_state.clear_counter += 1
        st.rerun()

    if st.button("✨ دمج الملفات"):
        with st.spinner("جاري الدمج..."):
            if st_lottie and LOTTIE_MERGE:
                st_lottie(LOTTIE_MERGE, height=140, key="lottie_merge_main")
            try:
                all_excel = all(f.name.lower().endswith('.xlsx') for f in merge_files)
                if all_excel:
                    merged_wb = Workbook()
                    merged_ws = merged_wb.active
                    merged_ws.title = "Merged_Data"
                    current_row = 1
                    for idx, file in enumerate(merge_files):
                        file_bytes = file.getvalue()
                        src_wb = load_workbook(filename=BytesIO(file_bytes), data_only=False)
                        src_ws = src_wb.active
                        if idx == 0:
                            # header
                            for row in src_ws.iter_rows(min_row=1, max_row=1):
                                for cell in row:
                                    dst = merged_ws.cell(current_row, cell.column, cell.value)
                                    if cell.has_style:
                                        try:
                                            if cell.font:
                                                dst.font = Font(name=cell.font.name, size=cell.font.size,
                                                                bold=cell.font.bold, italic=cell.font.italic,
                                                                color=cell.font.color)
                                            if cell.fill and cell.fill.fill_type:
                                                dst.fill = PatternFill(fill_type=cell.fill.fill_type,
                                                                       start_color=cell.fill.start_color,
                                                                       end_color=cell.fill.end_color)
                                            if cell.border:
                                                dst.border = Border(left=cell.border.left, right=cell.border.right,
                                                                    top=cell.border.top, bottom=cell.border.bottom)
                                            if cell.alignment:
                                                dst.alignment = Alignment(horizontal=cell.alignment.horizontal,
                                                                          vertical=cell.alignment.vertical,
                                                                          wrap_text=cell.alignment.wrap_text)
                                            dst.number_format = cell.number_format
                                        except Exception:
                                            pass
                            current_row += 1
                        for row in src_ws.iter_rows(min_row=2):
                            for cell in row:
                                dst = merged_ws.cell(current_row, cell.column, cell.value)
                                if cell.has_style:
                                    try:
                                        if cell.font:
                                            dst.font = Font(name=cell.font.name, size=cell.font.size,
                                                            bold=cell.font.bold, italic=cell.font.italic,
                                                            color=cell.font.color)
                                        if cell.fill and cell.fill.fill_type:
                                            dst.fill = PatternFill(fill_type=cell.fill.fill_type,
                                                                   start_color=cell.fill.start_color,
                                                                   end_color=cell.fill.end_color)
                                        if cell.border:
                                            dst.border = Border(left=cell.border.left, right=cell.border.right,
                                                                top=cell.border.top, bottom=cell.border.bottom)
                                        if cell.alignment:
                                            dst.alignment = Alignment(horizontal=cell.alignment.horizontal,
                                                                      vertical=cell.alignment.vertical,
                                                                      wrap_text=cell.alignment.wrap_text)
                                        dst.number_format = cell.number_format
                                    except Exception:
                                        pass
                            current_row += 1

                        # carry over column widths from latest file
                        try:
                            for col_letter in src_ws.column_dimensions:
                                width = src_ws.column_dimensions[col_letter].width
                                if width:
                                    merged_ws.column_dimensions[col_letter].width = width
                        except Exception:
                            pass

                    output_buffer = BytesIO()
                    merged_wb.save(output_buffer)
                    output_buffer.seek(0)
                    st.success("✅ تم الدمج مع المحافظة على التنسيقات")
                    st.download_button(
                        label="⬇️ تحميل الملف المدمج (Excel)",
                        data=output_buffer.getvalue(),
                        file_name="Merged_Consolidated_Formatted.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    # Mixed or CSV — merge with pandas (formatting not preserved)
                    all_dfs = []
                    for file in merge_files:
                        ext = file.name.split(".")[-1].lower()
                        df = pd.read_csv(file) if ext == "csv" else pd.read_excel(file)
                        all_dfs.append(df)
                    merged_df = pd.concat(all_dfs, ignore_index=True)
                    output_buffer = BytesIO()
                    merged_df.to_excel(output_buffer, index=False, engine='openpyxl')
                    output_buffer.seek(0)
                    st.success("✅ تم الدمج (قد لا تُحفظ التنسيقات لملفات CSV/مختلطة)")
                    st.download_button(
                        label="⬇️ تحميل الملف المدمج (Excel)",
                        data=output_buffer.getvalue(),
                        file_name="Merged_Consolidated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            except Exception as e:
                st.error(f"❌ خطأ أثناء الدمج: {e}")
else:
    st.info("⬆️ ارفع ملفات للدمج")


# ===================== Section: Excel Processor (from Edit 1) =====================
st.markdown("---")
st.markdown("### 🧰 معالج Excel (إعادة تسمية/حذف/ترتيب أعمدة)")

proc_file = st.file_uploader(
    "📂 ارفع ملف Excel للمعالجة (xlsx/xlsm)",
    type=["xlsx", "xlsm"],
    accept_multiple_files=False,
    key=f"processor_uploader_{st.session_state.clear_counter}",
)

# Settings replicated from Edit 1
COLUMNS_TO_DELETE = [
    "Status","Status Date","Assigned To","Employees Count","Attendees Count",
    "Not Listed Invitees Count","Cost Per Person","Governorate","Accomodation Type",
    "CSR Enabled","Early Bird Due Date","Reservation Type","Meal Type","Meal Title",
    "Delivery Type","Start Date","End Date","Budget Date","Delivery Date","Invoice Date",
    "Actual Cost","Deduct","Net Amount","Other Professionals","Customers","Reps",
    "Items\\Brands","Created At","Request Professional Classifications",
    "Request Professional Ids","Segments","Accounts","Account Professionals","Category",
    "Type","Request Serial Number","Shared","Updated","L5 Emp Name",
    "Sponsored Company Name","Item Type","Item Brand","Promotional Code","Venue",
    "Restaurant","Business Type","Highlighted","Link Details",
]

COLUMN_RENAME_MAP = {
    "L1 Emp Name": "MR",
    "L2 Emp Name": "DM",
    "L3 Emp Name": "AM",
}

FINAL_COLUMN_ORDER = [
    "Tracking Number","MR","DM","AM","bum","Line","Activity","Description",
    "Account Number","Vendor","Bank","Cost","Bricks","Professionl Accounts",
    "Request Professionals","Specialities","Request Date",
]

if proc_file:
    st.write("**الملف:**", proc_file.name)
    if st.button("⚙️ تنفيذ المعالجة"):
        try:
            wb = load_workbook(proc_file, data_only=False)
            ws = wb.active

            # Read headers
            header_row = 1
            headers = [ws.cell(header_row, col).value for col in range(1, ws.max_column + 1)]
            header_to_idx = {h: i+1 for i, h in enumerate(headers) if h is not None}

            # Rename
            st.write("🔤 جاري إعادة تسمية الأعمدة...")
            for old_name, new_name in COLUMN_RENAME_MAP.items():
                if old_name in header_to_idx:
                    cidx = header_to_idx[old_name]
                    ws.cell(header_row, cidx).value = new_name
                    headers[cidx-1] = new_name
                    header_to_idx[new_name] = cidx
                    del header_to_idx[old_name]
                    st.write(f"✔️ تم تغيير `{old_name}` → `{new_name}`")

            # Delete unwanted columns
            st.write("🧽 جاري حذف الأعمدة غير المطلوبة...")
            # Delete from right to left to keep indices stable
            to_delete_indices = [header_to_idx[c] for c in COLUMNS_TO_DELETE if c in header_to_idx]
            to_delete_indices.sort(reverse=True)
            for cidx in to_delete_indices:
                ws.delete_cols(cidx)
            st.write(f"🗑️ تم حذف {len(to_delete_indices)} عمودًا")

            # Rebuild headers & mapping
            headers = [ws.cell(header_row, col).value for col in range(1, ws.max_column + 1)]
            header_to_idx = {h: i+1 for i, h in enumerate(headers) if h is not None}

            # Reorder into a new workbook (preserving simple style)
            st.write("📐 جاري إعادة ترتيب الأعمدة...")
            col_mapping = []
            for name in FINAL_COLUMN_ORDER:
                if name in header_to_idx:
                    col_mapping.append(header_to_idx[name])
                    st.write(f"➕ تضمين `{name}`")
                else:
                    col_mapping.append(None)
                    st.warning(f"⚠️ العمود `{name}` غير موجود")

            new_wb = Workbook()
            new_ws = new_wb.active

            for row_idx in range(1, ws.max_row + 1):
                for new_col_idx, old_col_idx in enumerate(col_mapping, start=1):
                    if old_col_idx is None:
                        continue
                    old_cell = ws.cell(row_idx, old_col_idx)
                    new_cell = new_ws.cell(row_idx, new_col_idx)
                    new_cell.value = old_cell.value
                    if old_cell.has_style:
                        try:
                            if old_cell.font:
                                new_cell.font = Font(name=old_cell.font.name, size=old_cell.font.size,
                                                     bold=old_cell.font.bold, italic=old_cell.font.italic,
                                                     color=old_cell.font.color)
                            if old_cell.fill and old_cell.fill.fill_type:
                                new_cell.fill = PatternFill(fill_type=old_cell.fill.fill_type,
                                                            start_color=old_cell.fill.start_color,
                                                            end_color=old_cell.fill.end_color)
                            if old_cell.border:
                                new_cell.border = old_cell.border
                            if old_cell.alignment:
                                new_cell.alignment = Alignment(horizontal=old_cell.alignment.horizontal,
                                                               vertical=old_cell.alignment.vertical,
                                                               wrap_text=old_cell.alignment.wrap_text)
                            new_cell.number_format = old_cell.number_format
                        except Exception:
                            pass

            # Set uniform column widths
            for c in range(1, len(FINAL_COLUMN_ORDER) + 1):
                new_ws.column_dimensions[get_column_letter(c)].width = 15

            out_buf = BytesIO()
            new_wb.save(out_buf)
            out_buf.seek(0)
            st.success("✅ تم إتمام المعالجة بنجاح")
            base = os.path.splitext(proc_file.name)[0]
            st.download_button(
                label="⬇️ تحميل الملف المعالج",
                data=out_buf.getvalue(),
                file_name=f"{_safe_name(base)}_processed.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"❌ خطأ أثناء المعالجة: {e}")
else:
    st.info("⬆️ ارفع ملف Excel لبدء المعالجة")


# ===================== Section: Images → PDF =====================
st.markdown("---")
st.markdown("### 🖼️ تحويل الصور إلى PDF (جودة أصلية)")

uploaded_images = st.file_uploader(
    "📂 ارفع صور JPG/PNG للتحويل إلى PDF (يمكن تحديد عدة صور)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
    key=f"image_uploader_{st.session_state.clear_counter}",
)

if uploaded_images:
    display_uploaded_files(uploaded_images, "Image")
    if st.button("🧹 مسح الصور", key="clear_images"):
        st.session_state.clear_counter += 1
        st.rerun()

    if st.button("🖨️ إنشاء PDF (جودة أصلية)"):
        with st.spinner("جاري إنشاء ملف PDF..."):
            try:
                first = Image.open(uploaded_images[0]).convert("RGB")
                others = [Image.open(x).convert("RGB") for x in uploaded_images[1:]]
                pdf_buffer = BytesIO()
                first.save(pdf_buffer, format="PDF", save_all=True, append_images=others)
                pdf_buffer.seek(0)
                st.success("✅ تم إنشاء ملف PDF بنجاح")
                st.download_button(
                    label="⬇️ تحميل ملف PDF",
                    data=pdf_buffer.getvalue(),
                    file_name="Images_Combined.pdf",
                    mime="application/pdf",
                )
            except Exception as e:
                st.error(f"❌ خطأ أثناء إنشاء PDF: {e}")
else:
    st.info("⬆️ ارفع صورًا للتحويل إلى PDF")

# ---- Optional: Folder path → PDF (Edit 2 logic) ----
st.markdown("#### 🗂️ تحويل مجلد صور محلي إلى PDF (اختياري)")
folder_path = st.text_input("📁 مسار المجلد (على جهازك المحلي)", value="")
if st.button("🖨️ إنشاء PDF من مجلد محلي"):
    try:
        if not folder_path or not os.path.exists(folder_path):
            st.error("المسار غير صحيح أو غير موجود")
        else:
            image_files = [f for f in os.listdir(folder_path) if f.lower().endswith((".jpg", ".jpeg", ".png"))]
            if not image_files:
                st.warning("لا توجد صور في هذا المجلد")
            else:
                # Sort to ensure consistent order
                image_files.sort()
                first_path = os.path.join(folder_path, image_files[0])
                first_img = Image.open(first_path).convert("RGB")
                others = []
                for f in image_files[1:]:
                    p = os.path.join(folder_path, f)
                    img = Image.open(p).convert("RGB")
                    others.append(img)
                pdf_buffer = BytesIO()
                first_img.save(pdf_buffer, format="PDF", save_all=True, append_images=others)
                pdf_buffer.seek(0)
                st.success("✅ تم إنشاء PDF من المجلد")
                st.download_button(
                    label="⬇️ تحميل PDF المجلد",
                    data=pdf_buffer.getvalue(),
                    file_name="Folder_Images.pdf",
                    mime="application/pdf",
                )
    except Exception as e:
        st.error(f"❌ خطأ: {e}")


st.markdown("---")
st.markdown("#### 📞 تواصل: [WhatsApp](https://wa.me/01554694554)")
