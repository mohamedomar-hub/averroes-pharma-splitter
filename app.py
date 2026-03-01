# -*- coding: utf-8 -*-
"""
Streamlit App — Polished UI per user request
Changes vs previous revision:
- Removed: "Folder Images → PDF (optional)" section completely
- Added: custom modern UI theme (dark-friendly) via CSS
- Tweaks: section cards, colored headers, buttons, containers
- Kept: Split, Merge, Excel Processor, Images→PDF (original quality)
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

# Optional animations
try:
    from streamlit_lottie import st_lottie  # type: ignore
    import requests  # type: ignore
except Exception:
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

# ------------------ Custom CSS (polished UI) ------------------
custom_css = """
<style>
:root {
  --primary: #6c5ce7;         /* Purple */
  --primary-2: #a29bfe;       /* Soft purple */
  --accent: #00d1b2;          /* Teal */
  --bg-elev: #111418;         /* Card bg (dark) */
  --bg-elev-2: #0c0f13;       /* Header bg */
  --text: #e8ebee;            /* Main text */
  --muted: #9aa3ad;           /* Muted */
  --border: #232a33;          /* Border */
}
/* Light theme fallback */
@media (prefers-color-scheme: light){
  :root {
    --bg-elev: #ffffff;
    --bg-elev-2: #f7f8fa;
    --text: #202530;
    --muted: #5b6570;
    --border: #e8edf3;
  }
}

/* Base */
html, body, [class^="css"]  { font-family: 'Segoe UI', system-ui, -apple-system, Cairo, Tahoma, sans-serif; }

/* App wide background harmony */
section.main > div { padding-top: 10px; }

/* Title row */
.app-header {
  display:flex; align-items:center; gap:12px; padding:14px 18px;
  background: linear-gradient(135deg, var(--bg-elev-2), transparent);
  border: 1px solid var(--border); border-radius: 14px;
}
.app-title { margin:0; font-weight:800; letter-spacing:.3px; color: var(--text); }
.app-sub { margin:2px 0 0; color: var(--muted); font-size: 14px; }

/* Section card */
.card { border:1px solid var(--border); border-radius:16px; padding:18px 18px 8px; background: var(--bg-elev); margin: 8px 0 18px; }
.card h3 { margin-top:0; display:flex; align-items:center; gap:8px; color: var(--text);}
.card .hint { color: var(--muted); font-size: 13px; margin-top:-8px; margin-bottom:10px; }

/* Buttons */
.stButton > button {
  border-radius:12px; padding:10px 14px; font-weight:600; border:1px solid transparent;
  background: linear-gradient(135deg, var(--primary), var(--primary-2)); color:#fff; box-shadow: 0 3px 10px rgba(108,92,231,.25);
}
.stButton > button:hover { filter:brightness(1.06); transform: translateY(-1px); }
.stButton > button:active { transform: translateY(0); }

/* File uploader */
[data-testid="stFileUploader"] { background: rgba(255,255,255,.02); padding:10px; border-radius: 12px; border:1px dashed var(--border); }

/* Download buttons */
.stDownloadButton > button { border-radius:10px; border:1px solid var(--border); background: var(--bg-elev-2); color: var(--text); }

/* Dataframe wrapper */
.css-1m1b9qw, .stDataFrame { border-radius: 10px; overflow:hidden; border:1px solid var(--border); }

/* Divider */
hr { border: none; height: 1px; background: var(--border); margin: 14px 0; }

.badge { display:inline-flex; align-items:center; gap:6px; padding:6px 10px; border-radius:999px; background: rgba(0,209,178,.12); color:#28d6bd; font-weight:600; border:1px solid rgba(0,209,178,.25); font-size:12px }

</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ------------------ Small helpers ------------------

def display_uploaded_files(file_list, file_type="ملفات"):
    if file_list:
        st.markdown("**الملفات المرفوعة:**")
        for i, f in enumerate(file_list):
            st.caption(f"{i+1}. {f.name} — {f.size//1024} KB")

def _safe_name(s):
    return re.sub(r"[^A-Za-z0-9_-]+", "_", str(s))


def get_image_as_base64(image_path):
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except Exception:
        return None

# ------------------ Header ------------------
logo_b64 = get_image_as_base64("logo.png")
col_logo, col_title = st.columns([1,7])
with col_logo:
    if logo_b64:
        st.image(f"data:image/png;base64,{logo_b64}", width=48)
with col_title:
    st.markdown('<div class="app-header"><div><h2 class="app-title">Tricks For Excel</h2><p class="app-sub">أدوات سريعة لملفات Excel والصور • Split • Merge • Processor • PDF</p></div></div>', unsafe_allow_html=True)

if 'clear_counter' not in st.session_state:
    st.session_state.clear_counter = 0

# ===================== Split Card =====================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### ✂️ Split Excel/CSV File")
    st.markdown('<div class="hint">ارفع ملفًا ثم اختر عمود التقسيم، وسيُنشأ ملف ZIP للتحميل</div>', unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "📂 ارفع ملف Excel أو CSV",
        type=["xlsx", "csv"],
        accept_multiple_files=False,
        key=f"split_uploader_{st.session_state.clear_counter}",
    )

    if uploaded_file:
        display_uploaded_files([uploaded_file])
        c1, c2 = st.columns([1,1])
        with c1:
            if st.button("🧹 مسح الملفات", key="clear_split"):
                st.session_state.clear_counter += 1
                st.rerun()

        try:
            file_ext = uploaded_file.name.split(".")[-1].lower()
            if file_ext == "csv":
                df = pd.read_csv(uploaded_file)
                selected_sheet = "Sheet1"
                st.success("✅ تم رفع ملف CSV بنجاح")
            else:
                input_bytes = uploaded_file.getvalue()
                original_wb = load_workbook(filename=BytesIO(input_bytes), data_only=False)
                sheet_names = original_wb.sheetnames
                selected_sheet = st.selectbox("اختر الشيت للتقسيم", sheet_names)
                df = pd.read_excel(BytesIO(input_bytes), sheet_name=selected_sheet)

            st.dataframe(df.head(200), use_container_width=True)
            col_to_split = st.selectbox("اختر العمود للتقسيم", df.columns)
            split_option = st.radio(
                "طريقة التقسيم:",
                ["Split by Column Values", "Split Each Sheet into Separate File"],
                horizontal=True,
            )

            if st.button("🚀 ابدأ التقسيم"):
                with st.spinner("جاري التقسيم..."):
                    if st_lottie and LOTTIE_SPLIT:
                        st_lottie(LOTTIE_SPLIT, height=120, key="lottie_split")

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
                                zip_file.writestr(f"{clean_name(value)}.csv", csv_buffer.read())
                        zip_buffer.seek(0)
                        st.success("🎉 تم التقسيم بنجاح!")
                        st.download_button("⬇️ تحميل (ZIP)", zip_buffer.getvalue(), file_name=f"Split_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip", mime="application/zip")
                    else:
                        ws = original_wb[selected_sheet]
                        if split_option == "Split by Column Values":
                            col_idx = df.columns.get_loc(col_to_split) + 1
                            unique_values = df[col_to_split].dropna().unique()
                            zip_buffer = BytesIO()
                            with ZipFile(zip_buffer, "w") as zip_file:
                                for value in unique_values:
                                    new_wb = Workbook(); default_ws = new_wb.active; new_wb.remove(default_ws)
                                    new_ws = new_wb.create_sheet(title=clean_name(value))
                                    # Header
                                    for cell in ws[1]:
                                        dst = new_ws.cell(1, cell.column, cell.value)
                                        if cell.has_style:
                                            try:
                                                if cell.font:
                                                    dst.font = Font(name=cell.font.name, size=cell.font.size, bold=cell.font.bold, italic=cell.font.italic, color=cell.font.color)
                                                if cell.fill and cell.fill.fill_type:
                                                    dst.fill = PatternFill(fill_type=cell.fill.fill_type, start_color=cell.fill.start_color, end_color=cell.fill.end_color)
                                                if cell.border:
                                                    dst.border = cell.border
                                                if cell.alignment:
                                                    dst.alignment = Alignment(horizontal=cell.alignment.horizontal, vertical=cell.alignment.vertical, wrap_text=cell.alignment.wrap_text)
                                                dst.number_format = cell.number_format
                                            except Exception: pass
                                    # Rows
                                    row_out = 2
                                    for row in ws.iter_rows(min_row=2):
                                        if row[col_idx - 1].value == value:
                                            for src in row:
                                                dst = new_ws.cell(row_out, src.column, src.value)
                                                if src.has_style:
                                                    try:
                                                        if src.font:
                                                            dst.font = Font(name=src.font.name, size=src.font.size, bold=src.font.bold, italic=src.font.italic, color=src.font.color)
                                                        if src.fill and src.fill.fill_type:
                                                            dst.fill = PatternFill(fill_type=src.fill.fill_type, start_color=src.fill.start_color, end_color=src.fill.end_color)
                                                        if src.border:
                                                            dst.border = src.border
                                                        if src.alignment:
                                                            dst.alignment = Alignment(horizontal=src.alignment.horizontal, vertical=src.alignment.vertical, wrap_text=src.alignment.wrap_text)
                                                        dst.number_format = src.number_format
                                                    except Exception: pass
                                            row_out += 1
                                    try:
                                        for col_letter in ws.column_dimensions:
                                            width = ws.column_dimensions[col_letter].width
                                            if width:
                                                new_ws.column_dimensions[col_letter].width = width
                                    except Exception: pass
                                    fb = BytesIO(); new_wb.save(fb); fb.seek(0)
                                    zip_file.writestr(f"{clean_name(value)}.xlsx", fb.read())
                            zip_buffer.seek(0)
                            st.success("🎉 تم التقسيم بنجاح!")
                            st.download_button("⬇️ تحميل (ZIP)", zip_buffer.getvalue(), file_name=f"Split_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip", mime="application/zip")
                        else:
                            zip_buffer = BytesIO()
                            with ZipFile(zip_buffer, "w") as zip_file:
                                for sheet_name in original_wb.sheetnames:
                                    new_wb = Workbook(); default_ws = new_wb.active; new_wb.remove(default_ws)
                                    new_ws = new_wb.create_sheet(title=sheet_name)
                                    src_ws = original_wb[sheet_name]
                                    for row in src_ws.iter_rows():
                                        for src_cell in row:
                                            dst = new_ws.cell(src_cell.row, src_cell.column, src_cell.value)
                                            if src_cell.has_style:
                                                try:
                                                    if src_cell.font:
                                                        dst.font = Font(name=src_cell.font.name, size=src_cell.font.size, bold=src_cell.font.bold, italic=src_cell.font.italic, color=src_cell.font.color)
                                                    if src_cell.fill and src_cell.fill.fill_type:
                                                        dst.fill = PatternFill(fill_type=src_cell.fill.fill_type, start_color=src_cell.fill.start_color, end_color=src_cell.fill.end_color)
                                                    if src_cell.border:
                                                        dst.border = src_cell.border
                                                    if src_cell.alignment:
                                                        dst.alignment = Alignment(horizontal=src_cell.alignment.horizontal, vertical=src_cell.alignment.vertical, wrap_text=src_cell.alignment.wrap_text)
                                                    dst.number_format = src_cell.number_format
                                                except Exception: pass
                                    try:
                                        for col_letter in src_ws.column_dimensions:
                                            if src_ws.column_dimensions[col_letter].width:
                                                new_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width
                                    except Exception: pass
                                    fb = BytesIO(); new_wb.save(fb); fb.seek(0)
                                    zip_file.writestr(f"{_safe_name(sheet_name)}.xlsx", fb.read())
                            zip_buffer.seek(0)
                            st.success("🎉 تم التقسيم بنجاح!")
                            st.download_button("⬇️ تحميل (ZIP)", zip_buffer.getvalue(), file_name=f"SplitBySheets_{_safe_name(uploaded_file.name.rsplit('.',1)[0])}.zip", mime="application/zip")
        except Exception as e:
            st.error(f"❌ حدث خطأ أثناء التقسيم: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

# ===================== Merge Card =====================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 🔁 Merge Excel/CSV Files")
    st.markdown('<div class="hint">ارفع عدّة ملفات وسيتم دمجها في ملف واحد مع الحفاظ على التنسيقات لملفات Excel</div>', unsafe_allow_html=True)

    merge_files = st.file_uploader(
        "📂 ارفع ملفات Excel/CSV للدمج",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        key=f"merge_uploader_{st.session_state.clear_counter}",
    )

    if merge_files:
        display_uploaded_files(merge_files)
        c1, c2 = st.columns([1,1])
        with c1:
            if st.button("🧹 مسح الملفات", key="clear_merge"):
                st.session_state.clear_counter += 1
                st.rerun()
        with c2:
            if st.button("✨ دمج الملفات"):
                with st.spinner("جاري الدمج..."):
                    if st_lottie and LOTTIE_MERGE:
                        st_lottie(LOTTIE_MERGE, height=110, key="lottie_merge")
                    try:
                        all_excel = all(f.name.lower().endswith('.xlsx') for f in merge_files)
                        if all_excel:
                            merged_wb = Workbook(); merged_ws = merged_wb.active; merged_ws.title = "Merged_Data"
                            current_row = 1
                            for idx, file in enumerate(merge_files):
                                file_bytes = file.getvalue()
                                src_wb = load_workbook(filename=BytesIO(file_bytes), data_only=False)
                                src_ws = src_wb.active
                                if idx == 0:
                                    for row in src_ws.iter_rows(min_row=1, max_row=1):
                                        for cell in row:
                                            dst = merged_ws.cell(current_row, cell.column, cell.value)
                                            if cell.has_style:
                                                try:
                                                    if cell.font: dst.font = cell.font
                                                    if cell.fill and cell.fill.fill_type: dst.fill = cell.fill
                                                    if cell.border: dst.border = cell.border
                                                    if cell.alignment: dst.alignment = cell.alignment
                                                    dst.number_format = cell.number_format
                                                except Exception: pass
                                    current_row += 1
                                for row in src_ws.iter_rows(min_row=2):
                                    for cell in row:
                                        dst = merged_ws.cell(current_row, cell.column, cell.value)
                                        if cell.has_style:
                                            try:
                                                if cell.font: dst.font = cell.font
                                                if cell.fill and cell.fill.fill_type: dst.fill = cell.fill
                                                if cell.border: dst.border = cell.border
                                                if cell.alignment: dst.alignment = cell.alignment
                                                dst.number_format = cell.number_format
                                            except Exception: pass
                                    current_row += 1
                                try:
                                    for col_letter in src_ws.column_dimensions:
                                        width = src_ws.column_dimensions[col_letter].width
                                        if width:
                                            merged_ws.column_dimensions[col_letter].width = width
                                except Exception: pass
                            out = BytesIO(); merged_wb.save(out); out.seek(0)
                            st.success("✅ تم الدمج مع الحفاظ على التنسيق")
                            st.download_button("⬇️ تحميل الملف المدمج", out.getvalue(), file_name="Merged_Consolidated_Formatted.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        else:
                            all_dfs = []
                            for file in merge_files:
                                ext = file.name.split(".")[-1].lower()
                                df = pd.read_csv(file) if ext == "csv" else pd.read_excel(file)
                                all_dfs.append(df)
                            merged_df = pd.concat(all_dfs, ignore_index=True)
                            out = BytesIO(); merged_df.to_excel(out, index=False, engine='openpyxl'); out.seek(0)
                            st.success("✅ تم الدمج (تنسيق CSV قد لا يُحفظ)")
                            st.download_button("⬇️ تحميل الملف المدمج", out.getvalue(), file_name="Merged_Consolidated.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    except Exception as e:
                        st.error(f"❌ خطأ أثناء الدمج: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

# ===================== Excel Processor Card =====================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 🧰 معالج Excel (إعادة تسمية/حذف/ترتيب أعمدة)")
    st.markdown('<div class="hint">مطابق لإعداداتك: إعادة تسمية (MR/DM/AM) + حذف أعمدة محددة + ترتيب نهائي</div>', unsafe_allow_html=True)

    proc_file = st.file_uploader(
        "📂 ارفع ملف Excel للمعالجة (xlsx/xlsm)",
        type=["xlsx", "xlsm"],
        accept_multiple_files=False,
        key=f"processor_uploader_{st.session_state.clear_counter}",
    )

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

    COLUMN_RENAME_MAP = { "L1 Emp Name": "MR", "L2 Emp Name": "DM", "L3 Emp Name": "AM" }

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
                header_row = 1
                headers = [ws.cell(header_row, col).value for col in range(1, ws.max_column + 1)]
                header_to_idx = {h: i+1 for i, h in enumerate(headers) if h is not None}

                # Rename
                for old_name, new_name in COLUMN_RENAME_MAP.items():
                    if old_name in header_to_idx:
                        cidx = header_to_idx[old_name]
                        ws.cell(header_row, cidx).value = new_name
                        headers[cidx-1] = new_name
                        header_to_idx[new_name] = cidx
                        del header_to_idx[old_name]

                # Delete
                to_delete_indices = [header_to_idx[c] for c in COLUMNS_TO_DELETE if c in header_to_idx]
                to_delete_indices.sort(reverse=True)
                for cidx in to_delete_indices:
                    ws.delete_cols(cidx)

                # Recompute
                headers = [ws.cell(header_row, col).value for col in range(1, ws.max_column + 1)]
                header_to_idx = {h: i+1 for i, h in enumerate(headers) if h is not None}

                # Reorder into new workbook
                col_mapping = [header_to_idx.get(name) for name in FINAL_COLUMN_ORDER]
                new_wb = Workbook(); new_ws = new_wb.active
                for row_idx in range(1, ws.max_row + 1):
                    for new_col_idx, old_col_idx in enumerate(col_mapping, start=1):
                        if old_col_idx is None: continue
                        old_cell = ws.cell(row_idx, old_col_idx)
                        new_cell = new_ws.cell(row_idx, new_col_idx)
                        new_cell.value = old_cell.value
                        if old_cell.has_style:
                            try:
                                if old_cell.font: new_cell.font = old_cell.font
                                if old_cell.fill and old_cell.fill.fill_type: new_cell.fill = old_cell.fill
                                if old_cell.border: new_cell.border = old_cell.border
                                if old_cell.alignment: new_cell.alignment = old_cell.alignment
                                new_cell.number_format = old_cell.number_format
                            except Exception: pass
                for c in range(1, len(FINAL_COLUMN_ORDER) + 1):
                    new_ws.column_dimensions[get_column_letter(c)].width = 15

                out_buf = BytesIO(); new_wb.save(out_buf); out_buf.seek(0)
                st.success("✅ تم إتمام المعالجة بنجاح")
                base = os.path.splitext(proc_file.name)[0]
                st.download_button("⬇️ تحميل الملف المعالج", out_buf.getvalue(), file_name=f"{_safe_name(base)}_processed.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"❌ خطأ أثناء المعالجة: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

# ===================== Images → PDF Card =====================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 🖼️ تحويل الصور إلى PDF (جودة أصلية)")
    st.markdown('<div class="hint">ارفع صورة أو أكثر وسيتم دمجها في ملف PDF واحد</div>', unsafe_allow_html=True)

    uploaded_images = st.file_uploader(
        "📂 ارفع صور JPG/PNG",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key=f"image_uploader_{st.session_state.clear_counter}",
    )

    if uploaded_images:
        display_uploaded_files(uploaded_images, "Images")
        c1, c2 = st.columns([1,1])
        with c1:
            if st.button("🧹 مسح الصور", key="clear_images"):
                st.session_state.clear_counter += 1
                st.rerun()
        with c2:
            if st.button("🖨️ إنشاء PDF (جودة أصلية)"):
                with st.spinner("جاري إنشاء ملف PDF..."):
                    try:
                        first = Image.open(uploaded_images[0]).convert("RGB")
                        others = [Image.open(x).convert("RGB") for x in uploaded_images[1:]]
                        pdf_buffer = BytesIO(); first.save(pdf_buffer, format="PDF", save_all=True, append_images=others); pdf_buffer.seek(0)
                        st.success("✅ تم إنشاء ملف PDF بنجاح")
                        st.download_button("⬇️ تحميل ملف PDF", pdf_buffer.getvalue(), file_name="Images_Combined.pdf", mime="application/pdf")
                    except Exception as e:
                        st.error(f"❌ خطأ أثناء إنشاء PDF: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

# Footer
st.markdown("<hr>", unsafe_allow_html=True)
st.caption("© Tricks For Excel — تواصل: WhatsApp 01554694554")
