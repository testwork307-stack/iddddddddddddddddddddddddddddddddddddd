import streamlit as st
import pandas as pd
import os
import zipfile
from PIL import Image
import tempfile

# ====== دالة تفك الضغط ======
def extract_zip(uploaded_file):
    tmpdir = tempfile.mkdtemp()
    with zipfile.ZipFile(uploaded_file, "r") as zip_ref:
        zip_ref.extractall(tmpdir)

    # Debug: عرض الملفات المستخرجة
    st.write("=== Debug: Extracted files structure ===")
    for dirpath, _, filenames in os.walk(tmpdir):
        st.write("Folder:", dirpath, "->", filenames[:20])  # أول 20 ملف فقط

    return tmpdir

# ====== دالة البحث عن الصورة ======
def find_photo_path(base_dir, photo_filename):
    for dirpath, _, filenames in os.walk(base_dir):
        for fn in filenames:
            if fn.lower() == photo_filename.lower():  # تطابق كامل مع الاسم والامتداد
                return os.path.join(dirpath, fn)
    return None

# ====== Streamlit App ======
st.title("📸 Photo Finder from Excel + ZIP")

# رفع الإكسل
excel_file = st.file_uploader("Upload Excel file", type=["xlsx"])
# رفع الصور (ZIP)
zip_file = st.file_uploader("Upload Photos ZIP", type=["zip"])

if excel_file and zip_file:
    # قراءة البيانات من Excel
    df = pd.read_excel(excel_file)
    st.write("### Preview of Excel Data", df.head())

    # فك الصور
    tmpdir = extract_zip(zip_file)

    # محاولة التعرف على أعمدة الاسم والصورة
    name_col = None
    photo_col = None
    for col in df.columns:
        if str(col).strip().lower() in ["name", "الاسم"]:
            name_col = col
        if str(col).strip().lower() in ["photo", "الصورة", "photos", "صورة"]:
            photo_col = col

    if name_col and photo_col:
        for _, row in df.iterrows():
            name = str(row[name_col])
            photo_filename = str(row[photo_col]).strip()

            photo_path = find_photo_path(tmpdir, photo_filename)
            if photo_path:
                st.image(photo_path, caption=f"{name} ({photo_filename})", width=200)
            else:
                st.warning(f"📷 Photo not found for '{name}'. Requested: {photo_filename}")
    else:
        st.error(f"❌ Excel must contain a 'Name/الاسم' column and a 'Photo/الصورة' column. Found: {list(df.columns)}")
