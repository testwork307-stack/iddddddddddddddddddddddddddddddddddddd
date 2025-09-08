# -*- coding: utf-8 -*-
"""
Streamlit – HR ID Generator (ZIP only, Arabic-aware, debug for missing photos)
"""

import os, io, zipfile, tempfile, shutil
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from barcode import Code128
from barcode.writer import ImageWriter

# Arabic text handling
import arabic_reshaper
from bidi.algorithm import get_display

# ===================== CONFIG =====================
PHOTO_POS = (111, 168)
PHOTO_SIZE = (300, 300)
BARCODE_POS = (570, 465)
BARCODE_SIZE = (390, 120)

NAME_OFFSET_X = -40
NAME_OFFSET_Y = -20

# ===================== UI =========================
st.set_page_config(page_title="HR ID Card Generator", page_icon="🎫", layout="wide")
st.title("🎫 HR ID Card Generator")

with st.sidebar:
    st.markdown("**Tips**")
    st.markdown("- Excel columns: **الاسم**, **الوظيفة**, **الرقم**, **الرقم القومي**, **الصورة**.")
    st.markdown("- Photos must be in a **ZIP** file.")
    st.markdown("- For best Arabic rendering, upload proper TTF fonts.")

font_ar_file = st.sidebar.file_uploader("Arabic font (TTF/OTF)", type=["ttf", "otf"])
font_en_file = st.sidebar.file_uploader("English font (TTF/OTF)", type=["ttf", "otf"])

excel_file = st.file_uploader("📂 Upload Excel (.xlsx)", type=["xlsx"])
photos_archive = st.file_uploader("📦 Upload Photos (ZIP)", type=["zip"])
template_file = st.file_uploader("🖼 Upload Card Template (PNG/JPG)", type=["png", "jpg", "jpeg"])

# ================== Helpers =======================
def load_font_from_upload(upload, fallback_name: str, size: int):
    """Load a font from upload or fallback."""
    if upload is not None:
        try:
            return ImageFont.truetype(io.BytesIO(upload.read()), size)
        except Exception:
            st.warning(f"⚠️ Failed to load uploaded font for {fallback_name}. Using default.")
    for candidate in [
        "Amiri-Regular.ttf", "NotoNaskhArabic-Regular.ttf",
        "Arial.ttf", "Tahoma.ttf"
    ]:
        try:
            return ImageFont.truetype(candidate, size)
        except Exception:
            continue
    return ImageFont.load_default()

def prepare_text(text: str) -> str:
    """Arabic reshape + bidi."""
    if not text:
        return ""
    reshaped = arabic_reshaper.reshape(str(text))
    return get_display(reshaped)

def draw_aligned_text(draw, xy, text, font, fill="black", anchor="rt"):
    """Draw right-to-left aligned text."""
    if not text:
        return
    draw.text(xy, text, font=font, fill=fill, anchor=anchor)

def draw_bold_text(draw, xy, text, font, fill="black", anchor="rt"):
    """Fake bold by multiple draws."""
    for dx, dy in [(0,0), (1,0), (0,1), (1,1)]:
        draw_aligned_text(draw, (xy[0]+dx, xy[1]+dy), text, font, fill, anchor)

def find_photo_path(root_dir: str, requested: str):
    """Find photo by name in extracted ZIP."""
    if not requested:
        return None
    req_stem = Path(requested).stem.lower()
    for dirpath, _, filenames in os.walk(root_dir):
        for fn in filenames:
            if Path(fn).stem.lower() == req_stem:
                return os.path.join(dirpath, fn)
    return None

# ================== Main logic ====================
if excel_file and photos_archive and template_file:
    # Load fonts
    font_ar = load_font_from_upload(font_ar_file, "Arabic", 36)
    font_en = load_font_from_upload(font_en_file, "English", 30)

    # Read Excel
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"❌ Failed to read Excel: {e}")
        st.stop()

    # Show preview of excel
    st.subheader("Preview of Excel Data")
    st.dataframe(df.head())

    # Read template
    try:
        template = Image.open(template_file).convert("RGB")
    except Exception as e:
        st.error(f"❌ Failed to read template image: {e}")
        st.stop()

    tmpdir = tempfile.mkdtemp(prefix="idcards_")
    archive_path = os.path.join(tmpdir, photos_archive.name)
    with open(archive_path, "wb") as f:
        f.write(photos_archive.getbuffer())

    # Extract ZIP
    try:
        with zipfile.ZipFile(archive_path, "r") as zf:
            zf.extractall(tmpdir)
    except Exception as e:
        st.error(f"❌ Failed to extract archive: {e}")
        shutil.rmtree(tmpdir, ignore_errors=True)
        st.stop()

    # === Debug: show folder contents ===
    st.subheader("=== Debug: Extracted files structure ===")
    for dirpath, _, filenames in os.walk(tmpdir):
        rel = os.path.relpath(dirpath, tmpdir)
        st.write(f"Folder: {dirpath} →")
        st.write(filenames)

    # === Check for missing photos ===
    available_photos = set()
    for dirpath, _, filenames in os.walk(tmpdir):
        available_photos.update([fn.lower() for fn in filenames])

    requested_photos = set(str(x).strip().lower() for x in df["الصورة"].dropna())
    missing = requested_photos - available_photos
    if missing:
        st.error(f"❌ Missing {len(missing)} photos from ZIP")
        st.write(sorted(list(missing))[:50])  # show first 50 missing

    # === Generate cards ===
    output_cards = []
    progress = st.progress(0)
    status = st.empty()

    for idx, row in df.iterrows():
        status.info(f"Processing {idx+1}/{len(df)} – {row.get('الاسم', '')}")
        card = template.copy()
        draw = ImageDraw.Draw(card)

        # Prepare texts
        name = prepare_text(str(row.get("الاسم", "")).strip())
        job = prepare_text(str(row.get("الوظيفة", "")).strip())
        num = str(row.get("الرقم", "")).strip()
        national_id = str(row.get("الرقم القومي", "")).strip()
        photo_filename = str(row.get("الصورة", "")).strip()

        # Draw texts
        base_name_xy = (915, 240)
        name_xy = (base_name_xy[0] + NAME_OFFSET_X, base_name_xy[1] + NAME_OFFSET_Y)
        draw_bold_text(draw, name_xy, name, font_ar, "black", "rt")

        name_bbox = draw.textbbox((0, 0), name, font=font_ar)
        name_height = (name_bbox[3] - name_bbox[1]) + 20

        job_xy = (name_xy[0], name_xy[1] + name_height)
        draw_aligned_text(draw, job_xy, job, font_ar, "black", "rt")

        job_bbox = draw.textbbox((0, 0), job, font=font_ar)
        job_height = (job_bbox[3] - job_bbox[1]) + 25

        id_xy = (name_xy[0], job_xy[1] + job_height)
        job_id_label = prepare_text(f"الرقم الوظيفي: {num}")
        draw_aligned_text(draw, id_xy, job_id_label, font_ar, "black", "rt")

        # Place photo
        photo_path = find_photo_path(tmpdir, photo_filename)
        if photo_path and os.path.exists(photo_path):
            try:
                img = Image.open(photo_path).convert("RGB").resize(PHOTO_SIZE)
                card.paste(img, PHOTO_POS)
            except Exception as e:
                st.warning(f"⚠️ Failed to place photo for '{row.get('الاسم', '')}': {e}")
        else:
            st.warning(f"📷 Photo not found for '{row.get('الاسم', '')}'. Requested: {photo_filename}")

        # Place barcode
        try:
            if national_id:
                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_barcode:
                    out_noext = tmp_barcode.name[:-4]
                    barcode = Code128(national_id, writer=ImageWriter())
                    barcode_path = barcode.save(out_noext, {"write_text": False})
                    with Image.open(barcode_path) as bimg:
                        bimg = bimg.convert("RGB").resize(BARCODE_SIZE)
                        card.paste(bimg, BARCODE_POS)
                    os.remove(barcode_path)
        except Exception as e:
            st.warning(f"⚠️ Failed to generate barcode for '{row.get('الاسم', '')}': {e}")

        output_cards.append(card)
        progress.progress(int(((idx + 1) / max(len(df), 1)) * 100))

    status.empty()

    # Export PDF
    if output_cards:
        try:
            pdf_path = os.path.join(tmpdir, "All_ID_Cards.pdf")
            output_cards[0].save(pdf_path, save_all=True, append_images=output_cards[1:])
            with open(pdf_path, "rb") as f:
                st.download_button("⬇️ Download All ID Cards (PDF)", f, file_name="All_ID_Cards.pdf")
            st.success(f"✅ Generated {len(output_cards)} cards")
            st.image(output_cards[0], caption="Preview", width=320)
        except Exception as e:
            st.error(f"❌ Failed to write PDF: {e}")
    else:
        st.warning("No cards generated.")
else:
    st.info("👆 Upload Excel, Photos (ZIP), and Template image to start.")
