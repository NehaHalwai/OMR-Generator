import os
import re
import base64
import pandas as pd
import streamlit as st
from io import BytesIO
from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Table, TableStyle
import zipfile

# ===== PATH CONFIG =====
BASE_DIR = Path(__file__).parent

child_omr_template = str(BASE_DIR / "child_omr.jpg")   # relative path
master_omr_template = str(BASE_DIR / "master_omr.jpg") # relative path
LOGO_FILE = str(BASE_DIR / "logo.webp")                # relative path

# ===== Master OMR Bubble positions =====
MASTER_ROLL_X_CM = [10.1, 11.5, 12.9]
MASTER_BUBBLE_Y_TOP_CM = [22, 22, 22]
MASTER_BUBBLE_SPACING_CM = 0.62
MASTER_BUBBLE_RADIUS_CM = 0.24

# ===== Child OMR Bubble positions =====
SHIFT_X_CM = 6
SHIFT_Y_CM = -0.3
CHILD_ROLL_X_CM = [9.9 + SHIFT_X_CM, 11.3 + SHIFT_X_CM, 12.6 + SHIFT_X_CM]
CHILD_BUBBLE_Y_TOP_CM = [21.5 + SHIFT_Y_CM, 21.5 + SHIFT_Y_CM, 21.5 + SHIFT_Y_CM]
CHILD_BUBBLE_SPACING_CM = 0.61
CHILD_BUBBLE_RADIUS_CM = 0.23


# ----- Utility Functions -----
def normalize_col_name(s):
    return re.sub(r'[^a-z0-9]', '', str(s).lower().strip()) if s else ""


def find_column(df_cols_norm, aliases):
    for orig_col, norm in df_cols_norm.items():
        for a in aliases:
            if norm == a:
                return orig_col
    for orig_col, norm in df_cols_norm.items():
        for a in aliases:
            if a in norm:
                return orig_col
    return None


def safe_filename(s):
    s = str(s).strip()
    s = re.sub(r'[\\/*?:"<>|]', '_', s)
    s = re.sub(r'\s+', '_', s)
    return s[:200]


def format_roll_value(v):
    if pd.isna(v) or not str(v).strip():
        return "000"
    try:
        return str(int(float(v))).zfill(3)
    except ValueError:
        return str(v).zfill(3)[:3]


# ----- Roll Filling Functions -----
def fill_roll_bubbles_master(c, roll_no):
    roll_no = roll_no.zfill(3)
    for i, digit_char in enumerate(roll_no):
        if not digit_char.isdigit():
            continue
        digit = int(digit_char)
        x = (MASTER_ROLL_X_CM[i] * cm) / 2.2
        y = MASTER_BUBBLE_Y_TOP_CM[i] * cm - digit * MASTER_BUBBLE_SPACING_CM * cm - (2.6 * cm) + 0.03 * cm
        c.setFillColor(colors.black)
        c.circle(x, y, MASTER_BUBBLE_RADIUS_CM * cm, stroke=0, fill=1)


def fill_roll_bubbles_child(c, roll_no):
    roll_no = roll_no.zfill(3)
    for i, digit_char in enumerate(roll_no):
        if not digit_char.isdigit():
            continue
        digit = int(digit_char)
        x = (CHILD_ROLL_X_CM[i] * cm) / 2.2
        y = CHILD_BUBBLE_Y_TOP_CM[i] * cm - digit * CHILD_BUBBLE_SPACING_CM * cm - (2.6 * cm) + 0.2 * cm
        c.setFillColor(colors.black)
        c.circle(x, y, CHILD_BUBBLE_RADIUS_CM * cm, stroke=0, fill=1)


def draw_roll_number_text(c, roll_no, template="master"):
    roll_no = roll_no.zfill(3)
    if template == "master":
        text_y = MASTER_BUBBLE_Y_TOP_CM[0] * cm - (2.1 * cm)
        x_positions = MASTER_ROLL_X_CM
    else:
        text_y = CHILD_BUBBLE_Y_TOP_CM[0] * cm - (2.0 * cm)
        x_positions = CHILD_ROLL_X_CM

    c.setFont("Helvetica-Bold", 14)
    c.setFillColor(colors.black)
    for i, digit_char in enumerate(roll_no):
        x = (x_positions[i] * cm) / 2.2
        c.drawCentredString(x, text_y, digit_char)


# ----- Class Parsing -----
def parse_class_value(class_val):
    if pd.isna(class_val):
        return None
    s = str(class_val).strip().lower()

    if s.isdigit():
        return int(s)

    match = re.search(r'(\d+)', s)
    if match:
        return int(match.group(1))

    roman_map = {
        "i": 1, "ii": 2, "iii": 3, "iv": 4, "v": 5,
        "vi": 6, "vii": 7, "viii": 8, "ix": 9, "x": 10,
        "xi": 11, "xii": 12
    }
    if s in roman_map:
        return roman_map[s]

    words_map = {
        "first": 1, "second": 2, "third": 3, "fourth": 4, "fifth": 5,
        "sixth": 6, "seventh": 7, "eighth": 8, "ninth": 9, "tenth": 10,
        "eleventh": 11, "twelfth": 12
    }
    if s in words_map:
        return words_map[s]

    return None


# ===== Streamlit App =====
if Path(LOGO_FILE).exists():
    with open(LOGO_FILE, "rb") as f:
        data = f.read()
    img_base64 = base64.b64encode(data).decode()
    st.markdown(
        f'<div style="text-align:center;"><img src="data:image/webp;base64,{img_base64}" width="150"></div>',
        unsafe_allow_html=True
    )
else:
    st.image("https://placehold.co/150x50/3498db/ffffff?text=LOGO+Missing", width=150)
    st.warning(f"Note: Local logo file '{LOGO_FILE}' not found. Using placeholder.")

st.markdown(
    "<h1 style='text-align: center; color: white;'>OMR Sheet Generator</h1>",
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    # Ensure template files exist
    if not Path(child_omr_template).exists() or not Path(master_omr_template).exists():
        st.error("❌ OMR template files are missing. Please ensure child_omr.jpg and master_omr.jpg exist in the repo.")
        st.stop()

    with st.spinner("⏳ Please wait, generating PDFs..."):
        xls = pd.ExcelFile(uploaded_file)
        output_zip = BytesIO()

        with zipfile.ZipFile(output_zip, "w") as zipf:
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name, dtype=object)
                df_cols_norm = {orig: normalize_col_name(orig) for orig in df.columns}
                aliases = {
                    "school_name": ["schoolname", "scoolname", "school"],
                    "class": ["class", "grade", "standard"],
                    "division": ["division", "section"],
                    "roll_no": ["rollno", "rollnumber", "roll_no"],
                    "student_name": ["nameofthestudent", "name", "studentname"],
                }
                col_map = {canon: find_column(df_cols_norm, al_list) for canon, al_list in aliases.items()}

                pdf_buffer = BytesIO()
                c = canvas.Canvas(pdf_buffer, pagesize=A4)
                width, height = A4

                for _, row in df.iterrows():
                    student_name = row.get(col_map["student_name"], "") if col_map["student_name"] else ""
                    school_name = row.get(col_map["school_name"], "") if col_map["school_name"] else ""
                    class_name = row.get(col_map["class"], "") if col_map["class"] else ""
                    division = row.get(col_map["division"], "") if col_map["division"] else ""
                    roll_no_raw = row.get(col_map["roll_no"], "") if col_map["roll_no"] else ""
                    roll_no = format_roll_value(roll_no_raw)

                    parsed_class = parse_class_value(class_name)
                    if parsed_class is not None and parsed_class in [1, 2, 3]:
                        omr_template = child_omr_template
                        template_type = "child"
                    else:
                        omr_template = master_omr_template
                        template_type = "master"

                    try:
                        omr_img = ImageReader(omr_template)
                        c.drawImage(omr_img, 0, 0, width=width, height=height, preserveAspectRatio=True)
                    except Exception as e:
                        st.error(f"❌ Failed to load OMR template image: {omr_template} → {e}")
                        st.stop()

                    if template_type == "child":
                        fill_roll_bubbles_child(c, roll_no)
                    else:
                        fill_roll_bubbles_master(c, roll_no)

                    draw_roll_number_text(c, roll_no, template=template_type)

                    data = [
                        [f"Student Name: {student_name or ' '}"],
                        [f"School: {school_name or ' '}"],
                        [f"Class: {class_name or ' '}      Division: {division or ' '}"],
                        ["Question Paper Set: _____________"],
                    ]
                    table_width = width * 0.7
                    table = Table(data, colWidths=[table_width])
                    table.setStyle(TableStyle([
                        ("BOX", (0,0), (-1,-1), 0.8, colors.black),
                        ("INNERGRID", (0,0), (-1,-1), 0.5, colors.black),
                        ("FONTNAME", (0,0), (-1,-1), "Helvetica-Bold"),
                        ("FONTSIZE", (0,0), (-1,-1), 11),
                        ("ALIGN", (0,0), (-1,-1), "LEFT"),
                        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                        ("LEFTPADDING", (0,0), (-1,-1), 10),
                        ("RIGHTPADDING", (0,0), (-1,-1), 10),
                        ("TOPPADDING", (0,0), (-1,-1), 5),
                        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
                    ]))
                    w, h = table.wrap(0,0)
                    x = (width - w)/2
                    y = height - 4.5*cm - h
                    table.drawOn(c, x, y)

                    c.showPage()

                c.save()
                pdf_data = pdf_buffer.getvalue()
                pdf_buffer.close()
                pdf_filename = f"{safe_filename(sheet_name)}.pdf"
                zipf.writestr(pdf_filename, pdf_data)

    st.success(" PDFs Generated Successfully!")
    st.download_button(
        label="⬇ Download All PDFs (ZIP)",
        data=output_zip.getvalue(),
        file_name="Generated_OMRs.zip",
        mime="application/zip"
    )
