import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter, PageObject
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io, re, zipfile

# Register font (ensure this font file exists in your working folder)
pdfmetrics.registerFont(TTFont('BlissExtraBold', './Bliss Extra Bold.ttf'))

# Streamlit page setup
st.set_page_config(page_title="Certificate Generator", page_icon="🎓", layout="wide")
st.title("Automated Certificate Generator")

# File uploads
excel_file = st.file_uploader("Upload Participant List (Excel)", type=["xlsx"])
pdf_file = st.file_uploader("Upload Certificate Template (PDF)", type=["pdf"])

# Settings panel
st.markdown("### Certificate Text Settings")

col1, col2 = st.columns(2)

with col1:
    student_font_size = st.number_input("Student Name Font Size", 10, 60, 18)
    student_x = st.number_input("Student X Position", 0, 1224, 427)
    student_y = st.number_input("Student Y Position", 0, 800, 200)

with col2:
    school_font_size = st.number_input("School Name Font Size", 10, 60, 18)
    school_x = st.number_input("School X Position", 0, 1224, 306)
    school_y = st.number_input("School Y Position", 0, 800, 550)

# Main logic
if excel_file and pdf_file:
    # Read Excel and ensure headers are correct
    participants = pd.read_excel(excel_file, header=0)
    if participants.columns[0] == "" or participants.columns[0] is None:
        participants.columns = ["Student Name"]

    pdf_bytes = pdf_file.read()

    # Auto-detect columns
    student_col = next(
        (c for c in participants.columns if "student" in c.lower() or "name" in c.lower()),
        participants.columns[0]
    )

    school_col = next(
        (c for c in participants.columns if "school" in c.lower() or "institution" in c.lower()),
        None if len(participants.columns) == 1 else participants.columns[1]
    )

    st.markdown(f"Using column {student_col} for Student names.")
    st.markdown(f"Using column {school_col} for School names.")
    st.success(f" Loaded {len(participants)} participants.")

    # Generate button
    if st.button("Generate Certificates"):
        zip_buf = io.BytesIO()
        success, fail = 0, 0

        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zipf:
            for idx, row in participants.iterrows():
                try:
                    # Validate student name
                    raw_student = row[student_col]
                    if pd.isna(raw_student) or str(raw_student).strip() == "":
                        fail += 1
                        continue
                    student = str(raw_student).strip()

                    # Optional school name
                    school = ""
                    if school_col is not None:
                        raw_school = row[school_col]
                        if not pd.isna(raw_school) and str(raw_school).strip() != "":
                            school = str(raw_school).strip()

                    # Read base template
                    base_reader = PdfReader(io.BytesIO(pdf_bytes))
                    base_page = base_reader.pages[0]
                    media_box = base_page.mediabox
                    w = float(media_box.width)
                    h = float(media_box.height)

                    # Create overlay
                    overlay_packet = io.BytesIO()
                    c = canvas.Canvas(overlay_packet, pagesize=(w, h))
                    c.setFillColorRGB(18/255, 97/255, 160/255)
                    c.setFont("BlissExtraBold", student_font_size)
                    c.drawCentredString(student_x, student_y, student)

                    if school:
                        c.setFont("BlissExtraBold", school_font_size)
                        c.drawCentredString(school_x, school_y, school)

                    c.save()
                    overlay_packet.seek(0)

                    # Merge base + overlay correctly
                    overlay_reader = PdfReader(overlay_packet)
                    merged_page = PageObject.create_blank_page(
                        width=w, height=h
                    )
                    merged_page.merge_page(base_page)
                    merged_page.merge_page(overlay_reader.pages[0])

                    # Write output
                    out_buf = io.BytesIO()
                    writer = PdfWriter()
                    writer.add_page(merged_page)
                    writer.write(out_buf)
                    writer.close()
                    out_buf.seek(0)

                    # Safe filename
                    safe_name = re.sub(r'[<>:"/\\|?*]', "_", student.replace(" ", "_"))
                    filename = f"{idx+1:03d}_{safe_name}_certificate.pdf"

                    zipf.writestr(filename, out_buf.getvalue())
                    success += 1

                except Exception as e:
                    fail += 1
                    st.error(f"Error creating certificate for {student or 'Unknown'}: {e}")

        # Finalize and download
        zip_buf.seek(0)
        st.success(f"Completed: {success} successful, {fail} failed.")

        st.download_button(
            "Download All Certificates (.zip)",
            data=zip_buf.getvalue(),
            file_name="certificates.zip",
            mime="application/zip"
        )

else:
    st.info("📂 Please upload both the Excel file and Certificate template to begin.")
