import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import io, re, zipfile

# Streamlit page setup
st.set_page_config(page_title="Certificate Generator", page_icon="🎓", layout="wide")
st.title("Automated Certificate Generator")

# File Uploads
excel_file = st.file_uploader("Upload Participant List (Excel)", type=["xlsx"])
pdf_file = st.file_uploader("Upload Certificate Template (PDF)", type=["pdf"])

# Settings Panel 
st.markdown("Certificate Text Settings")

col1, col2 = st.columns(2)

with col1:
    student_font_size = st.number_input("Student Name Font Size", 10, 60, 18)
    student_x = st.number_input("Student X Position", 0, 1224, 306)
    student_y = st.number_input("Student Y Position", 400, 800, 610)

with col2:
    school_font_size = st.number_input("School Name Font Size", 10, 60, 18)
    school_x = st.number_input("School X Position", 0, 1224, 306)
    school_y = st.number_input("School Y Position", 400, 800, 550)

# Main Logic 
if excel_file and pdf_file:
    participants = pd.read_excel(excel_file, header = 1 )
    pdf_bytes = pdf_file.read()  # Read template into memory once

    # Auto-detect column names
    student_col = next((c for c in participants.columns if "student" in c.lower() or "name" in c.lower()), participants.columns[0])
    school_col = next((c for c in participants.columns if "school" in c.lower() or "institution" in c.lower()), 
                      participants.columns[1] if len(participants.columns) > 1 else participants.columns[0])

    st.markdown(f" Using column {student_col} for Student names.")
    st.markdown(f" Using column {school_col} for School names.")
    st.success(f"  Loaded {len(participants)} participants.")

    # Generate on Button Click
    if st.button("Generate Certificates"):
        zip_buf = io.BytesIO()
        success, fail = 0, 0

        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zipf:
            for idx, row in participants.iterrows():
                try:
                    student = str(row[student_col]).strip()
                    school = str(row[school_col]).strip() if not pd.isna(row[school_col]) else ""

                    # Skip if student name is empty
                    if not student or pd.isna(student):
                        fail += 1
                        continue

                    # Create overlay canvas for each 
                    overlay_packet = io.BytesIO()
                    w, h = letter
                    c = canvas.Canvas(overlay_packet, pagesize=(w, h))
                    c.setFillColorRGB(0, 0, 0)
                    c.setFont("Helvetica-Bold", student_font_size)
                    c.drawCentredString(student_x, student_y, student)
                    c.setFont("Helvetica-Bold", school_font_size)
                    c.drawCentredString(school_x, school_y, school)
                    c.save()
                    overlay_packet.seek(0)

                    # Merge overlay & base certificate template 
                    base_reader = PdfReader(io.BytesIO(pdf_bytes))
                    overlay_reader = PdfReader(overlay_packet)
                    base_page = base_reader.pages[0]
                    base_page.merge_page(overlay_reader.pages[0])

                    # Write merged page to memory buffer
                    out_buf = io.BytesIO()
                    writer = PdfWriter()
                    writer.add_page(base_page)
                    writer.write(out_buf)
                    writer.close()
                    out_buf.seek(0)

                    # Safe unique filename
                    safe_name = re.sub(r'[<>:"/\\|?*]', "_", student.replace(" ", "_"))
                    filename = f"{idx+1:03d}_{safe_name}_certificate.pdf"

                    zipf.writestr(filename, out_buf.getvalue())
                    success += 1

                except Exception as e:
                    fail += 1
                    st.error(f"Error creating certificate for {student or 'Unknown'}: {e}")

        # Finalize and rewind ZIP
        zip_buf.seek(0)

        st.success(f"Completed: {success} successful, {fail} failed.")

        # FIX: Use .getvalue() for stable Streamlit download
        st.download_button(
            "Download All Certificates (.zip)",
            data=zip_buf.getvalue(),
            file_name="certificates.zip",
            mime="application/zip"
        )

else:
    st.info("Please upload both the Excel file and Certificate template to begin.")
