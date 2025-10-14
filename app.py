import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import reportlab.rl_config
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io
import re
import zipfile

st.title("Automated Certificate Generator")

# File uploaders for participant Excel and certificate template PDF
excel_file = st.file_uploader("Upload Participant List (Excel)", type=["xlsx"])
pdf_file = st.file_uploader("Upload Certificate Template (PDF)", type=["pdf"])

# Optional customization area
st.markdown("Certificate Text Settings")
student_font_size = st.number_input("Student Name Font Size", value=18, min_value=10, max_value=60)
school_font_size = st.number_input("School Name Font Size", value=18, min_value=10, max_value=60)
student_y = st.number_input("Student Y Position", value=610, min_value=400, max_value=800)
school_y = st.number_input("School Y Position", value=550, min_value=400, max_value=800)
student_x = st.number_input("Student X Position", value=306, min_value=0, max_value=1224)
school_x = st.number_input("School X Position", value=306, min_value=0, max_value=1224)


if excel_file and pdf_file:
    participants = pd.read_excel(excel_file)
    template_pdf = PdfReader(pdf_file)

    # Feedback for loaded participants
    st.write(f"Loaded {len(participants)} participants.")

    # Prepare storage for ZIP
    zip_buf = io.BytesIO()
    success_count = 0
    fail_count = 0

    with zipfile.ZipFile(zip_buf, "w") as zipf:
        for i in range(len(participants)):
            student = "Unknown"
            try:
                student = str(participants.loc[i, 'Student']).strip()
                school = str(participants.loc[i, 'School']).strip()
                if pd.isna(participants.loc[i, 'Student']) or not student:
                    fail_count += 1
                    continue

                # Create overlay canvas
                packet = io.BytesIO()
                width, height = letter
                c = canvas.Canvas(packet, pagesize=(width * 2, height * 2))
                c.setFillColorRGB(0, 0, 0)
                c.setFont('Helvetica-Bold', student_font_size)
                c.drawCentredString(student_x, student_y, student)
                c.setFont('Helvetica-Bold', school_font_size)
                c.drawCentredString(school_x, school_y, school)
                c.save()
                packet.seek(0)

                # Readers and merge
                template_pdf = PdfReader(pdf_file)
                page = template_pdf.pages[0]
                overlay_pdf = PdfReader(packet)
                page.merge_page(overlay_pdf.pages[0])

                # Filename sanitizing
                safe_name = re.sub(r'[<>:"/\\|?*]', '_', student.replace(" ", "_"))
                cert_name = f"{safe_name}_{i+1}_certificate.pdf"

                # Write PDF to buffer then add to ZIP
                out_buf = io.BytesIO()
                writer = PdfWriter()
                writer.add_page(page)
                writer.write(out_buf)
                zipf.writestr(cert_name, out_buf.getvalue())
                packet.close()
                success_count += 1

            except Exception as e:
                fail_count += 1
                st.error(f"Error for {student}: {e}")

    st.write(f"Completed: {success_count} successful, {fail_count} failed.")

    st.download_button(
        "Download All Certificates (.zip)",
        data=zip_buf.getvalue(),
        file_name="certificates.zip",
        mime="application/zip"
    )
else:
    st.info("Please upload both the Excel and template PDF files to begin.")
