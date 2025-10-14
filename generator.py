import os
from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import reportlab.rl_config
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io
import pandas as pd
import re  # For better filename sanitization

# Ensure output directory exists
os.makedirs('Certificates', exist_ok=True)

# Read data once
participants = pd.read_excel('Input/participants.xlsx')
print(f"Loaded {len(participants)} participants.")  # Debug: Confirm total

# Register fonts ONCE outside the loop (must have .ttf files in working directory)
reportlab.rl_config.warnOnMissingFontGlyphs = 0
#pdfmetrics.registerFont(TTFont('Helvetica-Bold', 'Helvetica-Bold.ttf'))  # Bold for student
#pdfmetrics.registerFont(TTFont('Helvetica-Bold', 'Helvetica-Bold.ttf'))  # Regular for school

# Read template PDF ONCE outside the loop
#template_pdf = PdfReader(open("Input/certificate_template.pdf", "rb"))
#template_page = template_pdf.pages[0]  # Assume single page; clone later

total = len(participants)
successful = 0
failed = 0

for i in range(total):
    try:
        # Extract data
        student = str(participants.loc[i, 'Student']).strip()
        school = str(participants.loc[i, 'School']).strip()
        
        if pd.isna(participants.loc[i, 'Student']) or not student:  # Skip empty names
            print(f"Skipping row {i+1}: Empty student name.")
            continue
        
        print(f"Processing {i+1}/{total}: {student}")  # Debug: Track progress (remove if too verbose)

        # Create overlay canvas (doubled size for positioning)
        packet = io.BytesIO()
        width, height = letter
        c = canvas.Canvas(packet, pagesize=(width * 2, height * 2))
        
        # Student name (bold, large, gold)
        c.setFillColorRGB(0, 0, 0)
        c.setFont('Helvetica-Bold', 18)
        c.drawCentredString(306, 610, student)
        
        # School (regular, smaller, gold)
        c.setFillColorRGB(0, 0, 0)
        c.setFont('Helvetica-Bold', 18)
        c.drawCentredString(306, 550, school)
        
        c.save()
        packet.seek(0)

        packet.seek(0)

        with open("Input/certificate_template.pdf", "rb") as f:
          template_pdf = PdfReader(f)
          page = template_pdf.pages[0]  # always fresh

          overlay_pdf = PdfReader(packet)
          page.merge_page(overlay_pdf.pages[0])

          safe_name = re.sub(r'[<>:"/\\|?*]', '_', student.replace(" ", "_"))  # Sanitize filename
          certificado = f"Certificates/{safe_name}_certificate.pdf"            # File path to save PDF

          count = 1
          while os.path.exists(certificado):
             certificado = f"Certificates/{safe_name}_{count}_certificate.pdf"
             count += 1


          output = PdfWriter()
          output.add_page(page)
          with open(certificado, "wb") as outputStream:
              output.write(outputStream)

        
          packet.close()  # Cleanup
          successful += 1
        
    except KeyError as e:
        print(f"Error in row {i+1}: Missing column - {e}")
        failed += 1
    except FileNotFoundError as e:
        print(f"Error in row {i+1}: File issue - {e}")
        failed += 1
    except Exception as e:
        print(f"Unexpected error in row {i+1} ({student}): {e}")
        failed += 1

print(f"Completed: {successful} successful, {failed} failed out of {total}.")
