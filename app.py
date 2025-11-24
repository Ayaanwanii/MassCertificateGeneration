import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter, PageObject
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io, re, zipfile
import base64 
import copy # <-- Import the standard copy module

# Register font (ensure this font file exists in your working folder)
pdfmetrics.registerFont(TTFont('BlissExtraBold', './Bliss Extra Bold.ttf'))
pdfmetrics.registerFont(TTFont('Alliance-BoldItalic', './alliance-bolditalic.ttf'))

try:
    pdfmetrics.registerFont(TTFont('FrutigerLight', './FrutigerLight-TrueType.ttf'))
except Exception:
    pass 

# Streamlit page setup
st.set_page_config(page_title="Certificate Generator", page_icon="🎓", layout="wide")
st.title("Automated Certificate Generator")

# File uploads
excel_file = st.file_uploader("Upload Participant List (Excel)", type=["xlsx"])
pdf_file = st.file_uploader("Upload Certificate Template (PDF)", type=["pdf"])

# --- Utility Functions ---

def hex_to_rgb(hex_color):
    """Convert hex color to RGB tuple (0-1 range for ReportLab)."""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) / 255 for i in (0, 2, 4))

def generate_certificate_pdf(row, student_col, school_col, pdf_bytes, settings):
    """
    Core logic to generate a single certificate PDF buffer.
    """
    
    # Extract data
    student = str(row[student_col]).strip() if student_col in row and pd.notna(row[student_col]) else ""
    school = ""
    if school_col is not None and school_col in row:
        raw_school = row[school_col]
        if pd.notna(raw_school) and str(raw_school).strip() != "":
            school = str(raw_school).strip()
            
    if not student:
        raise ValueError("Student name is empty or invalid.")

    # Read base template
    base_reader = PdfReader(io.BytesIO(pdf_bytes))
    if not base_reader.pages:
        raise ValueError("PDF template has no pages.")
        
    base_page = base_reader.pages[0]
    media_box = base_page.mediabox
    w = float(media_box.width)
    h = float(media_box.height)

    # Create overlay (text layer)
    overlay_packet = io.BytesIO()
    c = canvas.Canvas(overlay_packet, pagesize=(w, h))
    
    # Draw Student Name
    c.setFillColorRGB(*hex_to_rgb(settings['student_color']))
    c.setFont(settings['student_font'], settings['student_font_size'])
    c.drawCentredString(settings['student_x'], settings['student_y'], student)

    # Draw School Name
    if school:
        c.setFillColorRGB(*hex_to_rgb(settings['school_color']))
        c.setFont(settings['school_font'], settings['school_font_size'])
        c.drawCentredString(settings['school_x'], settings['school_y'], school)

    c.save()
    overlay_packet.seek(0)

    # Merge base + overlay
    overlay_reader = PdfReader(overlay_packet)
    
    # Use deepcopy to ensure merged_page is a PageObject
    merged_page = copy.deepcopy(base_page) 
    merged_page.merge_page(overlay_reader.pages[0])
    
    # Write output
    out_buf = io.BytesIO()
    writer = PdfWriter()
    writer.add_page(merged_page)
    writer.write(out_buf)
    
    writer.close() 
    
    out_buf.seek(0) # Seek to start so it can be read later
    
    return out_buf, student


# --- Settings Panel ---
st.markdown("### Certificate Text Settings")

col1, col2 = st.columns(2)

# Prepare font list dynamically
available_fonts = [
    "Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique",
    "Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic",
    "Courier", "Courier-Bold", "Courier-Oblique", "Courier-BoldOblique",
    "Symbol", "ZapfDingbats", 
    "BlissExtraBold", 
    "Alliance-BoldItalic",
    "FrutigerLight",
]

with col1:
    student_font_size = st.number_input("Student Name Font Size", 10, 60, 18)
    student_x = st.slider("Student X Position", 0, 1224, 427)
    student_y = st.slider("Student Y Position", 0, 800, 200)
    student_font = st.selectbox("Student Font", available_fonts, index=available_fonts.index("Helvetica-Bold"))
    student_color = st.color_picker("Student Text Color", "#000000")

with col2:
    school_font_size = st.number_input("School Name Font Size", 10, 60, 18)
    school_x = st.slider("School X Position", 0, 1224, 306)
    school_y = st.slider("School Y Position", 0, 800, 550)
    school_font = st.selectbox("School Font", available_fonts, index=available_fonts.index("Helvetica-Bold"))
    school_color = st.color_picker("School Text Color", "#000000")

# Collect all settings into a dictionary for easy passing
settings = {
    'student_font_size': student_font_size,
    'student_x': student_x,
    'student_y': student_y,
    'student_font': student_font,
    'student_color': student_color,
    'school_font_size': school_font_size,
    'school_x': school_x,
    'school_y': school_y,
    'school_font': school_font,
    'school_color': school_color,
}


# --- Main Logic ---
if excel_file and pdf_file:
    pdf_bytes = pdf_file.read()
    
    # -----------------------------------------------------------------
    # CRITICAL FIX: DYNAMIC HEADER DETECTION
    # -----------------------------------------------------------------
    
    # Reset file pointer for reading headers/skip detection
    excel_file.seek(0)
    
    try:
        # Attempt to read the first few rows (as headers) to check for title/empty rows
        temp_df = pd.read_excel(excel_file, header=None, nrows=5) 
        
        # Check for the known pattern of the "Javier Test Certificate" file
        skip_rows_count = 0
        if temp_df.iloc[0, 0] is not None and 'test certificate' in str(temp_df.iloc[0, 0]).lower():
            skip_rows_count = 2

    except Exception:
        skip_rows_count = 0 # Default to standard reading
        
    excel_file.seek(0) # Reset pointer again for final reading
    
    # Read the final DataFrame, setting the header to the correct row index
    participants = pd.read_excel(excel_file, header=skip_rows_count) 
    
    # Handle unnamed first column
    if participants.columns.size > 0 and (participants.columns[0] == participants.columns.name or str(participants.columns[0]).startswith('Unnamed:')):
         participants.columns = ["Student Name"] + list(participants.columns[1:])
    
    # Auto-detect columns
    student_col = next(
        (c for c in participants.columns if "student" in str(c).lower() or "name" in str(c).lower()),
        participants.columns[0] if participants.columns.size > 0 else None
    )

    school_col = next(
        (c for c in participants.columns if "school" in str(c).lower() or "institution" in str(c).lower()),
        None if participants.columns.size <= 1 else participants.columns[1]
    )

    if student_col:
        st.markdown(f"Using column **`{student_col}`** for Student names.")
        if school_col:
             st.markdown(f"Using column **`{school_col}`** for School names.")
        else:
             st.warning("No dedicated 'School' column found. Only student names will be inserted.")
        
        st.success(f" Loaded {len(participants)} participants.")
    else:
        st.error("Could not find any columns in the Excel file.")

    st.markdown("---")
    
    # --- PREVIEW FEATURE (Download Link) ---
    st.subheader("1. Preview Certificate")
    
    # NEW CONTAINER FOR THE DOWNLOAD BUTTON
    preview_placeholder = st.empty() 

    # Initialize session state for preview visibility
    if 'show_preview' not in st.session_state:
        st.session_state.show_preview = False

    if st.button("Generate Preview"): 
        st.session_state.show_preview = True

    if st.session_state.show_preview:
        if participants.empty or student_col is None:
            st.error("The participant list is empty or the column structure is invalid.")
        else:
            try:
                # Decide which row to preview (second row if available, otherwise first)
                if len(participants) < 2:
                    st.warning("The participant list has fewer than two entries. Previewing the first entry (Index 0).")
                    first_row = participants.iloc[0]
                else:
                    first_row = participants.iloc[1] 
                
                with st.spinner(f"Generating preview for {first_row[student_col]}..."):
                    
                    # Generate the single PDF using the core function
                    preview_buf, student_name = generate_certificate_pdf(
                        first_row, student_col, school_col, pdf_bytes, settings
                    )
                    
                    # Display the generated PDF using an iframe
                    base64_pdf = base64.b64encode(preview_buf.getvalue()).decode('utf-8')
                    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="600px" type="application/pdf"></iframe>'
                    preview_placeholder.markdown(pdf_display, unsafe_allow_html=True)
                    
                    st.info("The preview is displayed above. Adjust settings (sliders, etc.) to see real-time updates.")
                    
            except Exception as e:
                st.error(f"Error creating preview. Please check your X/Y coordinates, font selection, and uploaded files. Error: {e}")


    # --- BATCH GENERATION FEATURE ---
    st.subheader("2. Generate and Download All")
    if st.button("Generate All Certificates"):
        if participants.empty or student_col is None:
            st.error("The participant list is empty or the column structure is invalid.")
        else:
            zip_buf = io.BytesIO()
            success, fail = 0, 0
            
            with st.spinner("Generating all PDFs and zipping..."):
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zipf:
                    for idx, row in participants.iterrows():
                        unique_id = f"Row_{idx+1}"
                        student = ""
                        try:
                            # Use the reusable function here too
                            out_buf, student = generate_certificate_pdf(
                                row, student_col, school_col, pdf_bytes, settings
                            )

                            # Safe filename
                            safe_name = re.sub(r'[<>:"/\\|?*]', "_", student.replace(" ", "_"))
                            filename = f"{idx+1:03d}_{safe_name}_certificate.pdf"

                            zipf.writestr(filename, out_buf.getvalue())
                            success += 1

                        except ValueError as ve:
                            fail += 1
                            if "empty" in str(ve).lower():
                                st.warning(f"Skipping row {idx+1}: Student name is empty.")
                            else:
                                st.error(f"Error in row {idx+1} ({student or unique_id}): {ve}")
                        except Exception as e:
                            fail += 1
                            st.error(f"Critical error creating certificate for **{student or unique_id}**: {e}")

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
    st.info("Please upload both the Excel file and Certificate template to begin.")
