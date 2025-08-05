import os
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Preformatted
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import comtypes.client

def create_watermark(watermark_text, output_file):
    c = canvas.Canvas(output_file)
    c.setFont("Helvetica", 40) # Font type and font size
    c.setFillColorRGB(0.5, 0.5, 0.5, 0.4)  # Set watermark color (RGB) and transparency
    c.saveState()
    c.translate(100, 200)  # Move the origin to a better position for rotation
    c.rotate(45)  # Rotate the text by 45 degrees
    c.drawString(0, 0, watermark_text)  # Draw text at the new origin
    c.restoreState()
    c.save()

def add_watermark(input_pdf_path, output_pdf_path, watermark_pdf_path):
    # Read the watermark PDF
    watermark_pdf = PdfReader(watermark_pdf_path)
    watermark_page = watermark_pdf.pages[0]

    # Read the input PDF
    input_pdf = PdfReader(input_pdf_path)
    writer = PdfWriter()

    # Add watermark to each page
    for page in input_pdf.pages:
        page.merge_page(watermark_page)
        writer.add_page(page)

    # Write the output PDF
    with open(output_pdf_path, "wb") as output_pdf_file:
        writer.write(output_pdf_file)


def convert_txt_to_pdf(txt_file_path: str, pdf_file_path: str, font_size=10):
    """
    Convert a text file to PDF while preserving original formatting.
    
    Args:
        txt_file_path (str): Path to the input .txt file
        pdf_file_path (str): Path to the output .pdf file
        font_size (int): Font size for the text (default: 10)
    """
    # Create a PDF document
    doc = SimpleDocTemplate(pdf_file_path, pagesize=letter,
                            rightMargin=36, leftMargin=36,
                            topMargin=36, bottomMargin=36)
    
    # Get default styles
    styles = getSampleStyleSheet()
    
    # Create a monospace style that preserves formatting
    preformatted_style = ParagraphStyle(
        'PreformattedText',
        parent=styles['Code'],
        fontName='Courier',  # Monospace font
        fontSize=font_size,
        leftIndent=0,
        rightIndent=0,
        spaceAfter=0,
        spaceBefore=0,
        wordWrap='LTR',  # Left to right, preserve spacing
    )
    
    # Read the text file
    try:
        with open(txt_file_path, 'r', encoding='utf-8') as file:
            content = file.read()
    except FileNotFoundError:
        print(f"Error: File '{txt_file_path}' not found.")
        return False
    except Exception as e:
        print(f"Error reading file: {e}")
        return False
    
    # Create story with preformatted text
    story = []
    
    # Use Preformatted flowable which preserves whitespace and formatting
    preformatted_text = Preformatted(content, preformatted_style)
    story.append(preformatted_text)
    
    # Build the PDF
    try:
        doc.build(story)
        print(f"PDF successfully created: {pdf_file_path}")
        return True
    except Exception as e:
        print(f"Error creating PDF: {e}")
        return False


def convert_docx_to_pdf(docx_file: str, pdf_file: str) -> bool:
    """
    Convert DOCX to PDF using Microsoft Word COM interface.
    This method preserves ALL formatting perfectly including:
    - Complex layouts, headers/footers
    - Images, charts, tables
    - Fonts, styles, colors
    - Page breaks, margins
    - Embedded objects

    Args:
        docx_file: file path of docx file
        pdf_file: file path of pdf file

    Returns:
        True if successful, False otherwise

    """
    comtypes.CoInitialize()

    try:
        # Create Word application object
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False  # Don't show Word window
        
        # Convert to absolute paths
        docx_path = os.path.abspath(docx_file)
        pdf_path = os.path.abspath(pdf_file)
        
        # Open the document
        doc = word.Documents.Open(docx_path)
        
        # Export as PDF with high quality settings
        doc.ExportAsFixedFormat(
            OutputFileName=pdf_path,
            ExportFormat=17,  # PDF format
            OpenAfterExport=False,
            OptimizeFor=0,  # Optimize for print (better quality)
            CreateBookmarks=0  # 0 for None, 1 for Headings
        )
        
        # Close document and Word
        doc.Close()
        word.Quit()
        
        # Verify conversion
        if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
            print(f"✓ Perfect conversion completed: {pdf_file}")
            return True
        else:
            print("✗ PDF conversion failed")
            return False
            
    except Exception as e:
        print(f"✗ Error with Word COM: {e}")
        # Try to cleanup Word process
        try:
            if 'word' in locals():
                word.Quit()
        except:
            pass
        return False
    
    finally:
        # Always uninitialize COM
        comtypes.CoUninitialize()
