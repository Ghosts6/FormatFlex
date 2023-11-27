#requirement:
#PyMuPDF
#python-docx
import fitz  
from docx import Document

def pdf_to_word(pdf_path, word_path):
    # Open the PDF file
    pdf_document = fitz.open(pdf_path)
    
    # Create word file
    doc = Document()
    
    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        text = page.get_text("text")
        
        # Write data in word file
        doc.add_paragraph(text)
    
    # Save file
    doc.save(word_path)
    print(f"PDF converted to Word: {word_path}")

# Replace 'input_file.pdf' and 'output_file.docx' with your file paths
pdf_to_word('input_file.pdf', 'output_file.docx')
