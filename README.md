# Convert_pdf_to_wordWithPython
today im gonna review two python code that i write with help of python and its model to convert pdf to word
first we hvae pdf_to_wordv1 which is use aspose.words model to convert pdf file in this method we add require def as a convert for simplify our work and with help of that at first we load pdf file then we save it as word

# Hint !
this method use to convert data into image and save it in word file so it keep details but we cant edit it

#pdf_to_wordv1.py
```python
import aspose.words as convert
# load pdf file
doc = convert.Document("Input.pdf") 
# save loaded file as word file
doc.save("Output.docx
```

for second program we try different way in this method we need to install some python file with help of pip:
1PyMuPDF 2python-docx
unlike first program in this one at first we take data from pdf file then we create word file and write data 
inside it,the important difference is with this method we add data as text not image and we can edit them

#pdf_to_wordv2.py
```python
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
```
