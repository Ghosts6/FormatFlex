![baner](https://github.com/Ghosts6/Local-website/blob/main/img/Baner.png)

# 💻Convert_pdf_to_wordWithPython
today im gonna review python porgrams  that i write with help of python and its model to convert pdf to word and etc,
first we hvae pdf_to_wordv1 which is use aspose.words model to convert pdf file in this method we add require def as a 
convert for simplify our work and with help of that at first we load pdf file then we save it as word,also we provide
convert.py which can convert more then one format in fact it can support foramt type like xml pdf word and txt.

# 📝pdf_to_wordv1.py
#🚨hint! 

this method use to convert data into image and save it in word file so it keep details but we cant edit it

```python
import aspose.words as convert
# load pdf file
doc = convert.Document("Input.pdf") 
# save loaded file as word file
doc.save("Output.docx
```
# 🔧pdf_to_wordv2.py

for second program we try different way in this method we need to install some python file with help of pip:
1PyMuPDF 2python-docx
unlike first program in this one at first we take data from pdf file then we create word file and write data 
inside it,the important difference is with this method we add data as text not image and we can edit them

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
# 📏Convert_program:

With convert program we can convert different format to each other,only  thing you need some basic lib and some extra like Docx fitz and reportlab after that by running
program and select formats and entering path of input and output we can easliy convert foramts

```python
import fitz
import docx  
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from xml.etree import ElementTree as ET

def pdf_to_word(pdf_path, word_path):
    pdf_document = fitz.open(pdf_path)
    doc = Document()
    
    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        text = page.get_text("text")
        doc.add_paragraph(text)

    doc.save(word_path)
    print(f"PDF converted to Word: {word_path}")

def xml_to_pdf(file_path, file_2_path):
    with open(file_path, 'r') as xml_file:
        xml_data = xml_file.read()
    
    root = ET.fromstring(xml_data)
    
    c = canvas.Canvas(file_2_path, pagesize=letter)
    
    y_offset = 750
    for element in root:
        if element.tag == 'title':
            c.setFont("Helvetica-Bold", 16)
            c.drawString(100, y_offset, element.text)
            y_offset -= 20
            
        elif element.tag == 'paragraph':
            c.setFont("Helvetica", 12)
            c.drawString(100, y_offset, element.text)
            y_offset -= 15
            
    c.save()
    print(f"XML converted to PDF: {file_2_path}")

def xml_to_pdf(file_path, file_2_path):
    with open(file_path, 'r') as xml_file:
        xml_data = xml_file.read()
    
    root = ET.fromstring(xml_data)
    
    c = canvas.Canvas(file_2_path, pagesize=letter)
    
    y_offset = 750
    for element in root:
        if element.tag == 'title':
            c.setFont("Helvetica-Bold", 16)
            c.drawString(100, y_offset, element.text)
            y_offset -= 20
            
        elif element.tag == 'paragraph':
            c.setFont("Helvetica", 12)
            c.drawString(100, y_offset, element.text)
            y_offset -= 15
            
    c.save()
    print(f"XML converted to PDF: {file_2_path}")

def xml_to_word(file_path, file_2_path):
    with open(file_path, 'r') as xml_file:
        xml_data = xml_file.read()
    
    root = ET.fromstring(xml_data)
    
    doc = Document()
    
    for element in root:
        if element.tag == 'title':
            doc.add_heading(element.text, level=1)
            
        elif element.tag == 'paragraph':
            doc.add_paragraph(element.text)
    
    doc.save(file_2_path)
    print(f"XML converted to Word: {file_2_path}")

def xml_to_txt(file_path, file_2_path):
    with open(file_path, 'r') as xml_file:
        xml_data = xml_file.read()
    
    root = ET.fromstring(xml_data)
    
    with open(file_2_path, 'w') as txt_file:
        for element in root:
            txt_file.write(element.text + '\n')
    
    print(f"XML converted to TXT: {file_2_path}")

def pdf_to_xml(pdf_path, xml_path):
    pdf_document = fitz.open(pdf_path)
    
    root = ET.Element("document")
    
    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        text = page.get_text()
        
        page_element = ET.SubElement(root, "page")
        
        page_element.text = text
    
    tree = ET.ElementTree(root)
    tree.write(xml_path)
    
    print(f"PDF converted to XML: {xml_path}")

def pdf_to_txt(pdf_path, txt_path):
    pdf_document = fitz.open(pdf_path)
    text = ""
    
    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        text += page.get_text()
    
    with open(txt_path, 'w') as txt_file:
        txt_file.write(text)
    
    print(f"PDF converted to TXT: {txt_path}")

def word_to_xml(word_path, xml_path):
    doc = docx.Document(word_path)
    
    root = ET.Element("document")
    
    for paragraph in doc.paragraphs:
        paragraph_element = ET.SubElement(root, "paragraph")

        paragraph_element.text = paragraph.text

    tree = ET.ElementTree(root)
    tree.write(xml_path)
    
    print(f"Word converted to XML: {xml_path}")

def word_to_pdf(file_path, file_2_path):
    doc = docx.Document(file_path)
    c = canvas.Canvas(file_2_path, pagesize=letter)
    
    for paragraph in doc.paragraphs:
        c.drawString(100, 750, paragraph.text)
        c.showPage()
    
    c.save()
    
    print(f"Word converted to PDF: {file_2_path}")

def word_to_txt(file_path, file_2_path):
    doc = docx.Document(file_path)
    
    with open(file_2_path, 'w') as txt_file:
        for paragraph in doc.paragraphs:
            txt_file.write(paragraph.text + '\n')
    
    print(f"Word converted to TXT: {file_2_path}")

def txt_to_xml(file_path, file_2_path):
    with open(file_path, 'r') as txt_file:
        txt_data = txt_file.read()
    
    root = ET.Element("document")
    
    for line in txt_data.split('\n'):
        paragraph_element = ET.SubElement(root, "paragraph")
        paragraph_element.text = line
    
    tree = ET.ElementTree(root)
    tree.write(file_2_path)
    
    print(f"TXT converted to XML: {file_2_path}")

def txt_to_pdf(file_path, file_2_path):
    with open(file_path, 'r') as txt_file:
        txt_data = txt_file.read()
    
    c = canvas.Canvas(file_2_path, pagesize=letter)
    c.drawString(100, 750, txt_data)
    c.save()
    
    print(f"TXT converted to PDF: {file_2_path}")

def txt_to_word(file_path, file_2_path):
    doc = docx.Document()
    
    with open(file_path, 'r') as txt_file:
        txt_data = txt_file.read()
    
    for line in txt_data.split('\n'):
        doc.add_paragraph(line)
    
    doc.save(file_2_path)
    
    print(f"TXT converted to Word: {file_2_path}")

print("---Welcome to our conversion program---")
choice = input("Select your input file type (xml/pdf/word/txt): ")
convert = input("Select the output file type (xml/pdf/word/txt): ")
choice_path = input("Please enter the path of the input file: ")
convert_path = input("Please enter the path of the output file: ")

if choice == "xml":
    if convert == "pdf":
        xml_to_pdf(choice_path, convert_path)
    elif convert == "word":
        xml_to_word(choice_path, convert_path)
    elif convert == "txt":
        xml_to_txt(choice_path, convert_path)
    else:
        print("Wrong type or same type as input file, please try again.")

elif choice == "pdf":
    if convert == "xml":
        pdf_to_xml(choice_path, convert_path)
    elif convert == "word":
        pdf_to_word(choice_path, convert_path)
    elif convert == "txt":
        pdf_to_txt(choice_path, convert_path)
    else:
        print("Wrong type or same type as input file, please try again.")

elif choice == "word":
    if convert == "pdf":
        word_to_pdf(choice_path, convert_path)
    elif convert == "xml":
        word_to_xml(choice_path, convert_path)
    elif convert == "txt":
        word_to_txt(choice_path, convert_path)
    else:
        print("Wrong type or same type as input file, please try again.")

elif choice == "txt":
    if convert == "xml":
        txt_to_xml(choice_path, convert_path)
    elif convert == "pdf":
        txt_to_pdf(choice_path, convert_path)
    elif convert == "word":
        txt_to_word(choice_path, convert_path)
    else:
        print("Wrong type or same type as input file, please try again.")
else:
    print("Invalid input file type. Please select xml/pdf/word/txt.")
```
