![baner](https://github.com/Ghosts6/Local-website/blob/main/img/Baner.png)

# 💻FormatFlex

today im gonna review python porgrams  that i write with help of python and its model to convert pdf to word and etc,
first we hvae pdf_to_wordv1 which is use aspose.words model to convert pdf file in this method we add require def as a 
convert for simplify our work and with help of that at first we load pdf file then we save it as word,also we provide
convert.py which can convert more then one format in fact it can support foramt type like xml pdf word and txt.

# 🖥️Convert excel to xml:

Here we have programs that i write with  python to convert excel fille to xml i also create custome one to create xml file which is use for telephone phonebook
for this work we write three differend program include:Convert_xlsx_to_xlm_v1.py,Convert_xlsx_to_xlm_v2.py and Convert excel to phonebook.xlm.py

```python
import  jpype     
import  asposecells     
jpype.startJVM() 
from asposecells.api import Workbook
workbook = Workbook("phonebook01.xlsx")
workbook.save("phonebook01.xml")
jpype.shutdownJVM()
```
#convert excel to xml v2:

```python
from openpyxl import load_workbook
import xml.etree.ElementTree as ET

def excel_to_xml(excel_file, xml_file):
    
    wb = load_workbook(excel_file)
    sheet = wb.active
    
    # Create the XML 
    root = ET.Element('data')
    
    for row in sheet.iter_rows(values_only=True):
        record = ET.SubElement(root, 'record')
        for value in row:
            ET.SubElement(record, 'item').text = str(value) if value is not None else ''
    
    tree = ET.ElementTree(root)
    
    tree.write(xml_file, encoding='utf-8', xml_declaration=True)

excel_input_file = 'input.xlsx'

xml_output_file = 'output.xml'


excel_to_xml(excel_input_file, xml_output_file)
```
#convert phonebook.xlsx to xml

```python
import pandas as pd
import xml.etree.ElementTree as ET

# Read the Excel 
excel_file = 'your_excel_file.xlsx'  # Replace with your Excel file name
data = pd.read_excel(excel_file)

# Define updated default values
default_values = {
    'line': '0',
    'ring': 'Auto',
    'group_id_name': 'all contacts',
    'default_photo': 'Default:default_contact_image.png',
    'other_number': '', 
    'auto_divert': ''    
}

# Create the XML structure
root = ET.Element('phonebook')

# Iterate through the Excel data and create XML elements
for _, row in data.iterrows():
    contact = ET.SubElement(root, 'contact')
    for col in ['display_name', 'display_number', 'mobil', 'other_number', 'auto_divert']:
        value = str(row[col]) if not pd.isnull(row[col]) else default_values.get(col, '')
        ET.SubElement(contact, col).text = value

# Create the XML
tree = ET.ElementTree(root)

# Save XML 
xml_output_file = 'phonebook.xml' 
tree.write(xml_output_file, encoding='utf-8', xml_declaration=True)
```


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
