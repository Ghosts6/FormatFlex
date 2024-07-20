![baner](https://github.com/Ghosts6/Local-website/blob/main/img/Baner.png)

# 💻FormatFlex

Today, I'm going to review Python programs that I wrote with the help of Python and its libraries to convert files between various formats. The repository includes several conversion tools, such as:

1. **pdf_to_word**: Utilizes the Aspose.Words library to convert PDF files to Word documents. This method simplifies our work by defining a `convert` function that loads a PDF file and saves it as a Word document.

2. **convert_xlsx_to_xml**: Converts Excel files (XLSX format) to XML format, making it easier to work with Excel data in XML-based systems.

3. **convert.py**: A versatile script that can handle multiple format conversions, including XML, PDF, Word, and TXT.

Additionally, we have a dedicated program for converting PSD files to PNG format

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

    pdf_document = fitz.open(pdf_path)

    doc = Document()
    
    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        text = page.get_text("text")

        doc.add_paragraph(text)

    doc.save(word_path)
    print(f"PDF converted to Word: {word_path}")

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
# 🖼️ PSD to PNG Conversion Program

This program converts PSD files to PNG format with ease. It traverses directories, converts all PSD files it finds, and saves the converted PNG files in a specified output directory. The program also includes robust error handling and logging to ensure smooth operation even when encountering problematic files.

## ✨ Features

- 🚀 **Automated Conversion:** Convert all PSD files in a directory and its subdirectories to PNG format.
- 📂 **Directory Handling:** Organizes converted PNG files in the specified output directory, mirroring the original directory structure.
- 🛠️ **Error Handling:** Continues processing remaining files even if some conversions fail.
- 📄 **Logging:** Logs errors and successful conversions to a `log.txt` file in the output directory for easy troubleshooting.

## 🛠️ Requirements

Make sure you have the following Python libraries installed:
- `Pillow`
- `psd_tools`

You can install them using pip:
```sh
pip install pillow psd-tools
```

## 📝 Example Code
```python
from PIL import Image
import psd_tools
import os
import logging

# logging
def setup_logging(output_base_directory):
    log_path = os.path.join(output_base_directory, 'log.txt')
    logging.basicConfig(filename=log_path, level=logging.DEBUG, 
                        format='%(asctime)s - %(levelname)s - %(message)s')

def convert_psd_to_png(directory, output_base_directory):
    for root, _, files in os.walk(directory):
        for filename in files:
            if filename.lower().endswith('.psd'):
                file_path = os.path.join(root, filename)
                try:
                    psd = psd_tools.PSDImage.open(file_path)
                    composite = psd.compose()

                    relative_path = os.path.relpath(root, directory)
                    output_directory = os.path.join(output_base_directory, relative_path)
                    os.makedirs(output_directory, exist_ok=True)

                    png_path = os.path.join(output_directory, os.path.splitext(filename)[0] + '.png')
                    composite.save(png_path)

                    print(f"Converted {filename} to {png_path}")
                    logging.info(f"Successfully converted {file_path} to {png_path}")
                except Exception as e:
                    print(f"Error converting {filename}: {e}")
                    logging.error(f"Failed to convert {file_path}: {e}")

def main():
    while True:
        directory = input("Enter the base directory path containing PSD files: ")
        output_base_directory = os.path.join(directory, 'png')
        os.makedirs(output_base_directory, exist_ok=True)
        
        setup_logging(output_base_directory)

        convert_psd_to_png(directory, output_base_directory)

        continue_choice = input("Do you want to convert files in another directory? (yes/no): ").strip().lower()
        if continue_choice != 'yes':
            break

if __name__ == "__main__":
    main()
```
# 🚀 Usage
Clone the Repository:
```sh
git clone https://github.com/yourusername/psd-to-png-converter.git
cd psd-to-png-converter
```
Run the Program:
```sh
#example
python3 psd_to_png.py
```
Follow the Prompts:

# 📬 Contact
For any questions or suggestions, please open an issue or contact us at kiarash@kiarashbashokian.com 

# Happy converting! 🎉