import pandas as pd
import xml.etree.ElementTree as ET

# Read the Excel file into a pandas DataFrame
excel_file = 'your_excel_file.xlsx'  # Replace with your Excel file name
data = pd.read_excel(excel_file)

# Define updated default values
default_values = {
    'line': '0',
    'ring': 'Auto',
    'group_id_name': 'all contacts',
    'default_photo': 'Default:default_contact_image.png',
    'other_number': '',  # Default value for 'other_number' when null
    'auto_divert': ''    # Default value for 'auto_divert' when null
}

# Create the XML structure
root = ET.Element('phonebook')

# Iterate through the Excel data and create XML elements
for _, row in data.iterrows():
    contact = ET.SubElement(root, 'contact')
    for col in ['display_name', 'display_number', 'mobil', 'other_number', 'auto_divert']:
        value = str(row[col]) if not pd.isnull(row[col]) else default_values.get(col, '')
        ET.SubElement(contact, col).text = value

# Create the XML tree
tree = ET.ElementTree(root)

# Save XML tree to a file
xml_output_file = 'phonebook.xml'  # Replace with desired XML output file name
tree.write(xml_output_file, encoding='utf-8', xml_declaration=True)
