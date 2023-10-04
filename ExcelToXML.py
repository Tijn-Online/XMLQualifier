import pandas as pd
import xml.etree.ElementTree as ET

# Read the Excel file
df = pd.read_excel(r'C:\Users\mahoo8\Downloads\Results.xlsx')

# Create the root element for the XML
root = ET.Element('DocumentElement')

# Iterate through rows in the DataFrame and convert to XML
for _, row in df.iterrows():
    response = ET.SubElement(root, 'Response')

    company = ET.SubElement(response, 'Company')
    company.text = str(row['Company'])

    voucher_number = ET.SubElement(response, 'VoucherNumber')
    voucher_number.text = str(row['VoucherNumber'])

    document_type = ET.SubElement(response, 'DocumentType')
    document_type.text = str(row['DocumentType'])

    payment_date = ET.SubElement(response, 'PaymentDate')
    payment_date.text = str(row['PaymentDate'])

    payment_reference = ET.SubElement(response, 'PaymentReference')
    payment_reference.text = str(row['PaymentReference'])

# Create the XML tree
tree = ET.ElementTree(root)

# Save the XML to a file
tree.write('output.xml', encoding='utf-8', xml_declaration=True)
