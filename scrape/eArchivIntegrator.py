import os
import xml.etree.ElementTree as ET
import pandas as pd

def extract_node_value(element, tag):
    node = element.find(tag)
    if node is not None and node.text:
        return node.text.strip()
    return ''

def format_date(date_string):
    # Assuming the input date format is YYYY-MM-DD HH:MM:SS
    parts = date_string.split(' ')[0].split('-')
    return '/'.join([parts[1], parts[2], parts[0]])

def convert_xml_to_excel(folder_path, output_file):
    xml_files = [file for file in os.listdir(folder_path) if file.endswith('.xml')]

    if not xml_files:
        print("No XML files found in the specified folder.")
        return

    data = []
    for file in xml_files:
        file_path = os.path.join(folder_path, file)
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()

            values = root.find('__values__')
            uuid = extract_node_value(values, 'Uuid')
            createddatetime = extract_node_value(values, 'IssueDate')
            invoice_number = extract_node_value(values, 'InvoiceNumber')
            sender_identifier = extract_node_value(values, 'SenderIdentifier')
            sender_title = extract_node_value(values, 'SenderCommercialTitle')
            tax_total = extract_node_value(values, 'TaxTotal')
            payable_amount = extract_node_value(values, 'PayableAmount')
            tax_exclusive_amount = extract_node_value(values, 'TaxExclusiveAmount')

            # Format the date to MM/DD/YYYY
            formatted_createddatetime = format_date(createddatetime)

            data.append({
                'UUID': uuid,
                'CreatedDateTime': formatted_createddatetime,
                'InvoiceNumber': invoice_number,
                'SenderIdentifier': sender_identifier,
                'SenderCommercialTitle': sender_title,
                'TaxTotal': tax_total,
                'PayableAmount': payable_amount,
                'TaxExclusiveAmount': tax_exclusive_amount
            })
        except (ET.ParseError, FileNotFoundError):
            print(f"Error reading the XML file: {file_path}")
            continue

    if not data:
        print("No valid data found in the XML files.")
        return

    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)
    print(f"Conversion completed. Excel file saved as {output_file}.")

# Provide the folder path containing the XML files and the output Excel file name
folder_path = r'C:\Users\mahoo8\PycharmProjects\RydooUserDownload\UploadTR'
output_file = 'upload.xlsx'

convert_xml_to_excel(folder_path, output_file)
