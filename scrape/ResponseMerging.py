import os
import pandas as pd
import xml.etree.ElementTree as ET

# Function to extract data from XML files and return as a list of dictionaries
def extract_data_from_xml(file_path):
    data = []
    tree = ET.parse(file_path)
    root = tree.getroot()

    for response in root.findall(".//Response"):
        print(response)
        if response.find("ErrorMessage") is not None:
            doc_id = response.find("docid").text
            error_message = response.find("ErrorMessage").text
            data.append({"DOCID": doc_id, "ErrorMessage": error_message})

    return data

# Function to process all XML files in the directory and its subdirectories
def process_xml_files(directory_path, excel_file_name):
    data_list = []
    for root, _, files in os.walk(directory_path):
        for file in files:
            if file.endswith(".bak"):
                file_path = os.path.join(root, file)
                print(file_path)
                data_list.extend(extract_data_from_xml(file_path))

    df = pd.DataFrame(data_list)
    df.to_excel(excel_file_name, index=False)

subfolder_path = "payloadResponses"
excel_file_name = "errorresponses.xlsx"

process_xml_files(subfolder_path, excel_file_name)
