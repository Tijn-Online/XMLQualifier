# Martijn Hoogendijk 20/07/2023 PYTHON 3.10

import os
import pandas as pd
from datetime import datetime, timedelta
import math


class RydooFileGenerator:
    def __init__(self, excel_file_path, xml_template_file_path):
        self.excel_file_path = excel_file_path
        self.xml_template_file_path = xml_template_file_path

    def read_excel(self):
        self.excel_file = pd.read_excel(self.excel_file_path, converters={'CURRENT_DATE': pd.to_datetime})

    def read_xml_template(self):
        with open(self.xml_template_file_path, 'r') as xml_template_file:
            self.xml_template = xml_template_file.read()

    def excel_date_to_string(self, excel_date):
        excel_date_int = int(excel_date)  # Convert to standard integer
        base_date = datetime(1899, 12, 30)  # Excel's base date
        date = base_date + timedelta(days=excel_date_int)
        formatted_date = date.strftime('%Y-%m-%dT%H:%M:%S')
        return formatted_date

    def row_to_xml(self, row):
        xml_template_row = self.xml_template
        for col in row.index:
            placeholder = f"@{col}"
            if col == "VALUEDATE":
                date_str = self.excel_date_to_string(row[col])
                print(row[col],date_str)
                xml_template_row = xml_template_row.replace(placeholder, date_str)
            else:
                xml_template_row = xml_template_row.replace(placeholder, str(row[col]))
        return xml_template_row

    def perform_replacements_in_xml(self, xml_data):
        xml_data = xml_data.replace("2.0TYPE", "analysis")
        xml_data = xml_data.replace("2TYPE", "analysis")
        xml_data = xml_data.replace("2.0SENSE", "debit")
        xml_data = xml_data.replace("2SENSE", "debit")
        xml_data = xml_data.replace("<tran:Number>1.0</tran:Number>", "<tran:Number>1</tran:Number>")
        xml_data = xml_data.replace("<tran:Number>2.0</tran:Number>", "<tran:Number>2</tran:Number>")
        return xml_data

    def generate_xml_files(self):
        column_names_without_at = [col.replace('@', '') for col in self.excel_file.columns]
        self.excel_file.columns = column_names_without_at

        if not os.path.exists("airpluscleanup"):
            os.makedirs("airpluscleanup")

        batch_size = 255
        total_files = len(self.excel_file)
        num_batches = math.ceil(total_files / batch_size)

        for batch in range(num_batches):
            batch_folder_name = f"airpluscleanup2/batch_{batch + 1}"
            if not os.path.exists(batch_folder_name):
                os.makedirs(batch_folder_name)

            start_index = batch * batch_size
            end_index = min((batch + 1) * batch_size, total_files)

            for index in range(start_index, end_index):
                row = self.excel_file.loc[index]
                xml_data = self.row_to_xml(row)
                document_id = row['DOCUMENTID']
                xml_file_name = f"{batch_folder_name}/TransferBooking_{document_id}.xml"

                xml_data = self.perform_replacements_in_xml(xml_data)

                with open(xml_file_name, 'w') as xml_file:
                    xml_file.write(xml_data)

if __name__ == "__main__":
    sourcedata = "DataLoad.xlsx"
    xmltemplate = "xml_template.xml"
    rydoo_file_generator = RydooFileGenerator(sourcedata, xmltemplate)
    rydoo_file_generator.read_excel()
    rydoo_file_generator.read_xml_template()
    rydoo_file_generator.generate_xml_files()
