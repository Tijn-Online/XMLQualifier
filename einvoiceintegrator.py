from zeep import Client
import xml.etree.ElementTree as ET


class InvoiceProcessor:
    def __init__(self, username, password, begin_date, end_date):
        self.username = username
        self.password = password
        self.begin_date = begin_date
        self.end_date = end_date
        self.auth_key = None
        self.invoice_client = None

    def get_auth_key(self):
        wsdl_url_auth = 'https://integration.visionplus.com.tr/vpefwebservice/authservice.asmx?wsdl'
        client = Client(wsdl_url_auth)
        response_token = client.service.GetAuthorizationKey(self.username, self.password)
        self.auth_key = response_token

    def search_invoices(self):
        wsdl_url_integrator = 'https://integration.visionplus.com.tr/vpefwebservice/einvoice/einvoiceintegrator.asmx?wsdl'
        self.invoice_client = Client(wsdl_url_integrator)
        daterange = {
            'AuthKey': self.auth_key,
            'BeginDate': self.begin_date,
            'EndDate': self.end_date,
        }
        invoices = self.invoice_client.service.SearchInvoiceIncome(**daterange)
        return invoices

    def save_pdf(self, file_content, filename):
        with open(filename, 'wb') as file:
            file.write(file_content)
            print("PDF file saved successfully.")

    def convert_to_xml(self, element, data):
        if isinstance(data, dict):
            for key, value in data.items():
                sub_element = ET.SubElement(element, key)
                self.convert_to_xml(sub_element, value)
        elif hasattr(data, '__dict__'):
            for key, value in data.__dict__.items():
                sub_element = ET.SubElement(element, key)
                self.convert_to_xml(sub_element, value)
        else:
            element.text = str(data)

    def save_xml(self, xml_content, filename):
        with open(filename, 'w', encoding='utf-8') as file:
            file.write(xml_content)
            print("XML file saved successfully.")

    def process_invoices(self):
        self.get_auth_key()
        invoices = self.search_invoices()
        uuid_list = [item['Uuid'] for item in invoices['IncomeResults']['EInvoiceIncomeDetail']]
        print(len(uuid_list), "invoices to process")
        for uuid in uuid_list:
            pdfload = {
                'AuthKey': self.auth_key,
                'UUID': uuid
            }
            pdffile = self.invoice_client.service.GetInvoiceIncomePDF(**pdfload)
            file_content = pdffile['File']
            is_completed = pdffile['isCompleted']
            if is_completed:
                pdf_filename = f'{uuid}.pdf'
                self.save_pdf(file_content, pdf_filename)

                xml_data = next(
                    item for item in invoices['IncomeResults']['EInvoiceIncomeDetail'] if item['Uuid'] == uuid)
                root_element = ET.Element('Root')
                self.convert_to_xml(root_element, xml_data)
                xml_content = ET.tostring(root_element, encoding='unicode')

                xml_filename = f'{uuid}.xml'
                self.save_xml(xml_content, xml_filename)
            else:
                print("Conversion and saving skipped as 'isCompleted' is False.")

    def run(self):
        self.process_invoices()


# Usage
username = 'WS847798'
password = '21196426'
begin_date = '2023-07-19T00:00:00'
end_date = '2023-07-20T23:59:59'

invoice_processor = InvoiceProcessor(username, password, begin_date, end_date)
invoice_processor.run()
