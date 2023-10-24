import requests
import datetime
import zipfile
import io
import base64

class EmailProcessor:
    def __init__(self, client_id, client_secret, tenant_id, email_address):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.email_address = email_address
        self.access_token = self.get_access_token()
        self.headers = {
            "Authorization": f"Bearer {self.access_token}",
        }
        self.inbox_folder_name = "Inbox"
        self.ap_mailbox_subfolder_name = "APMailbox"
        self.tool_pickup_subfolder_name = "ToolPickup"
        self.manual_pickup_subfolder_name = "ManualPickup"

    def log(self, message):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{timestamp}] {message}")

    def get_access_token(self):
        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/token"
        token_data = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "resource": "https://graph.microsoft.com/",
        }

        self.log("Requesting access token...")
        self.log(token_url)

        try:
            token_response = requests.post(token_url, data=token_data)
            token_response.raise_for_status()
            access_token = token_response.json()["access_token"]
            self.log("Access token obtained successfully.")
            return access_token
        except requests.exceptions.RequestException as e:
            self.log(f"Error making token request: {e}")
            exit(1)
        except KeyError as e:
            self.log(f"Error parsing token response: {e}")
            exit(1)

    def get_destination_folder_id(self, folder_name):
        folders_url = f"https://graph.microsoft.com/v1.0/users/{self.email_address}/mailFolders"
        try:
            response = requests.get(folders_url, headers=self.headers)
            response.raise_for_status()
            folders = response.json().get("value", [])
            for folder in folders:
                if folder["displayName"] == folder_name:
                    return folder["id"]
            return None  # Folder not found
        except requests.exceptions.RequestException as e:
            self.log(f"Error fetching mail folders: {e}")
            return None

    def message_exists(self, message_id):
        message_url = f"https://graph.microsoft.com/v1.0/users/{self.email_address}/messages/{message_id}"
        try:
            response = requests.get(message_url, headers=self.headers)
            response.raise_for_status()
            return True  # Message exists
        except requests.exceptions.RequestException as e:
            return False  # Message does not exist

    def process_zip_attachments(self, attachments):
        xml_attachment_count = 0
        zip_attachments_count = 0

        for attachment in attachments:
            content_type = attachment.get("contentType", "")
            filename = attachment.get("name", "").lower()  # Get the filename in lowercase
            self.log(f"Analyzing an attachment with filetype: {content_type}")

            # Check if the file has a ".odf" or ".ODF" extension and skip it
            if filename.endswith((".ofd", ".OFD", ".xlsm", ".XLSM")):
                self.log(f"Skipping a file on the ignore list: {filename}")
                continue

            if content_type.startswith("application/zip") or content_type.startswith(
                    "application/x-zip-compressed") or content_type.startswith("application/octet-stream"):
                zip_attachments_count += 1
                zip_data = attachment.get("contentBytes", b"")
                if isinstance(zip_data, str):
                    zip_data = zip_data.encode("utf-8")
                if zip_data:
                    xml_attachment_count += self.process_nested_zips(io.BytesIO(base64.b64decode(zip_data)))

            elif content_type.startswith("application/xml") or content_type.startswith("text/xml"):
                xml_attachment_count += 1
                self.log("Found XML attachment.")

        self.log(f"Total ZIP attachments: {zip_attachments_count}")
        self.log(f"Total XML attachments (including within ZIPs): {xml_attachment_count}")

        if xml_attachment_count == 1:
            self.log("Analysis Result: Moving to ToolPickup - Only 1 XML file found in the message")
            return self.tool_pickup_subfolder_name
        else:
            self.log("Analysis Result: Moving to ManualPickup")
            return self.manual_pickup_subfolder_name

    def process_nested_zips(self, zip_data):
        xml_count = 0
        try:
            with zipfile.ZipFile(zip_data, "r") as zip_file:
                for zip_info in zip_file.infolist():
                    if zip_info.filename.lower().endswith(".xml"):
                        self.log(f"Found XML file '{zip_info.filename}' within the ZIP attachment.")
                        xml_count += 1
                    elif zip_info.filename.lower().endswith(".zip"):
                        self.log(f"Found nested ZIP file '{zip_info.filename}' within the ZIP attachment.")
                        nested_zip_data = zip_file.read(zip_info.filename)
                        xml_count += self.process_nested_zips(io.BytesIO(nested_zip_data))
        except zipfile.BadZipFile:
            self.log("Error: Not a valid ZIP file. Skipping.")
        except Exception as e:
            self.log(f"Error extracting ZIP file: {e}")
        return xml_count

    def process_email_messages(self):
        folders_url = f"https://graph.microsoft.com/v1.0/users/{self.email_address}/mailFolders/{self.inbox_folder_name}/childFolders"

        self.log(f"Fetching folders inside the '{self.inbox_folder_name}' folder...")

        try:
            response = requests.get(folders_url, headers=self.headers)
            response.raise_for_status()
            folders = response.json().get("value", [])
            folder_names = [folder["displayName"] for folder in folders]
            self.log(f"Folders inside '{self.inbox_folder_name}': {', '.join(folder_names)}")
        except requests.exceptions.RequestException as e:
            self.log(f"Error fetching folders: {e}")
            exit(1)

        self.log("Subfolders within Inbox:")

        for folder in folders:
            self.log(f"- {folder['displayName']}")

        for folder in folders:
            if folder["displayName"] == self.ap_mailbox_subfolder_name:
                ap_mailbox_subfolder_id = folder["id"]
                inbox_messages_url = f"https://graph.microsoft.com/v1.0/users/{self.email_address}/mailFolders/{ap_mailbox_subfolder_id}/messages?$expand=attachments"

                self.log(f"Fetching email messages from the '{self.ap_mailbox_subfolder_name}' subfolder...")

                # Initialize pagination parameters
                page = 1
                page_size = 1000  # Adjust this as needed

                while True:
                    # Add pagination parameters to the URL
                    page_url = f"{inbox_messages_url}&$top={page_size}&$skip={(page - 1) * page_size}"

                    try:
                        response = requests.get(page_url, headers=self.headers)
                        response.raise_for_status()
                        inbox_messages = response.json().get("value", [])
                        self.log(f"Received {len(inbox_messages)} email messages from page {page}.")

                        if not inbox_messages:
                            break  # No more messages to process

                        for message in inbox_messages:
                            attachments = message.get("attachments", [])
                            move_folder = self.process_zip_attachments(attachments)

                            self.log(
                                f"Analyzing message with subject: {message['subject']} from the '{self.ap_mailbox_subfolder_name}' subfolder...")

                            # Verify the message ID
                            message_id = message['id']
                            message_url = f"https://graph.microsoft.com/v1.0/users/{self.email_address}/messages/{message_id}"
                            try:
                                response = requests.get(message_url, headers=self.headers)
                                response.raise_for_status()
                                message_exists = True
                            except requests.exceptions.RequestException as e:
                                message_exists = False

                            if message_exists:
                                self.log(f"Message with subject: {message['subject']} exists. Proceeding to move...")
                                self.log(f"Message ID: {message_id}")

                                # Verify the destination folder ID
                                destination_folder_id = self.get_destination_folder_id(move_folder)
                                if destination_folder_id is not None:
                                    self.move_email_message(message_id, destination_folder_id)
                                else:
                                    self.log(f"Destination folder '{move_folder}' not found.")
                            else:
                                self.log(f"Message with subject: {message['subject']} no longer exists.")

                        page += 1  # Move to the next page
                    except requests.exceptions.RequestException as e:
                        self.log(f"Error fetching email messages: {e}")
                        break  # Stop pagination on error

    def move_email_message(self, message_id, destination_folder_id):
        move_url = f"https://graph.microsoft.com/v1.0/users/{self.email_address}/messages/{message_id}/move"
        move_data = {
            "destinationId": destination_folder_id,
        }

        self.log(f"Moving message with ID {message_id} to '{destination_folder_id}' folder...")

        try:
            move_response = requests.post(move_url, headers=self.headers, json=move_data)
            move_response.raise_for_status()
            self.log(f"Message with ID {message_id} moved successfully to '{destination_folder_id}'.")
        except requests.exceptions.RequestException as e:
            self.log(f"Error moving message with ID {message_id}: {e}")
            self.log(f"Move request failed with {move_response.status_code} status code.")
            self.log(f"Response content: {move_response.text}")

if __name__ == "__main__":
    client_id = "ac6e12bf-fe16-4479-a941-a143e735cfcd"
    client_secret = "t5jsK-_sL-qQD1wcj7h5R~O3-o0Uxc_uWp"
    tenant_id = "a33c6ac4-a52e-45c5-af07-b972df9bd004"
    email_address = "iig.e-invoices@inter.ikea.com"

    processor = EmailProcessor(client_id, client_secret, tenant_id, email_address)
    processor.process_email_messages()
