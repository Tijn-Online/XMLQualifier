import requests
import json
import datetime
import os


# Define a function to log messages with timestamps
def log(message):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")


# Define your application's credentials
client_id = "ac6e12bf-fe16-4479-a941-a143e735cfcd"
client_secret = "t5jsK-_sL-qQD1wcj7h5R~O3-o0Uxc_uWp"
tenant_id = "a33c6ac4-a52e-45c5-af07-b972df9bd004"
resource_url = "https://graph.microsoft.com/"
email_address = "iig.e-invoices@inter.ikea.com"
inbox_folder_name = "Inbox"  # The main Inbox folder
ap_mailbox_subfolder_name = "APMailbox"  # Subfolder within Inbox
tool_pickup_subfolder_name = "ToolPickup"  # Subfolder within Inbox


# Define a function to fetch attachments for a message
def fetch_attachments(message_id, save_folder, headers):
    try:
        attachments_url = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/attachments"
        response = requests.get(attachments_url, headers=headers)
        response.raise_for_status()
        attachment_items = response.json().get('value', [])

        for attachment in attachment_items:
            file_name = attachment['name']
            attachment_id = attachment['id']
            attachment_content_url = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/attachments/{attachment_id}/$value"
            attachment_response = requests.get(attachment_content_url, headers=headers)
            attachment_response.raise_for_status()

            log(f"Saving file {file_name}...")
            with open(os.path.join(save_folder, file_name), 'wb') as file:
                file.write(attachment_response.content)

        return True
    except requests.exceptions.RequestException as e:
        log(f"Error fetching or saving attachments: {e}")
        return False


# Define a function to check if an attachment has an XML content type
def is_xml_attachment(attachment):
    content_type = attachment.get("contentType", "")
    return content_type.startswith("application/xml")


# Authenticate and obtain an access token
# (Assuming you have the generate_access_token function)
access_token = generate_access_token(app_id=client_id, scopes=["Mail.ReadWrite"])
headers = {
    'Authorization': 'Bearer ' + access_token['access_token']
}

# Retrieve folders inside the Inbox
folders_url = f"https://graph.microsoft.com/v1.0/users/{email_address}/mailFolders/{inbox_folder_name}/childFolders"
try:
    response = requests.get(folders_url, headers=headers)
    response.raise_for_status()
    folders = response.json().get("value", [])
except requests.exceptions.RequestException as e:
    log(f"Error fetching folders: {e}")
    exit(1)

# Process email messages in the "Inbox/APMailbox" subfolder
for folder in folders:
    if folder["displayName"] == ap_mailbox_subfolder_name:
        ap_mailbox_subfolder_id = folder["id"]
        inbox_messages_url = f"https://graph.microsoft.com/v1.0/users/{email_address}/mailFolders/{ap_mailbox_subfolder_id}/messages"

        log(f"Fetching email messages from the '{ap_mailbox_subfolder_name}' subfolder...")
        try:
            response = requests.get(inbox_messages_url, headers=headers)
            response.raise_for_status()
            inbox_messages = response.json().get("value", [])
            log(f"Received {len(inbox_messages)} email messages from the '{ap_mailbox_subfolder_name}' subfolder.")
        except requests.exceptions.RequestException as e:
            log(f"Error fetching email messages: {e}")
            continue

        for message in inbox_messages:
            # Check if the message meets your condition
            attachments = message.get("attachments", [])
            xml_attachments = [attachment for attachment in attachments if is_xml_attachment(attachment)]

            if len(xml_attachments) == 1:
                log(f"Analyzing message with subject: {message['subject']} from the '{ap_mailbox_subfolder_name}' subfolder...")

                # Additional logging for analyzing attachments
                for idx, attachment in enumerate(xml_attachments, start=1):
                    content_type = attachment.get("contentType", "")
                    log(f"Attachment {idx} - Content Type: {content_type}")
                    log(f"Attachment {idx} - Name: {attachment.get('name', '')}")
                    log(f"Attachment {idx} - Size: {attachment.get('size', 0)} bytes")

                # Fetch and save attachments for this message
                fetch_attachments(message['id'], '/path/to/save/folder', headers)

                # Move the message to the "Inbox/ToolPickup" subfolder as needed
                # You can implement the move operation here
                # Example:
                # move_message(message['id'], tool_pickup_subfolder_id, headers)

log("Script execution completed.")
