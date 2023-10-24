import os
import xml.etree.ElementTree as ET

def remove_extra_extension(filename):
    return filename.split(".xml")[0] + ".xml"

def truncate_comment_value(comment, max_chars=49):
    return comment[:max_chars]

def process_xml_files(folder_path):
    for filename in os.listdir(folder_path):
        if filename.endswith(".bak"):
            full_path = os.path.join(folder_path, filename)

            # Action 1: Remove everything after .xml from the filename
            new_filename = remove_extra_extension(filename)
            new_full_path = os.path.join(folder_path, new_filename)

            os.rename(full_path, new_full_path)

            # Action 2: Process the XML file
            tree = ET.parse(new_full_path)
            root = tree.getroot()

            for comment_element in root.findall(".//Comment"):
                comment = comment_element.text
                if comment and len(comment) > 49:
                    comment_element.text = truncate_comment_value(comment)

            tree.write(new_full_path)

if __name__ == "__main__":
    folder_path = "XMLFiles"
    process_xml_files(folder_path)
