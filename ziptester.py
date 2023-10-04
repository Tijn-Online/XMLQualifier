import zipfile

# Replace 'your_zip_file.zip' with the path to your ZIP file
zip_file_path = r'C:\Users\mahoo8\OneDrive - Inter IKEA Group\Desktop\dzfp_22442000000003600274_20221129114046.zip'

try:
    with zipfile.ZipFile(zip_file_path, 'r') as zip_file:
        # List the names of files in the ZIP archive
        file_names = zip_file.namelist()
        print("Files in the ZIP archive:")
        for file_name in file_names:
            print(file_name)
except zipfile.BadZipFile as e:
    print(f"Error: Not a valid ZIP file. {e}")
except Exception as e:
    print(f"Error extracting ZIP file: {e}")