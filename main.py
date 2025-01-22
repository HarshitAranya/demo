
import os
import sys
import datetime
import time
import win32com.client

# Get the directory where the .exe file is located
current_directory = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.realpath(__file__))

# List all .docx files in the directory
docx_files = [f for f in os.listdir(current_directory) if f.endswith('.docx')]

# If there are .docx files in the directory
if docx_files:
    # Get the latest .docx file based on the modification time
    latest_file = max(docx_files, key=lambda f: os.path.getmtime(os.path.join(current_directory, f)))
    latest_file_path = os.path.join(current_directory, latest_file)
    
    # Print or use the path of the latest file
    print(f"The latest .docx file is: {latest_file}")
    print(f"File path: {latest_file_path}")
else:
    print("No .docx files found in the directory.")
    time.sleep(10)
    exit()

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

if "~$" in latest_file_path:
    print(f"The file {latest_file} is currently open.")
    document_count = word.Documents.Count
    print(f"Number of open documents: {document_count}")
    doc.Close(False)
    word.Quit()
    time.sleep(5)
    exit()
else:
    try:
        doc = word.Documents.Open(latest_file_path)
    except Exception as e:
        print(f"Error opening document: {e}")
        word.Quit()

paragraph_dict = {}
# Loop through all paragraphs and print index and text
x=1
for index, paragraph in enumerate(doc.Paragraphs, start=1):
    # if index in prange:
    ptext = paragraph.Range.Text.strip()
    ptext = ptext.replace('\r', '').replace('\x07', '')
    paragraph_dict[x] = ptext
    x += 1

 for key, value in paragraph_dict.items():
     # print(f"{key}: {value}")
     if "Data/Code" in value:
         # print(f"{key}: {value}")
         ocrType = paragraph_dict[key+1]
         print(ocrType)
         print(paragraph_dict[key+1])
         # print(paragraph_dict[key+0])

 if (doc.FullName == latest_file_path):
     doc.Close(False)  # False means don't save changes

 word.Quit()
 exit()
