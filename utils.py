from cred import *
import requests
from lxml import etree
import zipfile
import time
import os
import io

TASK_FOLDERS = [
    { "folder_name": "Task 01 - Renewable Energy", 
    "task_drive_id": TASK_01_DRIVE_ID, 
    "working_folder_id" : TASK_01_WORKING_FILES_ID,
    "outgoing_folder_id": TASK_01_OUTGOING_FILES_ID
    },
    { "folder_name": "Task 02 - Climate Adaptation", 
    "task_drive_id": TASK_02_DRIVE_ID, 
    "working_folder_id" : TASK_02_WORKING_FILES_ID,
    "outgoing_folder_id": TASK_02_OUTGOING_FILES_ID
    },
    { "folder_name": "Task 03 - Strategy", 
    "task_drive_id": TASK_03_DRIVE_ID, 
    "working_folder_id" : TASK_03_WORKING_FILES_ID,
    "outgoing_folder_id": TASK_03_OUTGOING_FILES_ID
    },
    { "folder_name": "Task 04 - Reporting", 
    "task_drive_id": TASK_04_DRIVE_ID, 
    "working_folder_id" : TASK_04_WORKING_FILES_ID,
    "outgoing_folder_id": TASK_04_OUTGOING_FILES_ID
    },
    { "folder_name": "Task 05 - GHG & Air Quality", 
    "task_drive_id": TASK_05_DRIVE_ID, 
    "working_folder_id" : TASK_05_WORKING_FILES_ID,
    "outgoing_folder_id": TASK_05_OUTGOING_FILES_ID
    },
]
"""
Get the id of the .msg file in Email Library
"""
def get_weburl_item_id(web_url, drives_id,  headers):
    file_name = web_url.split("/")[-1]
    msg_item_url = f'{GRAPH_API_URL}/drives/{drives_id}/root:/{file_name}'
    response = requests.get(msg_item_url, headers=headers)
    item_id = response.json().get('id')
    print("THE ITEM ID")
    print(item_id)
    return item_id

def list_children(site_id, drive_id, item_id, headers):
    url = f"{GRAPH_API_URL}/sites/{site_id}/drives/{drive_id}/items/{item_id}/children"
    response = requests.get(url, headers=headers)
    return response.json()

def fetch_file_content(site_id, drive_id, file_id, headers):
    url = f"{GRAPH_API_URL}/sites/{site_id}/drives/{drive_id}/items/{file_id}/content"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.content  # Return raw file content
    else:
        return None

"""
Using search in graphapi to recursively find the partial match of the file name in SharePoint folder
"""
def search_files_in_folder(site_id, drive_id, folder_id, query, headers):
    url = f"{GRAPH_API_URL}/sites/{site_id}/drives/{drive_id}/items/{folder_id}/search(q='{query}')"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()  # Return the search results
    else:
        return None

"""
Check for matching extension for .docx and .pptx from the returned search results
"""
def find_file_in_subfolders(site_id, drive_id, parent_item_id, filename, headers):
    search_results = search_files_in_folder(site_id, drive_id, parent_item_id, filename, headers)
    if search_results:
        print("THE SEARCH RESULTS")
        print(search_results)
        valid_extensions = ('.docx', '.pptx')
        for item in search_results['value']:
            # Check if the file name matches and has a valid extension
            if item['name'].lower().startswith(filename.lower()) and item['name'].lower().endswith(valid_extensions):
                return item  # Return the file item if found
    return None  # Return None if the file is not found


"""
Returns the following statuses in an dictionary

1. whether file exists
2. whether the DV sheet exists in specified folder
3. whether the DV sheet is filled out or is blank

example result
{
    "is_file_exists": True,
    "is_dv_sheet_exists": True,
    "is_dv_sheet_filled": True
}

"""
def DV_sheet_exists_status(headers, 
                           task_drive_id,
                           task_folder_id,
                           attachment_name_without_extension):
    
    response_dict = {
        "is_file_exists": False,
        "is_dv_sheet_exists": False,
        "is_dv_sheet_filled": False
    }
    file = find_file_in_subfolders(site_id=SITE_ID, drive_id=task_drive_id, parent_item_id=task_folder_id, filename=attachment_name_without_extension, headers=headers)

    if file:
        print(f"File found: {file['name']}")
        response_dict["is_file_exists"] = True
    else:
        print("File not found")
        return response_dict
    
    working_item_content_url  = f'{GRAPH_API_URL}/drives/{file["parentReference"]["driveId"]}/items/{file["id"]}/content'
    working_msg_response = requests.get(working_item_content_url, headers=headers)

    file_extension = os.path.splitext(file['name'])[1].lower()

    if file_extension == '.pptx':
        print("The file is a PPTX.")
    elif file_extension == '.docx':
        print("The file is a DOCX.")

        doc = io.BytesIO(working_msg_response.content)
        with zipfile.ZipFile(doc) as docx_zip:
            xml_content = docx_zip.read('word/document.xml')
        
        target_fields = [
            "Signature", "Filename", "Description", 
            "Prepared by", "Checked by", "Approved by", "Name"
        ]

        tree = etree.XML(xml_content)
        namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        text_elements = tree.xpath("//w:t", namespaces=namespaces)
        content = "\n".join([element.text for element in text_elements if element.text])
        print(content) 

        if "Document Verification" in content:
            print("DV Sheet found in Working Folder")
            response_dict["is_dv_sheet_exists"] = True

            for field in target_fields:
                for idx, text in enumerate(content):
            
                    # Check if the current text is one of the target fields
                    if field in text:
                        # Check the next element or the surrounding text (if any) to see if it's the person's name
                        # For simplicity, check the next text in the list after the field text
                        next_text = content[idx + 1] if idx + 1 < len(content) else None
                        
                        # Print field and next text (for inspection)
                        #print(f"Field: {field}, Next Text: {next_text}")

                        # You can add custom logic here to check if the next text is a name (e.g., by matching patterns)
                        if next_text and "Name" in next_text:
                            print(f"Detected name after '{field}': {next_text}")
                            response_dict["is_dv_sheet_filled"] = True
        
        else:
            print("No DV Sheet in Working Folder")
    else:
        print("The file is neither a pptx nor a DOCX.")
    
    return response_dict