from flask import Flask, redirect, url_for, session, request, jsonify
import extract_msg
import io
from msal import ConfidentialClientApplication
import requests
import time
import os
import time
from cred import *

app = Flask(__name__)
app.secret_key = FLASK_SECRET_KEY

@app.route('/')
def index():
    return 'QA Automation Script'

@app.route('/login')
def login():
    client = ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    auth_url = client.get_authorization_request_url(SCOPE, redirect_uri=url_for('authorized', _external=True))

    print("AUTH URL")
    print(auth_url)
    return redirect(auth_url)

@app.route(REDIRECT_PATH)
def authorized():
    client = ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    code = request.args.get('code')
    result = client.acquire_token_by_authorization_code(code, scopes=SCOPE, redirect_uri=url_for('authorized', _external=True))
    if 'access_token' in result:
        session['access_token'] = result['access_token']
        return redirect(url_for('graph_call'))
    return 'Login failed'


"""
Debug getting msg content
"""
@app.route('/get_msg_content')
def test_get_msg_content():
    token = session.get('access_token')
    if not token:
        return redirect(url_for('login'))
    headers = {'Authorization': 'Bearer ' + token}
    drives_url  = f'{GRAPH_API_URL}/drives/{EMAIL_LIBRARY_DRIVE_ID}/items/item-id/content'
    response = requests.get(drives_url, headers=headers)
    msg = extract_msg.Message(io.BytesIO(response.content))
    # Extract attachments
    attachments = msg.attachments

    for attachment in attachments:
        print(f"Attachment: {attachment.longFilename}")
        attachment_data = attachment.data


    print(attachments)
    return {"message": 'ok'}

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

@app.route('/graph_call')
def graph_call():
    token = session.get('access_token')
    if not token:
        return redirect(url_for('login'))
    headers = {'Authorization': 'Bearer ' + token}

    drives_url  = f'{GRAPH_API_URL}/sites/{SITE_ID}/lists/{EMAIL_LIBRARY_LIST_ID}/items?$expand=fields&$filter=fields/Arup_AttachmentsCount gt 0&$top={TOP_EMAIL_LIBRARY_COUNT}'

    response = requests.get(drives_url, headers=headers)

    data = response.json().get('value', [])

    print("data length", len(data))
    filtered_data = [
        item for item in data
        if '@HSR' in item.get('fields', {}).get('Arup_To', '')
    ]

    print(len(filtered_data))
    # Output the filtered data
    print(filtered_data[0])

    for i, item in enumerate(filtered_data):
        print("====Item processing===", i)
        try:
            web_url = item['webUrl']
            item_id = get_weburl_item_id(web_url=web_url, drives_id=EMAIL_LIBRARY_DRIVE_ID,headers=headers)
            item_content_url  = f'{GRAPH_API_URL}/drives/{EMAIL_LIBRARY_DRIVE_ID}/items/{item_id}/content'
            msg_response = requests.get(item_content_url, headers=headers)
            msg = extract_msg.Message(io.BytesIO(msg_response.content))
            attachments = msg.attachments
            print("Attachments")
            for attachment in attachments:

                print(f"Attachment: {attachment.longFilename}")
                attachemnt_name = attachment.longFilename
                attachment_name_without_extension = os.path.splitext(attachemnt_name)[0]
                attachment_name_without_extension = attachment_name_without_extension.strip()

                print("FINDING:", attachment_name_without_extension)
                file = find_file_in_subfolders(site_id=SITE_ID, drive_id=TASK_01_DRIVE_ID, parent_item_id=TASK_01_WORKING_FILES_ID, filename=attachment_name_without_extension, headers=headers)

                if file:
                    print(f"File found: {file['name']}")
                else:
                    print("File not found")

        except Exception as e:
            print("An error occured processing: ", e)

    return filtered_data
    
if __name__ == '__main__':
    app.run(host='localhost', port=5000, debug=True)