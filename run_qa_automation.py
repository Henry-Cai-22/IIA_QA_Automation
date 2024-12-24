from flask import Flask, redirect, url_for, session, request, jsonify
import extract_msg
import io
from msal import ConfidentialClientApplication
import requests
import time
from cred import *
from utils import *
import pandas as pd

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

@app.route('/graph_call')
def graph_call():

    dv_filled_in_working_and_outgoing_no_dv_sheet_df = pd.DataFrame()
    dv_filled_in_working_and_outgoing_has_filled_dv_sheet_df = pd.DataFrame()
    dv_not_filled_in_working_outgoing_no_filled_dv_sheet_df = pd.DataFrame()
    dv_not_filled_in_working_outgoing_no_dv_sheet_df = pd.DataFrame()
    no_dv_in_working_or_outgoing_df = pd.DataFrame()
    attachment_not_found_in_outgoing_folder_df = pd.DataFrame()

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
        valid_extensions = ('.docx', '.pptx', '.pdf')


        if i != 11 and i != 16:
            continue

        try:
            web_url = item['webUrl']
            item_id = get_weburl_item_id(web_url=web_url, drives_id=EMAIL_LIBRARY_DRIVE_ID,headers=headers)
            item_content_url  = f'{GRAPH_API_URL}/drives/{EMAIL_LIBRARY_DRIVE_ID}/items/{item_id}/content'
            msg_response = requests.get(item_content_url, headers=headers)
            msg = extract_msg.Message(io.BytesIO(msg_response.content))
            attachments = msg.attachments
            print("Attachments")

            for attachment in attachments:

                if not attachment.longFilename.lower().endswith(valid_extensions):
                    print("not valid ext skipping: " + attachment.longFilename)
                    continue

                print(f"Attachment: {attachment.longFilename}")
                attachemnt_name = attachment.longFilename
                attachment_name_without_extension = os.path.splitext(attachemnt_name)[0]
                attachment_name_without_extension = attachment_name_without_extension.strip()

                print("FINDING:", attachment_name_without_extension)

                for i, task_folder in enumerate(TASK_FOLDERS):
                    task_drive_id = task_folder['task_drive_id']
                    task_working_folder_id = task_folder['working_folder_id']
                    task_outgoing_folder_id = task_folder['outgoing_folder_id']
                    
                    print("Scanning working folder")
                    working_folder_statuses = DV_sheet_exists_status(headers=headers,
                                           task_drive_id=task_drive_id,
                                           task_folder_id=task_working_folder_id,
                                           attachment_name_without_extension=attachment_name_without_extension)    

                    if working_folder_statuses['is_file_exists']:
                        print("Working folder file exists")
                        break

                for i, task_folder in enumerate(TASK_FOLDERS):
                    task_drive_id = task_folder['task_drive_id']
                    task_working_folder_id = task_folder['working_folder_id']
                    task_outgoing_folder_id = task_folder['outgoing_folder_id']

                    print("Scanning outgoing folder")
                    outgoing_folder_statuses = DV_sheet_exists_status(
                        headers=headers,
                        task_drive_id=task_drive_id,
                        task_folder_id=task_outgoing_folder_id,
                        attachment_name_without_extension=attachment_name_without_extension) 

                    if outgoing_folder_statuses['is_file_exists']:
                        print("Outgoing folder file exists")
                        break
                
                print("==========RESULTS===========")
                print("Working folder statuses")
                print(working_folder_statuses)
                print("Outgoing folder statuses")
                print(outgoing_folder_statuses)

                if not working_folder_statuses['is_file_exists'] or not outgoing_folder_statuses['is_file_exists']:
                    no_dv_in_working_or_outgoing_df = attachment_not_found_in_outgoing_folder_df.append({
                        "Attachment Name": attachemnt_name,
                        "Working Folder Exists": working_folder_statuses['is_file_exists'],
                        "Outgoing Folder Exists": outgoing_folder_statuses['is_file_exists']
                    }, ignore_index=True)

                if working_folder_statuses['is_file_exists'] or not outgoing_folder_statuses['is_file_exists']:
                    attachment_not_found_in_outgoing_folder_df = attachment_not_found_in_outgoing_folder_df.append({
                        "Attachment Name": attachemnt_name,
                        "Working Folder Exists": working_folder_statuses['is_file_exists'],
                        "Outgoing Folder Exists": outgoing_folder_statuses['is_file_exists']
                    }, ignore_index=True)

        except Exception as e:
            print("An error occured processing: ", e)
            

    with pd.ExcelWriter("QA_Automation_Output.xlsx", engine="xlsxwriter") as writer:
        # Separate data based on condition and write to different sheetz

        # Write each DataFrame to a different worksheet
        no_dv_in_working_or_outgoing_df.to_excel(writer, sheet_name="5_no_dv_in_working_or_outgoing", index=False)
        attachment_not_found_in_outgoing_folder_df.to_excel(writer, sheet_name="6_no_outgoing_copy", index=False)

    return filtered_data
    
if __name__ == '__main__':
    app.run(host='localhost', port=5000, debug=True)