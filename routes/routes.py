from flask import Blueprint, redirect, url_for, redirect, url_for, session, request
from msal import ConfidentialClientApplication
from cred import *
from utils.utils import *
import threading

routes = Blueprint('routes', __name__)

@routes.route('/')
def index():
    print("Redirecting to login")
    return redirect(url_for('routes.login'))

@routes.route('/login')
def login():
    client = ConfidentialClientApplication(
        client_id=CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    auth_url = client.get_authorization_request_url(SCOPE, redirect_uri=url_for('routes.authorized', _external=True))
    return redirect(auth_url)

@routes.route(REDIRECT_PATH)
def authorized():
    client = ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    code = request.args.get('code')
    result = client.acquire_token_by_authorization_code(code, scopes=SCOPE, redirect_uri=url_for('routes.authorized', _external=True))

    if 'access_token' in result and 'refresh_token' in result:
        session['access_token'] = result['access_token']
        session['refresh_token'] = result['refresh_token']
        return redirect(url_for('routes.graph_call'))
    return 'Login failed'

"""
This route is used to call the QA Automation process in the background.
"""
@routes.route('/call_qa_automation')
def graph_call():
    token = session.get('access_token')
    if not token:
        return redirect(url_for('routes.login'))
    
    refresh_token = session.get('refresh_token')
    headers = {'Authorization': 'Bearer ' + token}
    drives_url  = f'{GRAPH_API_URL}/sites/{SITE_ID}/lists/{EMAIL_LIBRARY_LIST_ID}/items?$expand=fields&$filter=fields/Arup_AttachmentsCount gt 0&$top={TOP_EMAIL_LIBRARY_COUNT}'
    
    # Run QA Automation in background so that the frontend can return a response immediately
    task_thread_a = threading.Thread(target=run_qa_automation_A_in_background, args=(drives_url, headers, refresh_token))
    task_thread_a.start()

    task_thread_b = threading.Thread(target=run_qa_automation_B_in_background, args=(drives_url, headers, refresh_token))
    task_thread_b.start()

    return "Running QA Automation in background. Once the process is completed, \
the results will be saved to an .xlsx file within this directory."




"""
==== THE ROOTS BELOW ARE FOR DEBUGGING PURPOSES ====
"""

"""
Debug getting msg content
"""
@routes.route('/get_msg_content')
def test_get_msg_content():
    token = session.get('access_token')
    if not token:
        return redirect(url_for('debug_routes.login'))
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
Get list of internal members from SharePoint list to use find filter out internal members with @HSR account
"""
@routes.route('/list_internal_members')
def list_internal_members():
    token = session.get('access_token')
    if not token:
        return redirect(url_for('debug_routes.login'))
    headers = {'Authorization': 'Bearer ' + token}

    drives_url = f'{GRAPH_API_URL}/sites/{SITE_ID}/lists/{INTERNAL_MEMBERS_LIST_ID}/items?$expand=fields'
    response = requests.get(drives_url, headers=headers)
    data = response.json()

    internal_members = []

    # Extract internal members names
    for item in data['value']:
        internal_members.append(item['fields']['Title'])
    
    drives_url_email_library  = f'{GRAPH_API_URL}/sites/{SITE_ID}/lists/{EMAIL_LIBRARY_LIST_ID}/items?$expand=fields&$filter=fields/Arup_AttachmentsCount gt 0&$top={5000}'
    response = requests.get(drives_url_email_library, headers=headers)

    data = response.json().get('value', [])

    print("LENGTH OF DATA FROM EmAIL LIBRARY", len(data))

    # filtered_data = [
    #     item for item in data
    #     if '@HSR' in item.get('fields', {}).get('Arup_To', '')
    # ]

    # print("LENGTH OF FILTERED DATA", len(filtered_data))

    result_filted_data = []

    for item in data:
        email = item['fields'].get('Arup_To')
        if not email:
            continue

        name = switch_name_format(email)
        if name:
            if name not in internal_members:
                print("NAME NOT IN:", name)
                result_filted_data.append(item)
        # else:
            # print(f"{email.strip()} is not a valid HSR email.")

    print("LENGTH OF RESULT FILTERED DATA for non internal members", len(result_filted_data))
    return {"result_filterd_data": result_filted_data}