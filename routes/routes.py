from flask import Blueprint, redirect, url_for, redirect, url_for, session, request
from msal import ConfidentialClientApplication
from cred import *
from utils import *
import threading

routes = Blueprint('routes', __name__)

@routes.route('/')
def index():
    return redirect(url_for('routes.login'))

@routes.route('/login')
def login():
    client = ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    auth_url = client.get_authorization_request_url(SCOPE, redirect_uri=url_for('routes.authorized', _external=True))

    print("AUTH URL")
    print(auth_url)
    return redirect(auth_url)

@routes.route(REDIRECT_PATH)
def authorized():
    client = ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    code = request.args.get('code')
    result = client.acquire_token_by_authorization_code(code, scopes=SCOPE, redirect_uri=url_for('routes.authorized', _external=True))
    if 'access_token' in result:
        session['access_token'] = result['access_token']
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
    
    headers = {'Authorization': 'Bearer ' + token}
    drives_url  = f'{GRAPH_API_URL}/sites/{SITE_ID}/lists/{EMAIL_LIBRARY_LIST_ID}/items?$expand=fields&$filter=fields/Arup_AttachmentsCount gt 0&$top={TOP_EMAIL_LIBRARY_COUNT}'
    
    # Run QA Automation in background so that the frontend can return a response immediately
    task_thread = threading.Thread(target=run_qa_automation_in_background, args=(drives_url, headers))
    task_thread.start()

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
    return "TODO: Implement this route"