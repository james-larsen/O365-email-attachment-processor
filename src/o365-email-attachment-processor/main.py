"""Check multiple O365 service accounts for new emails, evaluates them against pre-defined patterns, and delivers any attachments accordingly or forwards the email"""
#%%
import os
import datetime
import io
import json
import base64
from pathlib import Path
from azure.identity import ClientSecretCredential
from msgraph.core import GraphClient
import openpyxl
import boto3
# import pysftp
# import configparser
# pylint: disable=import-error
# from utils.password import get_password as pw
# pylint: enable=import-error

#%%

def authenticate(tenant_id, client_id, client_secret):
    """Authenticate with the O365 server to return a client object"""

    # Create a ClientSecretCredential object
    credential = ClientSecretCredential(tenant_id=tenant_id, client_id=client_id, client_secret=client_secret)

    # Create a GraphClient object
    client = GraphClient(credential=credential)

    return client

def get_sharepoint_folder(sharepoint_client, o365_site_address, o365_site_name, o365_site_folderpath):
    """Retrieve the address of the Sharepoint folder holding rules definitions"""
    
    try:
        site_id = sharepoint_client.get(f'https://graph.microsoft.com/v1.0/sites/{o365_site_address}:/sites/{o365_site_name}').json()['id']

        folder_list = sharepoint_client.get(f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives?$select=id,name').json()['value']
        folder_id = ''
        path_folders = o365_site_folderpath.split("/")

        # retrieve ID for the "Documents" base folder
        for folder in folder_list:
            if folder['name'] == 'Documents':
                documents_folder_id = folder['id']

        folder_list = sharepoint_client.get(f'https://graph.microsoft.com/v1.0/drives/{documents_folder_id}/root/children?$select=folder,name,id').json()['value']
        
        # traverse through underlying subfolders
        for path_folder in path_folders[1:]:
            for folder in folder_list:
                if folder['name'] == path_folder:
                    folder_id = folder['id']
                    # folder_name = folder['name']
                    folder_list = sharepoint_client.get(f'https://graph.microsoft.com/v1.0/drives/{documents_folder_id}/items/{folder_id}/children').json()['value']

        return f'https://graph.microsoft.com/v1.0/drives/{documents_folder_id}/items/{folder_id}/children'
    except:
        return None

def retrieve_rules(email_account_name):
    """Retrieve the pattern and delivery rules from local JSON and optional Sharepoint .xlsx files"""
    global account
    json_filename = f'{email_account_name}_email_rules.json'

    if os.path.exists(json_filename):
        with open(json_filename, mode='rb') as f:
            content = f.read().decode('utf-8').replace('\\', '\\\\')
            email_rules = json.loads(content)
    elif os.path.exists('default_email_rules.json'):
        with open('default_email_rules.json', mode='rb') as f:
            content = f.read().decode('utf-8').replace('\\', '\\\\')
            email_rules = json.loads(content)
    else:
        pass

    if 'sharepoint_account' in account:
        sharepoint_account = account['sharepoint_account']
        sharepoint_account_name = sharepoint_account['account_name']
        # o365_sharepoint_username = sharepoint_account['o365_username']
        o365_site_address = sharepoint_account['o365_site_address']
        o365_site_name = sharepoint_account['o365_site_name']
        o365_site_folderpath = sharepoint_account['o365_site_folderpath']
        # o365_sharepoint__user_id = sharepoint_account['o365_user_id']
        o365_sharepoint_tenant_id = sharepoint_account['o365_tenant_id']
        o365_sharepoint_client_id = sharepoint_account['o365_client_id']
        sharepoint_password_key = sharepoint_account['o365_password_key']
        o365_sharepointpassword = pw(sharepoint_account_name, sharepoint_password_key)
        sharepoint_client = authenticate(o365_sharepoint_tenant_id, o365_sharepoint_client_id, o365_sharepointpassword)

        sharepoint_folder = get_sharepoint_folder(sharepoint_client, o365_site_address, o365_site_name, o365_site_folderpath)
        
        if sharepoint_folder is not None:
            try:
                files_list = sharepoint_client.get(sharepoint_folder).json()['value']

                for item in files_list:
                    # check if the item is an Excel file
                    if item['name'].endswith('.xlsx'):
                        # get the download URL
                        download_url = item['@microsoft.graph.downloadUrl']
                        # access the file content
                        file_content = sharepoint_client.get(download_url).content
                        # print(item['@microsoft.graph.downloadUrl'])
                        # print(item['name'])
                        file_obj = io.BytesIO(file_content)
                        # load the Excel workbook from file_obj
                        workbook = openpyxl.load_workbook(file_obj)
                        # get the active sheet (i.e., the first sheet)
                        # worksheet = workbook.active
                        worksheet = workbook['Email Rules']
                        # print the value in cell A1
                        # print(sheet['A1'].value)

                        for row in worksheet.iter_rows(min_row=2, values_only=True):
                            # create a new condition dictionary
                            condition = {}
                            
                            # add the name
                            condition["name"] = row[0]
                            
                            # add the pattern dictionary
                            condition["pattern"] = {}
                            
                            # add the sender
                            condition["pattern"]["sender"] = row[1]
                            
                            # add the subject as a list
                            subject_list = row[2]
                            subject_list = subject_list.replace("\n", "|").replace("||", "|")
                            condition["pattern"]["subject"] = [x.strip() for x in subject_list.split("|")]
                            
                            # add the body as a list
                            body_list = row[3]
                            body_list = body_list.replace("\n", "|").replace("||", "|")
                            condition["pattern"]["body"] = [x.strip() for x in body_list.split("|")]
                            
                            # add the attachments filename as a list
                            attachments_list = row[4]
                            attachments_list = attachments_list.replace("\n", "|").replace("||", "|")
                            condition["pattern"]["attachments"] = [{"filename": [x.strip()]} for x in attachments_list.split("|")]
                            
                            # add the recipients as a list
                            recipients_list = row[5]
                            recipients_list = recipients_list.replace("\n", "|").replace("||", "|")
                            condition["delivery"] = {"target": "email_forward", "recipients": [x.strip() for x in recipients_list.split("|")], "body": row[6]}
                        
                            # add the condition to the output dictionary
                            email_rules["conditions"].append(condition)

                # print(json.dumps(email_rules, indent=4))
            except:
                return email_rules
    
    return email_rules

def transmit_files(condition_name, target, delivery_details, email_date, attachment_name, attachment_content):
    """Transmit files to an target location"""
    # attachment_content = part.get_payload(decode=True)
    if 'append_datetime' in delivery_details and str(delivery_details['append_datetime'].lower()) == 'true':
        # datetime_string = datetime.datetime.fromtimestamp(datetime.datetime.now().timestamp()).strftime("%Y-%m-%d_%H%M%S")
        datetime_obj = datetime.datetime.strptime(email_date, "%Y-%m-%dT%H:%M:%SZ")
        datetime_string = datetime_obj.strftime("%Y-%m-%d_%H%M%S")
        parts = attachment_name.split('.')
        base_attachment_name = '.'.join(parts[:-1])
        extension = parts[-1]
        attachment_name = f"{base_attachment_name}_{datetime_string}.{extension}"
    
    if target == 'local':
        delivery_path = delivery_details['path']
        if not os.path.exists(delivery_path):
            os.makedirs(delivery_path)
        filepath = Path(delivery_path) / attachment_name
        with open(filepath, 'wb') as f:
            f.write(attachment_content)

    if target == 's3':
        bucket_region = delivery_details['region']
        bucket_name = delivery_details['bucket']
        if 'subfolder' in delivery_details:
            subfolder_name = delivery_details['subfolder']
            attachment_name = subfolder_name + attachment_name
        s3_access_key = pw(condition_name, "S3AccessKey")
        s3_secret_key = pw(condition_name, "S3SecretKey")

        # Create an S3 client using the access key and secret key
        s3 = boto3.client('s3', aws_access_key_id=s3_access_key, aws_secret_access_key=s3_secret_key, region_name=bucket_region)

        s3.put_object(Body=attachment_content, Bucket=bucket_name, Key=attachment_name)

    # if target == 'sftp':
    #     sftp_hostname = delivery_details['hostname']
    #     sftp_port = int(delivery_details['port'])
    #     sftp_username = delivery_details['username']
    #     sftp_password_key = delivery_details['password_key']
    #     if 'subfolder' in delivery_details:
    #         sftp_subfolder = delivery_details['subfolder']
    #     sftp_password = pw(condition_name, sftp_password_key)
        
    #     cnopts = pysftp.CnOpts()
    #     cnopts.hostkeys = None

    #     with pysftp.Connection(host=sftp_hostname, username=sftp_username, password=sftp_password, cnopts=cnopts) as sftp:
    #         # Change to the remote directory
    #         sftp.chdir(sftp_subfolder)
            
    #         # Get a list of files in the directory
    #         file_list = sftp.listdir()

    #         # Print the file names
    #         for file_name in file_list:
    #             print(file_name)

def forward_email(message, delivery_details):
    """Forward the email to any number of recipients"""
    global email_client, o365_email_user_id
    message_id = message['id']
    recipients = delivery_details['recipients']
    # custom_subject = delivery_details.get('subject', '') # have not been able to get overwriting the subject line working
    custom_body = delivery_details.get('body', '')
    forward_endpoint = f"/users/{o365_email_user_id}/messages/{message_id}/forward"

    # Define the recipient list
    to_recipients = [{'emailAddress': {'address': recipient}} for recipient in recipients]

    # Define the payload
    payload = {
        'toRecipients': to_recipients,
        'comment': custom_body,
        'send': True
    }

    # if custom_subject != '':
    #     payload['subject'] = custom_subject

    # Forward the email
    email_client.post(forward_endpoint, json=payload)


current_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(current_dir)

if os.path.exists('o365_accounts_local.json'):
    with open('o365_accounts_local.json', encoding='utf-8') as f:
        o365_accounts = json.load(f)
else:
    with open('o365_accounts.json', encoding='utf-8') as f:
        o365_accounts = json.load(f)

#%%
for account in o365_accounts['o365_accounts']:
    # choose appropriate password method
    password_method = account['password_method'].lower()
    # pylint: disable=import-error
    if password_method == 'keyring':
        from utils.password_keyring import get_password as pw
    elif password_method == 'secretsmanager':
        from utils.password_aws import get_password as pw
    elif password_method == 'ssm':
        from utils.password_ssm import get_password as pw
    if password_method == 'custom':
        from utils.password_custom import get_password as pw
    # pylint: enable=import-error

    email_account = account['email_account']
    email_account_name = email_account['account_name']
    o365_email_username = email_account['o365_username']
    o365_email_user_id = email_account['o365_user_id']
    o365_email_tenant_id = email_account['o365_tenant_id']
    o365_email_client_id = email_account['o365_client_id']
    email_password_key = email_account['o365_password_key']
    o365_emailpassword = pw(email_account_name, email_password_key)
    # build initial email client to check for any unread messages
    email_client = authenticate(o365_email_tenant_id, o365_email_client_id, o365_emailpassword)

    # search_parameters = {
    #     "$search": "isRead eq false",
    #     "$orderby": "receivedDateTime desc"
    # }

    # define the API endpoint for retrieving emails
    api_endpoint = f"/users/{o365_email_user_id}/mailfolders/inbox/messages?$filter=isRead eq false&$orderby=receivedDateTime desc"
    result = email_client.get(api_endpoint).json()

    # for message in result['value']:
    #     print(message['subject'].lower() + ' | ' + message['receivedDateTime'])

    if len(result) == 0:
        print('No Unread messages to process')
        continue

    # read rules
    email_rules = retrieve_rules(email_account_name)

    # rebuild email client after potentially checking Sharepoint - building the Sharepoint client seems to affect the email client
    email_client = authenticate(o365_email_tenant_id, o365_email_client_id, o365_emailpassword)

    # loop through all unread emails
    for message in result['value']:
        # Fetch the email content
        email_id = message['id']

        # extract relevant fields from the email message
        email_subject = message['subject'].lower()
        email_from = message['from']['emailAddress']['address'].lower()
        email_to = [recipient['emailAddress']['address'].lower() for recipient in message['toRecipients']]
        email_date = message['receivedDateTime']

        if 'undeliverable' in email_subject:
            # mark the email as read
            email_client.patch(f"/users/{o365_email_user_id}/messages/{email_id}", json={'isRead': True})
            continue

        response = email_client.get(f"{api_endpoint}", params={'$select': 'body'})
        email_body = ''

        # check if the email is multipart
        if 'multipart' in response.json()['value'][0]['body']['contentType']:
            for part in response.json()['body']['content']:
                # Check if the part contains plain text
                if 'text/plain' in part['contentType']:
                    email_body = part['content']
                    break
        else:
            email_body = response.json()['value'][0]['body']['content']

        # decode the email body if it is base64 encoded
        if 'base64' in response.json()['value'][0]['body']:
            email_body = base64.b64decode(email_body).decode()

        # convert the email body to lowercase
        email_body = email_body.lower()

        # retrieve any attachment names.  Could potentially be used to check attachment patterns without having to actually fully download them
        # for message in result['value']:
        #     # Check if the message has attachments
        #     if message['hasAttachments']:
        #         # Retrieve the attachment names
        #         attachment_endpoint = f"/users/{o365_email_user_id}/messages/{message['id']}/attachments?$select=name"
        #         attachments = email_client.get(attachment_endpoint).json()
        #         attachment_names = [attachment['name'] for attachment in attachments['value']]

        # check if the message has attachments
        if message['hasAttachments']:
            # Retrieve the attachments
            attachment_endpoint = f"/users/{o365_email_user_id}/messages/{message['id']}/attachments"
            attachments = email_client.get(attachment_endpoint).json()

        # mark the email as read
        email_client.patch(f"/users/{o365_email_user_id}/messages/{email_id}", json={'isRead': True})

        # check if email meets any of the defined patterns
        for condition in email_rules['conditions']:
            condition_name = condition['name']
            if condition_name == 'example_entry_will_be_ignored':
                continue

            pattern = condition['pattern']

            meets_criteria = True
                
            # ensure the proper definitions are present
            if not ('attachments' in pattern or 'sender' in pattern or 'subject' in pattern or 'body' in pattern):
                meets_criteria = False
                continue

            # check sender pattern
            if 'sender' in pattern and pattern['sender'].lower() not in email_from:
                meets_criteria = False
                continue

            # check subject patterns
            if 'subject' in pattern:
                subject_patterns = pattern['subject']
                subject_matches = [pattern for pattern in subject_patterns if pattern.lower() in email_subject.lower()]
                if not subject_matches:
                    meets_criteria = False
                    continue

            # check body patterns
            if 'body' in pattern:
                body_patterns = pattern['body']
                body_matches = [pattern for pattern in body_patterns if pattern.lower() in email_body.lower()]
                if not body_matches:
                    meets_criteria = False
                    continue

            if 'delivery' in condition:
                delivery_target = condition['delivery']['target']
            else:
                delivery_target = None
            
            # check attachment pattern
            if meets_criteria and 'attachments' in pattern:
                attachment_matches = [pattern.lower() for pattern in pattern['attachments'][0]['filename']]
                for attachment in attachments['value']:
                    attachment_name = attachment['name'].lower()
                    if all(pattern in attachment_name for pattern in attachment_matches):
                        print(f'Attachment {attachment_name} meets the condition: {condition_name}')
                        if delivery_target != 'email_forward':
                            attachment_content = base64.b64decode(attachment['contentBytes'])
                            transmit_files(condition_name, delivery_target, condition['delivery'], email_date, attachment['name'], attachment_content)
                        any_attachment_matched = True
                if not any_attachment_matched:
                    meets_criteria = False
                    continue

            if meets_criteria and delivery_target == 'email_forward':
                forward_email(message, condition['delivery'])

            # this prevents the email from being compared against further patterns.  If you wish to have the email evaluated against other conditions, such as to extract other attachments, remove these lines
            if meets_criteria == True:
                break

# %%
