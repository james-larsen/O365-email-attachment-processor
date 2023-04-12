"""Check multiple o365 servers for new emails, evaluate them against pre-defined patterns, and deliver any attachments accordingly"""
#%%
import os
# import o365lib
import email
import json
import base64
from pathlib import Path
# import keyring
import configparser
# pylint: disable=import-error
from utils.password import get_password as pw
# pylint: enable=import-error
import boto3
from botocore.exceptions import ClientError
from azure.identity import ClientSecretCredential
from msgraph.core import GraphClient
import datetime

#%%

def authenticate(tenant_id, client_id, client_secret):
    """Authenticate with the O365 server to return a client object"""
    # client_secret = pw.get_password(account_name, o365_password_key)

    # Create a ClientSecretCredential object
    credential = ClientSecretCredential(tenant_id=tenant_id, client_id=client_id, client_secret=client_secret)

    # Create a GraphClient object
    client = GraphClient(credential=credential)

    return client

def retrieve_rules(email_account_name):
    json_filename = f'{email_account_name}_email_rules.json'

    if os.path.exists(json_filename):
        with open(json_filename, encoding='utf-8') as f:
            config_data = json.load(f)
    else:
        with open("default_email_rules.json", encoding='utf-8') as f:
            config_data = json.load(f)
    
    return config_data

def transmit_files(condition_name, target, delivery_details, attachment_name):
    """Transmit files to an target location"""
    attachment_content = part.get_payload(decode=True)
    
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
    email_account = account['email_account']
    email_account_name = email_account['account_name']
    o365_email_username = email_account['o365_username']
    o365_email_user_id = email_account['o365_user_id']
    o365_email_tenant_id = email_account['o365_tenant_id']
    o365_email_client_id = email_account['o365_client_id']
    email_password_key = email_account['o365_password_key']
    o365_emailpassword = pw(email_account_name, email_password_key)
    email_client = authenticate(o365_email_tenant_id, o365_email_client_id, o365_emailpassword)

    sharepoint_account = account['sharepoint_account']
    sharepoint_account_name = sharepoint_account['account_name']
    o365_sharepoint_username = sharepoint_account['o365_username']
    o365_sharepoint__user_id = sharepoint_account['o365_user_id']
    o365_sharepoint_tenant_id = sharepoint_account['o365_tenant_id']
    o365_sharepoint_client_id = sharepoint_account['o365_client_id']
    sharepoint_password_key = sharepoint_account['o365_password_key']
    o365_sharepointpassword = pw(sharepoint_account_name, sharepoint_password_key)
    sharepoint_client = authenticate(o365_sharepoint_tenant_id, o365_sharepoint_client_id, o365_sharepointpassword)

    # user_principal_name = "206705394@tfayd.com"
    # response = email_client.get(f"/users/{user_principal_name}")
    # user = response.json()
    o365_email_user_id = "06091a12-c517-425f-aab8-e4842c01da14"
    # o365_email_user_id = "f1144887-d3fb-4ef9-911e-e6851753679e"

    # Read Rules
    config_data = retrieve_rules(email_account_name)

    search_parameters = {
        "$search": "isRead ne true",
        "$orderby": "receivedDateTime desc"
    }

    # Define the API endpoint for retrieving emails
    api_endpoint = f"/users/{o365_email_user_id}/messages"

    # Retrieve all unread emails
    result = email_client.get(api_endpoint, params=search_parameters).json()

    if len(result) == 0:
        print('No Unread messages to process')
        continue

    # Loop through all unread emails
    for message in result['value']:
        # Fetch the email content
        email_id = message['id']

        # Mark the email as read
        # client.patch(f"{api_endpoint}/{email_id}", json={'isRead': True})

        # Extract relevant fields from the email message
        email_subject = message['subject'].lower()
        email_from = message['from']['emailAddress']['address'].lower()
        email_to = [recipient['emailAddress']['address'].lower() for recipient in message['toRecipients']]
        email_date = message['receivedDateTime']

        response = email_client.get(f"{api_endpoint}/{email_id}", params={'$select': 'body'})
        email_body = ''

        # Check if the email is multipart
        if 'multipart' in response.json()['body']['contentType']:
            for part in response.json()['body']['content']:
                # Check if the part contains plain text
                if 'text/plain' in part['contentType']:
                    email_body = part['content']
                    break
        else:
            email_body = response.json()['body']['content']

        # Decode the email body if it is base64 encoded
        if 'base64' in response.json()['body']:
            email_body = base64.b64decode(email_body).decode()

        # Convert the email body to lowercase
        email_body = email_body.lower()

        # retrieve any attachments
        attachment_list = []
        for message in result['value']:
            # Check if the message has attachments
            if message['hasAttachments']:
                # Retrieve the attachments
                attachment_endpoint = f"/users/{o365_email_user_id}/messages/{message['id']}/attachments"
                attachments = email_client.get(attachment_endpoint).json()

                # Loop through all the attachments and append to the list
                for attachment in attachments['value']:
                    attachment_list.append(attachment)

        # attachments = []
        # if message['hasAttachments']:
        #     # email_id = message['id']
        #     attachments_endpoint = f"/users/{user_id}/messages/{email_id}/attachments"
        #     attachments = client.get(attachments_endpoint).json()

        #     # attachments is a list of dictionaries, each containing information about an attachment
        #     for attachment in attachments['value']:
        #         attachment_id = attachment['id']
        #         attachment_name = attachment['name']
        #         attachment_content_type = attachment['contentType']
        #         attachment_size = attachment['size']

        #         # Get the content of the attachment
        #         attachment_endpoint = f"/users/{user_id}/messages/{email_id}/attachments/{attachment_id}/$value"
        #         attachment_content = client.get(attachment_endpoint).content

        # check if email meets any of the defined patterns
        for condition in config_data['conditions']:
            condition_name = condition['name']
            if condition_name == 'example_entry_will_be_ignored':
                continue

            if 'delivery' in condition:
                delivery_target = condition['delivery']['target']
                # delivery_path = condition['delivery']['path']
            else:
                target = None
                # path = None

            pattern = condition['pattern']

            meets_criteria = True
                
            # ensure the proper definitions are present
            if not 'attachments' in pattern or not ('sender' in pattern or 'subject' in pattern or 'body' in pattern):
                meets_criteria = False
                continue

            # check sender pattern
            if 'sender' in pattern and pattern['sender'].lower() not in email_from:
                meets_criteria = False
                continue

            # check subject patterns
            if 'subject' in pattern:
                subject_patterns = pattern['subject']
                subject_matches = [pattern for pattern in subject_patterns if pattern.lower() in email_subject]
                if not subject_matches:
                    meets_criteria = False
                    continue

            # check body patterns
            if 'body' in pattern:
                body_patterns = pattern['body']
                body_matches = [pattern for pattern in body_patterns if pattern.lower() in email_body]
                if not body_matches:
                    meets_criteria = False
                    continue
            
            # check attachment pattern
            if meets_criteria and 'attachments' in pattern:
                any_attachment_matched = False
                for i, attachment in enumerate(attachments):
                    attachment_matches = [pattern.lower() for pattern in pattern['attachments'][0]['filename']]
                    if all(pattern in attachment.lower() for pattern in attachment_matches):
                        print(f'Attachment {i+1} meets the condition: {condition_name}')
                        if delivery_target:
                            transmit_files(condition_name, delivery_target, condition['delivery'], attachment)
                        any_attachment_matched = True
                if not any_attachment_matched:
                    meets_criteria = False
                    continue

                # mark email as read
                # o365_conn.store(email_id, '+FLAGS', '\\Flagged')

                # this prevents the email from being compared against other patterns.  If you wish to have the email evaluated against other conditions, such as to extract other attachments, remove these lines
                if meets_criteria == True:
                    break

    # close the o365 connection
    o365_conn.close()
    o365_conn.logout()
