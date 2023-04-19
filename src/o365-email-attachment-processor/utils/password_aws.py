"""Retrieve password using AWS Secret Manager"""
import boto3
from botocore.exceptions import ClientError
import base64

client = boto3.client(
    'secretsmanager',
    endpoint_url='https://my-secret-manager-instance.example.com',
    region_name='your_region_name',
    aws_access_key_id='your_access_key',
    aws_secret_access_key='your_secret_key',
    aws_session_token='your_session_token'
)

def get_password(account_name, password_key, encoding='utf-8'):
    """Return password based on account name and secret key"""
    
    secret_name = f"{account_name}_{password_key}"
    
    try:
        response = client.get_secret_value(SecretId=secret_name)
    except ClientError as e:
        print(f"Error retrieving secret: {e}")
        return None
    else:
        if 'SecretString' in response:
            secret_value = response['SecretString']
        else:
            supported_encodings = ['utf-8', 'ascii', 'latin-1', 'utf-16']
            
            if encoding not in supported_encodings:
                print(f"Unsupported encoding: {encoding}")
                return None
            
            secret_value = base64.b64decode(response['SecretBinary']).decode(encoding)
        return secret_value
