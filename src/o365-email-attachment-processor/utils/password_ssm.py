"""Retrieve password using Systems Manager Parameter Store"""
import boto3
from botocore.exceptions import ClientError

client = boto3.client(
    'ssm',
    endpoint_url='https://my-ssm-instance.example.com',
    region_name='your_region_name',
    aws_access_key_id='your_access_key',
    aws_secret_access_key='your_secret_key',
    aws_session_token='your_session_token'
)

def get_password(account_name, password_key, encoding='utf-8'):
    """Return password based on account name and secret key"""

    password_path = '/aws/prod/email_processor/passwords'
    parameter_name = f"{account_name}_{password_key}"
    if password_path is not None and password_path != '':
        parameter_name = f"{password_path}/{account_name}_{password_key}"
    try:
        response = client.get_parameter(Name=parameter_name, WithDecryption=True)
    except ClientError as e:
        print(f"Error retrieving secret: {e}")
        return None
    else:
        secret_value = response['Parameter']['Value']
        supported_encodings = ['utf-8', 'ascii', 'latin-1', 'utf-16']
        if encoding not in supported_encodings:
            print(f"Unsupported encoding: {encoding}")
            return None
        if isinstance(secret_value, bytes):
            secret_value = secret_value.decode(encoding)
    
    return secret_value
