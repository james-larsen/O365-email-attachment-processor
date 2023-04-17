"""Retrieve password using Systems Manager Parameter Store"""
import boto3
from botocore.exceptions import ClientError

client = boto3.client('ssm')

def get_password(account_name, password_key):
    """Return password based on account name and secret key"""

    password_path = '/aws/prod/email_processor/passwords' # remove if passwords are stored in the root
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
        return secret_value
