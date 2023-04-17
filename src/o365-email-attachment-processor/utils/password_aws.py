"""Retrieve password using AWS Secret Manager"""
import boto3
from botocore.exceptions import ClientError

client = boto3.client('secretsmanager')

def get_password(account_name, password_key):
    """Return password based on account name and secret key"""

    secret_name = f"{account_name}_{password_key}"
    try:
        response = client.get_secret_value(SecretId=secret_name)
    except ClientError as e:
        print(f"Error retrieving secret: {e}")
        return None
    else:
        secret_value = response['SecretString']
        return secret_value
