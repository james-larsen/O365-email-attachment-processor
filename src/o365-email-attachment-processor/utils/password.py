"""Retrieve password"""
import keyring

def get_password(account_name, password_key):
    """Return password based on username and secret key"""
    
    return keyring.get_password(account_name, password_key)
