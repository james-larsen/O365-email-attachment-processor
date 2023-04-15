# O365 Email Attachment Processor

The purpose of this application is to allow people to set up a number of o365 email accounts, login to them periodically, and check unread emails against a set of rules, including sender, subject, body, and attachment patterns, as well as delivery instructions.  When a match is found, the attachment is routed to a specified location locally or to an S3 bucket.  The rule can also be configured to forward the email to multiple recipients.  New o365 accounts and patterns can be added using JSON files.  Patterns and delivery instructions can also be provided via JSON files hosted locally and .xslx files hosted on Sharepoint.

## Requirements

python = "^3.8"

boto3 = "^1.26.45"

azure-identity = "^1.12.0"

msgraph-core = "^0.2.2"

openpyxl = "^3.1.2"

keyring = "^23.13.1" # Optional

## Installation

### Via poetry (Installation instructions [here](https://python-poetry.org/docs/)):

```python
poetry install
```

### Via pip:
```python
pip install boto3
pip install azure-identity
pip install msgraph-core
pip install openpyxl
pip install keyring # Optional
```

## Usage

Configure the following files:

* ./src/o365-email-attachment-processor/o365_accounts.json
* ./src/o365-email-attachment-processor/default_email_rules.json

Optionally add rules .xlsx files to a Sharepoint location

```python
python3 src/o365-email-attachment-processor/main.py
```

## Passwords

The module for retrieving database passwords is located at **'./src/o365-email-attachment-processor/utils/password.py'**.  By default it uses the 'keyring' library, accepts two strings of 'secret_key' and 'user_name' and returns a string of 'password'.  If you wish to use a different method of storing and retrieving database passwords, modify this .py file.

If you require more significant changes to how the password is retrieved (Eg. need to pass a different number of parameters), it is called by the **'./src/o365-email-attachment-processor/main.py'** module.

If you do wish to use the keyring library, create the below password entries ("account_name" is specified in the **o365_accounts.json** file below, "name" is specified under a given condition in the **_email_rules.json** files below):

* For each o365 account:
    * account_name, o365_password_key

* For each S3 delivery target:
    * name, "S3AccessKey"
    * name, "S3SecretKey"

## Account Permissions

You will need to make sure to register an app in Azure Active Directory to have a Tenant ID, Client ID, and Secret Key provided.  They should also provide the following permissions for the service account:

Email Permissions:
```
MailReadWrite
MailSend
User.Read.All
```

Sharepoint Permissions:
```
User.Read
Sites.ReadWrite.All
Files.ReadWrite.All
Sites.Selected
```

## App Configuration

The application is primarily controlled by .json files.  The first is:

**./src/o365-email-attachment-processor/o365_accounts.json**

Contains the details for each o365 account to be connected to.  It looks like the below:

``` json
{
    "o365_accounts": [
        {
            "email_account": {
                "account_name": "XYZ_Department_Sales_Email",
                "o365_username": "user1",
                "o365_user_id": "xxxxxxxx-xxxx-xxxx-xxxxxxxxxxxx",
                "o365_tenant_id": "xxxxxxxx-xxxx-xxxx-xxxxxxxxxxxx",
                "o365_client_id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
                "o365_password_key": "password_key"
            },
            "sharepoint_account": {
                "account_name": "XYZ_Department_Sales_SharePoint",
                "o365_username": "user2",
                "o365_site_address": "mycompany.sharepoint.com",
                "o365_site_name": "MySharepointSite",
                "o365_site_folderpath": "Documents/General/My Folder",
                "o365_user_id": "xxxxxxxx-xxxx-xxxx-xxxxxxxxxxxx",
                "o365_tenant_id": "xxxxxxxx-xxxx-xxxx-xxxxxxxxxxxx",
                "o365_client_id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
                "o365_password_key": "password_key"
            }
        }
    ]
}
```

### Email:
* ***account_name***:  General name for the file processing area.  It could be named based on the email address being monitored, or the department the files are being delivered for, etc.
* ***o365_username***:  Login username for o365 server
* ***o365_user_id***:  o365 email account internal ID
* ***o365_tenant_id***:  o365 Tenant ID
* ***o365_client_id***:  o365 Client ID
* ***o365_password_key***:  Key to be used along with "account_name" for retrieving the correct Secret Key

### Sharepoint (optional):
* ***account_name***:  General name for the file processing area.  Can be the same as email account_name
* ***o365_username***:  Login username for o365 server
* ***o365_site_address***:  Address for your Sharepoint site
* ***o365_site_name***:  Site name containing the rules folder
* ***o365_site_folderpath***:  Full path to the rules folder.  Must start with "Documents"
* ***o365_user_id***:  o365 Sharepoint account internal ID
* ***o365_tenant_id***:  o365 Tenant ID
* ***o365_client_id***:  o365 Client ID
* ***o365_password_key***:  Key to be used along with "account_name" for retrieving the correct Secret Key

**Note:**  The Sharepoint functionality is the most likely to have issues, depending on how your site is configured.  Knowledge of the "msgraph.core.GraphClient" library my be needed to point the application to the correct folder

---

**./src/o365-email-attachment-processor/default_email_rules.json**

Holds the specific patterns to look for and where to deliver the files when they are detected.

``` json
{
  "conditions": [
    {
      "name": "example_entry_will_be_ignored",
      "pattern": {
        "sender": "@domain",
        "subject": ["Pattern01", "Pattern02"],
        "body": ["Pattern01", "Pattern02"],
        "attachments": [
          {
            "filename": [".csv"]
          }
        ]
      },
      "delivery": {
        "target": "local",
        "path": "/path/to/save/files", 
        "append_datetime": "True"
      },
      "delivery": {
        "target": "s3", 
        "region": "us-west-1",
        "bucket": "my-bucket-name", 
        "subfolder": "sub-folder1/sub-folder2/", 
        "append_datetime": ""
      },
      "delivery": {
        "target": "email_forward",
        "recipients": ["email@server.com"],
        "body": "This is a custom email body"
      }
    }
  ]
}

```
**Note: Multiple "delivery" keys shown for reference, but only one should be used per condition entry**

* ***name***:  Name for the condition being defined
* ***pattern***
    * ***sender***:  Single string pattern to check against the "Sender" field
    * ***subject***:  List of strings to check against the "Subject" field
    * ***body***:  List of strings to check against the "Body" field
    * ***filename***:  List of strings to check against the "Filename" field of each attachment
* ***delivery***
    * ***target***:  Delivery target type (local, s3 or email_forward)
    * ***append_datetime***:  Whether to append the email datetime to the end of the attachment.  Accepts "True", anything else will be evaluated to False.  Date will be in the format of "_YYYY-MM-DD_HHMISS" in the UTC timezone
    * ***path***:  Local file path to deliver attachments
    * ***region***:  S3 bucket region
    * ***bucket***:  S3 bucket name
    * ***subfolder***:  Subfolder(s) to deliver within bucket (optional)
    * ***recipients***:  Email recipients for the email to be forwarded to
    * ***body***:  Custom body text of the forwarded email (optional)

**Notes:**

* The first entry "example_entry_will_be_ignored" will be ignored by the program.  Leaving this here may make it easier to refer to the proper syntax and which patterns are available
* This file is required.  If you plan to solely use the Sharepoint method discussed below, still leave the default "example_entry_will_be_ignored" entry
* For items defined as lists above, leave them in [], even if only a single pattern is desired
* Every condition must have atleast one of "sender", "subject", "body" or "filename" sections defined.  Ideally multiple should be defined to avoid a rule being applied to an incorrect email
* If the body is of type "multipart/alternative", the "text/plain" version will be preferred over the "text/html" version
* Currently the program can only deliver files locally and to an S3 bucket, or to forward the email.  Eventually the program will be enhanced to deliver to other locations (FTP Servers, Sharepoint, etc.)

---

**Account-Specific Conditions**

By default all accounts will use the **default_email_rules.json** file to determine their conditions.  However, you can create account-specific condition files.  While processing a particular account, the application will look for a file with the pattern "**{account_name}_email_rules.json**" and use it instead of the default.  Note that this uses the "account_name" from the "email_account" section.

For example, if you have an account_name in your "o365_accounts.json" with a value of "Acme_Marketing_Files", then it will look for **Acme_Marketing_Files_email_rules.json**.

---

### **Sharepoint-Hosted Rules**

In additional to the local JSON files to define rules, you can specify a Sharepoint location to host multiple Excel (.xlsx) files containing rules.  This is useful if you'd like to expose certain rules documents to business users to maintain themselves.  These rules will be appended to the JSON rules.  The document structure is outlined below, and a sample file can be found at **./templates/Sample Email Rules.xlsx**.

| *Name* | *Sender* | *Subject* | *Body* | *Filename* | *Recipients* | *Custom Email Body* |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
|My email rule|@company.com|Sales file for today |Here is today's \| contact your sales rep|daily_sales\|.csv|name@company.com \| name2@company.com|Attached is today's forwarded sales file|

**Notes:**

* You can add additional sheets to the Excel document, but don't rename the "Email Rules" sheet.  Similarly, you can add additional columns to the right of "Custom Email Body", but don't change the order of the existing columns
* Do not include any .xlsx files in this folder that are not rules documents
* Sender, Subject, Body and Filename are search text patterns, similar to those described above
* For Subject, Body, Filename and Recipients, multiple patterns can be represented with a pipe ("|"), a line break (Alt + Enter in cell), or both.  Sender only allows a single pattern
* "Custom Email Body" is an optional field.  If left blank, the email will be forwarded with no body text
* Text is not case sensitive

[## Logging]:#

[TBD]:#

## About the Author

My name is James Larsen, and I have been working professionally as a Business Analyst, Database Architect and Data Engineer since 2007.  While I specialize in Data Modeling and SQL, I am working to improve my knowledge in different data engineering technologies, particularly Python.

[https://www.linkedin.com/in/jameslarsen42](https://www.linkedin.com/in/jameslarsen42/)  
[https://github.com/james-larsen](https://github.com/james-larsen)