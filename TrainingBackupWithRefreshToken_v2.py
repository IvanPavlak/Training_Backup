import os
import os.path
import time
import json
import urllib
import requests
from getpass import getpass
from docx2pdf import convert 
from pypdf import PdfWriter, PdfReader
from urllib.parse import urlparse, parse_qs
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow


# Convert .docx to .pdf
print("--------------------------------------------------------------------------------")
print('Converting ".docx" to ".pdf":\n')
docx_path = r"E:\Accountability\Training\2024\ThePRogram2024.docx"  
pdf_path = r"E:\Accountability\Training\2024\ThePRogram2024.pdf"
convert(docx_path, pdf_path)

credentials_path = r"E:\Accountability\Training\2024\TrainingBackupCredentials\credentials.json"
google_token_path = r"E:\Accountability\Training\2024\TrainingBackupCredentials\google_token.json"
onedrive_token_path = r"E:\Accountability\Training\2024\TrainingBackupCredentials\onedrive_token.json"

training_pdf_path = r"E:\Accountability\Training\2024\Training.pdf"

def clean_local_folder(file_path):
    file = os.path.basename(file_path)
    if os.path.exists(file_path):
        os.remove(file_path)
        print(f'"{file}" deleted successfully from local folder!')
        print("--------------------------------------------------------------------------------")
    else:
        print(f'"{file}" does not exist in the specified path!')
        print("--------------------------------------------------------------------------------")

# Authenticate with One Drive
def authenticate_onedrive():
    client_id = "d8e6ca27-806b-4368-aa7f-6df4db14976b"  
    redirect_uri = "http://localhost:8080/"  
    scope = "files.readwrite offline_access"

    if os.path.exists(onedrive_token_path):
        with open(onedrive_token_path, "r") as token_file:
            token_data = json.load(token_file)
            access_token = token_data.get("access_token")
            expires_at = token_data.get("expires_at")
            refresh_token = token_data.get("refresh_token")

            if access_token and expires_at > time.time():
                return access_token

            if refresh_token:
                access_token, expires_at, refresh_token = refresh_access_token(refresh_token)
                if access_token:
                    return access_token

    auth_url = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
    auth_params = {
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "scope": scope,
        "response_type": "code"
    }
    
    print("--------------------------------------------------------------------------------")
    print("Click on this link to authenticate with OneDrive:\n")
    print(auth_url + "?" + urllib.parse.urlencode(auth_params))
    
    print("--------------------------------------------------------------------------------")
    url = input("Copy the redirected URL here:\n\n")
    code = parse_qs(urlparse(url).query).get('code', [''])[0]   

    access_token, expires_at, refresh_token = exchange_code_for_tokens(code)

    with open(onedrive_token_path, "w") as token_file:
        token_data = {
            "access_token": access_token,
            "expires_at": expires_at,
            "refresh_token": refresh_token
        }
        json.dump(token_data, token_file)
    
    return access_token

def exchange_code_for_tokens(code):
    client_id = "d8e6ca27-806b-4368-aa7f-6df4db14976b"
    client_secret = "psu8Q~QyZ4odKo.CfBLwvWGkTZc1m5eXbGVKWasK" 
    redirect_uri = "http://localhost:8080/"
    token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    
    token_params = {
        "client_id": client_id,
        "client_secret": client_secret,
        "code": code,
        "redirect_uri": redirect_uri,
        "grant_type": "authorization_code"
    }
    
    response = requests.post(token_url, data=token_params)
    token_data = response.json()
    
    access_token = token_data.get("access_token")
    expires_in = token_data.get("expires_in", 3600)
    expires_at = time.time() + expires_in
    refresh_token = token_data.get("refresh_token")
    
    return access_token, expires_at, refresh_token

def refresh_access_token(refresh_token):
    client_id = "d8e6ca27-806b-4368-aa7f-6df4db14976b"
    client_secret = "psu8Q~QyZ4odKo.CfBLwvWGkTZc1m5eXbGVKWasK" 
    token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    
    token_params = {
        "client_id": client_id,
        "client_secret": client_secret,
        "refresh_token": refresh_token,
        "grant_type": "refresh_token"
    }
    
    response = requests.post(token_url, data=token_params)
    token_data = response.json()
    
    access_token = token_data.get("access_token")
    expires_in = token_data.get("expires_in", 3600)
    expires_at = time.time() + expires_in
    new_refresh_token = token_data.get("refresh_token", refresh_token)
    
    return access_token, expires_at, new_refresh_token

access_token = authenticate_onedrive()

# OneDrive Upload
try:
    reader = PdfReader(pdf_path)
    last_page = reader.pages[-1]

    writer = PdfWriter()
    writer.add_page(last_page)

    with open(training_pdf_path, "wb") as output_pdf:
        writer.write(output_pdf)

    upload_url = "https://graph.microsoft.com/v1.0/me/drive/root:/Training.pdf:/content"
    headers = {"Authorization": "Bearer " + access_token}
    file_content = open(training_pdf_path, "rb")
    response = requests.put(upload_url, headers=headers, data=file_content)
    file_content.close()

    if response.status_code == 200:
        print("--------------------------------------------------------------------------------")
        print("Uploaded file to OneDrive!")
    else:
        print("--------------------------------------------------------------------------------")
        print("Error uploading file to OneDrive:", response.text)

except Exception as e:
    print("--------------------------------------------------------------------------------")
    print("Error:", str(e))

# Authenticate with Google Drive
def authenticate_google_drive():
    SCOPES = ["https://www.googleapis.com/auth/drive"]
    credentials = None
    if os.path.exists(google_token_path):
        credentials = Credentials.from_authorized_user_file(google_token_path, SCOPES)

    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            credentials = flow.run_local_server(port=0)

        with open(google_token_path, "w") as token:
            token.write(credentials.to_json())

    return credentials

# Google Drive Upload
try:
    credentials = authenticate_google_drive()
    service = build("drive", "v3", credentials=credentials)

    response = service.files().list(
        q="name='PRogram1' and mimeType='application/vnd.google-apps.folder'",
        spaces="drive"
    ).execute()

    if not response["files"]:
        file_metadata = {
            "name": "PRogram1",
            "mimeType": "application/vnd.google-apps.folder"
        }

        file = service.files().create(body=file_metadata, fields="id").execute()

        folder_id = file.get("id")
    else:
        folder_id = response["files"][0]["id"]

    response = service.files().list(
        q="name='ThePRogram2024.pdf' and '" + folder_id + "' in parents",
        spaces="drive"
    ).execute()

    if response["files"]:
        file_id = response["files"][0]["id"]
        service.files().delete(fileId=file_id).execute()
        print("--------------------------------------------------------------------------------")
        print("Deleted existing file from Google Drive!")

    file_name = "ThePRogram2024.pdf"
    file_metadata = {
        "name": file_name,
        "parents": [folder_id]
    }

    media = MediaFileUpload(pdf_path)
    upload_file = service.files().create(body=file_metadata,
                                         media_body=media,
                                         fields="id").execute()
    
    media.stream().close()
    print("--------------------------------------------------------------------------------")
    print("Uploaded file to Google Drive!")

except HttpError as e:
    print("--------------------------------------------------------------------------------")
    print("Error: " + str(e))

print("--------------------------------------------------------------------------------")
clean_local_folder(training_pdf_path)
clean_local_folder(pdf_path)

time.sleep(5)
