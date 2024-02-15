import os
import time
import json
import urllib.parse
import requests
import subprocess
from docx2pdf import convert 
from pypdf import PdfWriter, PdfReader
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload

# Paths
training_folder = r"E:\Accountability\Training\2024"

docx_path = os.path.join(training_folder, "ThePRogram2024.docx")
pdf_path = os.path.join(training_folder, "ThePRogram2024.pdf")
training_pdf_path = os.path.join(training_folder, "Training.pdf")

google_token_path = os.path.join(training_folder, "TrainingBackupCredentials", "google_token.json")
credentials_path = os.path.join(training_folder, "TrainingBackupCredentials", "credentials.json")

onedrive_client_id = "d8e6ca27-806b-4368-aa7f-6df4db14976b"
onedrive_client_secret = "psu8Q~QyZ4odKo.CfBLwvWGkTZc1m5eXbGVKWasK" 

redirect_uri = "http://localhost:8080/"
token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"

# Hide the output of convert()
def convert(docx_path, pdf_path):
    # Redirect stdout and stderr to NUL
    with open('NUL', 'w') as nul:
        subprocess.run(['docx2pdf', docx_path, pdf_path], stdout=nul, stderr=subprocess.STDOUT)

def clean_local_folder(file_path):
    if os.path.exists(file_path):
        os.remove(file_path)   

# Authenticate with Google Drive
def authenticate_google_drive():
    SCOPES = ["https://www.googleapis.com/auth/drive"]
    credentials = None
    if os.path.exists(google_token_path):
        credentials = Credentials.from_authorized_user_file(google_token_path, SCOPES)

    if not credentials or not credentials.valid:
        flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
        credentials = flow.run_local_server(port=0)

        with open(google_token_path, "w") as token:
            token.write(credentials.to_json())

    return credentials

def google_drive_upload():
    try:
        credentials = authenticate_google_drive()
        service = build("drive", "v3", credentials=credentials)
        response = service.files().list(
            q="name='PRogram1' and mimeType='application/vnd.google-apps.folder'",
            spaces="drive"
        ).execute()

        folder_id = response["files"][0]["id"] if response["files"] else create_folder(service)

        response = service.files().list(
            q="name='ThePRogram2024.pdf' and '{}' in parents".format(folder_id),
            spaces="drive"
        ).execute()

        if response["files"]:
            service.files().delete(fileId=response["files"][0]["id"]).execute()

        media = MediaFileUpload(pdf_path)
        file_metadata = {"name": "ThePRogram2024.pdf", "parents": [folder_id]}
        upload_file = service.files().create(body=file_metadata, media_body=media, fields="id").execute()
        
        media.stream().close()

    except HttpError as e:
        print("Error: " + str(e))

# OneDrive
def authenticate_onedrive(): 
    redirect_uri = "http://localhost:8080/"  
    scope = "files.readwrite offline_access"
    onedrive_token_path = os.path.join(training_folder, "TrainingBackupCredentials", "onedrive_token.json")

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
        "client_id": onedrive_client_id,
        "redirect_uri": redirect_uri,
        "scope": scope,
        "response_type": "code"
    }

    print(auth_url + "?" + urllib.parse.urlencode(auth_params))
    url = input("Copy the redirected URL here:\n\n")
    code = urllib.parse.parse_qs(urllib.parse.urlparse(url).query).get('code', [''])[0]   

    access_token, expires_at, refresh_token = exchange_code_for_tokens(code)

    with open(onedrive_token_path, "w") as token_file:
        token_data = {"access_token": access_token, "expires_at": expires_at, "refresh_token": refresh_token}
        json.dump(token_data, token_file)
    
    return access_token

def onedrive_upload():
    access_token = authenticate_onedrive()
    try:
        reader = PdfReader(pdf_path)
        last_page = reader.pages[-1]
        writer = PdfWriter()
        writer.add_page(last_page)

        with open(os.path.join(training_folder, "Training.pdf"), "wb") as output_pdf:
            writer.write(output_pdf)

        upload_url = "https://graph.microsoft.com/v1.0/me/drive/root:/Training.pdf:/content"
        headers = {"Authorization": "Bearer " + access_token}
        with open(training_pdf_path, "rb") as file_content:
            response = requests.put(upload_url, headers=headers, data=file_content)

    except Exception as e:
        print("Error:", str(e))

# Shared functions
def create_folder(service):
    file_metadata = {"name": "PRogram1", "mimeType": "application/vnd.google-apps.folder"}
    file = service.files().create(body=file_metadata, fields="id").execute()
    return file.get("id")

def refresh_access_token(refresh_token): 
    token_params = {
        "client_id": onedrive_client_id,
        "client_secret": onedrive_client_secret,
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

def exchange_code_for_tokens(code):
    token_params = {
        "client_id": onedrive_client_id,
        "client_secret": onedrive_client_secret,
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

# Magic happens after this comment: 
convert(docx_path, pdf_path)

google_drive_upload()
onedrive_upload()

clean_local_folder(pdf_path)
clean_local_folder(training_pdf_path)


