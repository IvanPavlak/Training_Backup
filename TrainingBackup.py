import os
import os.path
import time
import json
import urllib
import socket
import requests
from docx2pdf import convert
from datetime import datetime
from pypdf import PdfWriter, PdfReader
from urllib.parse import urlparse, parse_qs
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# Get the hostname of the system
hostname = socket.gethostname()

# Read the configuration file
with open("configuration.json") as f:
    config = json.load(f)

# Get the paths based on the hostname
if hostname in config:
    paths = config[hostname]
    training_folder = paths["training_folder"]
    credentials_folder = paths["credentials_folder"]
else:
    print("Hostname not found in configuration file!")

docx_path = os.path.join(training_folder, "ThePRogram2024.docx")
pdf_path = os.path.join(training_folder, "ThePRogram2024.pdf")
training_pdf_path = os.path.join(training_folder, "Training.pdf")

google_token_path = os.path.join(credentials_folder, "google_token.json")
google_credentials_path = os.path.join(credentials_folder, "google_credentials.json")

onedrive_token_path = os.path.join(credentials_folder, "onedrive_token.json")
one_drive_credentials_path = os.path.join(credentials_folder, "onedrive_credentials.json")

# Global Variables
with open(one_drive_credentials_path, "r") as f:
    credentials = json.load(f)

onedrive_client_id = credentials.get("client_id")
onedrive_client_secret = credentials.get("client_secret")
redirect_uri = "http://localhost:8080/"
token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"

# Convert .docx to .pdf
print('\nConverting ".docx" to ".pdf":')
convert(docx_path, pdf_path)
print()

# Delete local files
def clean_local_folder(file_path):
    """
    Deletes a file from the local folder if it exists.
    
    Args:
        file_path (str): Path of the file to be deleted.
    """
    file = os.path.basename(file_path)
    if os.path.exists(file_path):
        os.remove(file_path)
        print(f'"{file}" deleted successfully from local folder!')  
    else:
        print(f'"{file}" does not exist in the specified path!')        

# One Drive Authentication
def authenticate_onedrive(): 
    """
    Authenticates with OneDrive using OAuth2.
    If tokens are expired or missing, it initiates the authentication process.

    Returns:
        str: Access token for OneDrive API.
    """
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
        "client_id": onedrive_client_id,
        "redirect_uri": redirect_uri,
        "scope": scope,
        "response_type": "code"
    }
    
    print("\nClick on this link to authenticate with OneDrive:\n")
    print(auth_url + "?" + urllib.parse.urlencode(auth_params))
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
    """
    Exchanges the authorization code for access and refresh tokens.

    Args:
        code (str): Authorization code received from the OAuth2 flow.

    Returns:
        tuple: Access token, expiry time, and refresh token.
    """
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

def refresh_access_token(refresh_token): 
    """
    Refreshes the access token using the refresh token.

    Args:
        refresh_token (str): Refresh token for obtaining a new access token.

    Returns:
        tuple: New access token, expiry time, and refresh token.
    """
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

# One Drive Upload
def upload_to_onedrive(access_token):
    """
    Uploads the PDF and DOCX files to a specified folder in OneDrive.

    Args:
        access_token (str): Access token for OneDrive API.
    """
    try:
        reader = PdfReader(pdf_path)
        last_page = reader.pages[-1]

        writer = PdfWriter()
        writer.add_page(last_page)

        with open(training_pdf_path, "wb") as output_pdf:
            writer.write(output_pdf)

        files_to_upload = {
            "Training.pdf": training_pdf_path,
            "ThePRogram2024.docx": docx_path
        }

        for file_name, file_path in files_to_upload.items():
            upload_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/Training/{file_name}:/content"
            headers = {"Authorization": "Bearer " + access_token}
            file_content = open(file_path, "rb")
            response = requests.put(upload_url, headers=headers, data=file_content)
            file_content.close()

            if response.status_code == 200: 
                print(f'=> Uploaded "{file_name}" to OneDrive!\n')
            else: 
                print(f'Error uploading "{file_name}" to OneDrive:\n\n {response.text}\n')

    except Exception as e:
        print(f"Error:\n\n {str(e)}")

# Google Drive Authentication
def authenticate_google_drive():
    """
    Authenticates with Google Drive using OAuth2.

    Returns:
        google.oauth2.credentials.Credentials: Google API credentials.
    """
    SCOPES = ["https://www.googleapis.com/auth/drive"]
    credentials = None

    if os.path.exists(google_token_path):
        try:
            credentials = Credentials.from_authorized_user_file(google_token_path, SCOPES)
        except Exception as e:
            print(f"Error loading credentials: {e}")
            credentials = None

    if not credentials or not credentials.valid:
        try:
            if credentials and credentials.expired and credentials.refresh_token:
                response = requests.post(
                    "https://oauth2.googleapis.com/token",
                    data={
                        "client_id": credentials.client_id,
                        "client_secret": credentials.client_secret,
                        "refresh_token": credentials.refresh_token,
                        "grant_type": "refresh_token"
                    }
                )
                response.raise_for_status()
                refresh_data = response.json()
                
                credentials.token = refresh_data.get("access_token")
                credentials.expiry = datetime.fromtimestamp(
                    time.time() + refresh_data.get("expires_in", 3600)
                )
            else:
                flow = InstalledAppFlow.from_client_secrets_file(google_credentials_path, SCOPES)
                credentials = flow.run_local_server(port=0)

            with open(google_token_path, "w") as token:
                token.write(credentials.to_json())

        except Exception as e:
            print(f"Authentication error: {e}")
            flow = InstalledAppFlow.from_client_secrets_file(google_credentials_path, SCOPES)
            credentials = flow.run_local_server(port=0)

            with open(google_token_path, "w") as token:
                token.write(credentials.to_json())

    return credentials

# Google Drive Upload
def upload_to_google_drive(credentials):
    """
    Uploads a PDF file to specified folders in the Google Drive.
    If there already is a file with the same name, it is deleted before the upload.
    - this deletion is necessary because two files with the same name can be uploaded to Google Drive, which results in unnecessary clutter.

    Args:
        credentials: Google API credentials.
    """
    folders = ["PRogram1", "PRogram2", "10xEngineers"]

    try:
        service = build("drive", "v3", credentials=credentials)

        for folder_name in folders:
            response = service.files().list(
                q=f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder'",
                spaces="drive"
            ).execute()

            if not response["files"]:
                file_metadata = {
                    "name": folder_name,
                    "mimeType": "application/vnd.google-apps.folder"
                }
                file = service.files().create(body=file_metadata, fields="id").execute()
                folder_id = file.get("id")
            else:
                folder_id = response["files"][0]["id"]

            response = service.files().list(
                q=f"name='ThePRogram2024.pdf' and parents='{folder_id}'",
                spaces="drive"
            ).execute()
            for existing_file in response.get("files", []):
                existing_file_id = existing_file["id"]
                service.files().delete(fileId=existing_file_id).execute()
                print(f'Deleted existing file "{existing_file["name"]}" from folder "{folder_name}"')

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
            print(f'=> Uploaded file to Google Drive folder "{folder_name}"\n')

    except HttpError as e:
        print(f"Error:\n\n {str(e)}")

upload_to_onedrive(authenticate_onedrive())
upload_to_google_drive(authenticate_google_drive())
clean_local_folder(training_pdf_path)
clean_local_folder(pdf_path)
print()