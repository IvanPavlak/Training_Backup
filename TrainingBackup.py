import os
import os.path
import time
import json
import urllib
import socket
import requests
import platform
import subprocess
from docx2pdf import convert
from datetime import datetime
from pypdf import PdfWriter, PdfReader
from urllib.parse import urlparse, parse_qs, quote
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

RESET = "\033[0m"
DARK_CYAN = "\033[36m"
GREEN = "\033[92m"
RED = "\033[91m"

hostname = socket.gethostname()

with open("configuration.json") as f:
    config = json.load(f)

if hostname in config:
    paths = config[hostname]
    training_folder = paths["training_folder"]
    credentials_folder = paths["credentials_folder"]
else:
    print(RED + "=> Hostname not found in configuration file!" + RESET)
    exit("Exiting due to configuration error.")

docx_path = os.path.join(training_folder, "ThePRogram2025.docx")
pdf_path = os.path.join(training_folder, "ThePRogram2025.pdf")
training_pdf_path = os.path.join(training_folder, "Training.pdf")

google_token_path = os.path.join(credentials_folder, "google_token.json")
google_credentials_path = os.path.join(credentials_folder, "google_credentials.json")

onedrive_token_path = os.path.join(credentials_folder, "onedrive_token.json")
one_drive_credentials_path = os.path.join(credentials_folder, "onedrive_credentials.json")

with open(one_drive_credentials_path, "r") as f:
    onedrive_creds_json = json.load(f)

onedrive_client_id = onedrive_creds_json.get("client_id")
onedrive_client_secret = onedrive_creds_json.get("client_secret")
redirect_uri = "http://localhost:8080/"
token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"

def clean_local_folder(file_path):
    file_name = os.path.basename(file_path)
    if os.path.exists(file_path):
        os.remove(file_path)
        print(GREEN + f'=> "{file_name}" deleted successfully from local folder!' + RESET)
    else:
        print(f'"{file_name}" does not exist in the specified path to delete!')

def authenticate_onedrive():
    scope = "files.readwrite offline_access"
    if os.path.exists(onedrive_token_path):
        with open(onedrive_token_path, "r") as token_file:
            token_data = json.load(token_file)
            access_token = token_data.get("access_token")
            expires_at = token_data.get("expires_at")
            refresh_token = token_data.get("refresh_token")

            if access_token and expires_at and expires_at > time.time():
                print(GREEN + "=> One Drive token authenticated successfully!" + RESET)
                return access_token

            if refresh_token:
                print("One Drive token expired or invalid, attempting to refresh...")
                refreshed_access_token, new_expires_at, new_refresh_token = refresh_access_token(refresh_token)
                if refreshed_access_token:
                    print(GREEN + "=> One Drive token refreshed successfully!" + RESET)
                    with open(onedrive_token_path, "w") as new_token_file:
                        new_token_data = {
                            "access_token": refreshed_access_token,
                            "expires_at": new_expires_at,
                            "refresh_token": new_refresh_token
                        }
                        json.dump(new_token_data, new_token_file)
                    return refreshed_access_token
                else:
                    print(RED + "=> Failed to refresh One Drive token. Proceeding to full authentication!" + RESET)

    auth_url = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
    auth_params = {
        "client_id": onedrive_client_id,
        "redirect_uri": redirect_uri,
        "scope": scope,
        "response_type": "code"
    }

    print("\nClick on this link to authenticate with One Drive:\n")
    print(auth_url + "?" + urllib.parse.urlencode(auth_params))
    url_input = input("Copy the redirected URL here:\n\n")
    code = parse_qs(urlparse(url_input).query).get('code', [''])[0]

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
    token_params = {
        "client_id": onedrive_client_id,
        "client_secret": onedrive_client_secret,
        "refresh_token": refresh_token,
        "grant_type": "refresh_token",
        "scope": "files.readwrite offline_access"
    }
    response = requests.post(token_url, data=token_params)
    if response.status_code == 200:
        token_data = response.json()
        access_token = token_data.get("access_token")
        expires_in = token_data.get("expires_in", 3600)
        expires_at = time.time() + expires_in
        new_refresh_token = token_data.get("refresh_token", refresh_token)
        return access_token, expires_at, new_refresh_token
    else:
        print(RED + f"=> Error refreshing One Drive token: {response.status_code} - {response.text}" + RESET)
        return None, None, None

def download_file_from_onedrive(access_token, onedrive_folder_name, onedrive_file_name, local_file_path):
    local_dir = os.path.dirname(local_file_path)
    if local_dir and not os.path.exists(local_dir):
        os.makedirs(local_dir, exist_ok=True)
        print(f"Created directory: {local_dir}")

    item_path = f"/{onedrive_folder_name}/{onedrive_file_name}"
    encoded_item_path = quote(item_path)

    download_url = f"https://graph.microsoft.com/v1.0/me/drive/root:{encoded_item_path}:/content"
    headers = {"Authorization": "Bearer " + access_token}

    response = requests.get(download_url, headers=headers, stream=True)

    if response.status_code == 200:
        with open(local_file_path, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        print(GREEN + f'=> Downloaded "{onedrive_file_name}" to "{local_file_path}" successfully!' + RESET)
        return True
    else:
        print(RED + f'=> Error downloading "{onedrive_file_name}" from One Drive: {response.status_code}' + RESET)
        if response.text:
            try:
                error_details = response.json()
                if error_details.get("error", {}).get("code") == "itemNotFound":
                    print(RED + f'=> File "{onedrive_file_name}" not found in OneDrive folder "{onedrive_folder_name}".' + RESET)
                else:
                    print(RED + f'Response: {response.text}\n' + RESET)
            except json.JSONDecodeError:
                 print(RED + f'Response: {response.text}\n' + RESET)
        return False


def delete_file_from_onedrive(access_token, onedrive_folder_name, onedrive_file_name):
    """
    Deletes a file from a specified folder in One Drive.
    """
    item_path = f"/{onedrive_folder_name}/{onedrive_file_name}"
    encoded_item_path = quote(item_path)
    delete_url = f"https://graph.microsoft.com/v1.0/me/drive/root:{encoded_item_path}"
    headers = {"Authorization": "Bearer " + access_token}

    response = requests.delete(delete_url, headers=headers)

    if response.status_code == 204:
        print(GREEN + f'=> Successfully deleted "{onedrive_file_name}" from One Drive!' + RESET)
        return True
    elif response.status_code == 404:
        print(f'File "{onedrive_file_name}" not found on One Drive for deletion!')
        return True
    else:
        print(RED + f'=> Error deleting "{onedrive_file_name}" from One Drive: {response.status_code}' + RESET)
        print(RED + f'URL: {delete_url}' + RESET)
        print(RED + f'Response: {response.text}\n' + RESET)
        return False

def upload_to_onedrive(access_token):
    try:
        files_to_upload = {}
        if not os.path.exists(pdf_path):
            print(RED + f'=> "{pdf_path}" not found. Cannot create "Training.pdf" for One Drive upload!' + RESET)
            if os.path.exists(docx_path):
                files_to_upload = {"ThePRogram2025.docx": docx_path}
                print("Proceeding to upload only ThePRogram2025.docx to One Drive!")
            else:
                print(RED + "=> Neither PDF nor DOCX found for One Drive upload. Skipping!" + RESET)
                return
        else:
            reader = PdfReader(pdf_path)
            if not reader.pages:
                print(RED + f'=> "{pdf_path}" contains no pages. Cannot create "Training.pdf"!' + RESET)
                if os.path.exists(docx_path):
                    files_to_upload = {"ThePRogram2025.docx": docx_path}
                    print("Proceeding to upload only ThePRogram2025.docx to One Drive!")
                else:
                    print(RED + "=> DOCX also not found. Skipping One Drive upload!" + RESET)
                    return
            else:
                last_page = reader.pages[-1]
                writer = PdfWriter()
                writer.add_page(last_page)
                with open(training_pdf_path, "wb") as output_pdf:
                    writer.write(output_pdf)

                files_to_upload = {
                    "Training.pdf": training_pdf_path,
                    "ThePRogram2025.docx": docx_path
                }

        for file_name, current_file_path in files_to_upload.items():
            if not os.path.exists(current_file_path):
                print(f'File "{current_file_path}" for One Drive upload as "{file_name}" not found. Skipping this file!')
                continue

            item_path = f"/Training/{file_name}"
            encoded_item_path = quote(item_path)
            upload_url = f"https://graph.microsoft.com/v1.0/me/drive/root:{encoded_item_path}:/content"

            headers = {"Authorization": "Bearer " + access_token, "Content-Type": "application/octet-stream"}

            with open(current_file_path, "rb") as file_content:
                response = requests.put(upload_url, headers=headers, data=file_content)

            if response.status_code == 200 or response.status_code == 201:
                print(GREEN + f'=> Uploaded "{file_name}" to One Drive!' + RESET)
            else:
                print(RED + f'Error uploading "{file_name}" to One Drive: =>' + RESET)
                print(RED + f'URL: {upload_url}\nStatus: {response.status_code}\nResponse: {response.text}\n' + RESET)

    except Exception as e:
        print(RED + f"Error during One Drive upload preparation or execution: =>\n\n {str(e)}" + RESET)

def authenticate_google_drive():
    SCOPES = ["https://www.googleapis.com/auth/drive"]
    google_creds = None
    if os.path.exists(google_token_path):
        try:
            google_creds = Credentials.from_authorized_user_file(google_token_path, SCOPES)
            print(GREEN + "=> Google Drive token authenticated successfully!" + RESET)
        except Exception as e:
            print(RED + f"=> Error loading Google credentials from token file: {e}" + RESET)
            google_creds = None

    if not google_creds or not google_creds.valid:
        if google_creds and google_creds.expired and google_creds.refresh_token:
            print("Google token expired, attempting to refresh...")
            try:
                google_creds.refresh(Request())
                print(GREEN + "=> Google token refreshed successfully" + RESET)
            except Exception as e:
                print(RED + f"=> Error refreshing Google token: {e}" + RESET)
                google_creds = None

        if not google_creds:
            print("Performing new Google Drive authentication flow...")
            try:
                flow = InstalledAppFlow.from_client_secrets_file(google_credentials_path, SCOPES)
                google_creds = flow.run_local_server(port=0)
            except Exception as e:
                print(RED + f"Error during new Google authentication flow: {e} =>" + RESET)
                return None

        if google_creds:
            try:
                with open(google_token_path, "w") as token:
                    token.write(google_creds.to_json())
            except Exception as e:
                print(RED + f"Error saving Google token: {e} =>" + RESET)
    return google_creds

def upload_to_google_drive(google_creds):
    folders = ["PRogram"]
    try:
        service = build("drive", "v3", credentials=google_creds)
        for folder_name in folders:
            response = service.files().list(
                q=f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
                spaces="drive",
                fields="files(id, name)"
            ).execute()

            if not response["files"]:
                print(f'Folder "{folder_name}" not found in Google Drive, creating it...')
                file_metadata = { "name": folder_name, "mimeType": "application/vnd.google-apps.folder" }
                created_folder = service.files().create(body=file_metadata, fields="id").execute()
                folder_id = created_folder.get("id")
                print(GREEN + f"Folder '{folder_name}' created with ID: {folder_id} =>" + RESET)
            else:
                folder_id = response["files"][0]["id"]

            file_to_upload_name = "ThePRogram2025.pdf"
            response = service.files().list(
                q=f"name='{file_to_upload_name}' and parents in '{folder_id}' and trashed=false",
                spaces="drive", fields="files(id, name)"
            ).execute()
            for existing_file in response.get("files", []):
                existing_file_id = existing_file["id"]
                service.files().delete(fileId=existing_file_id).execute()
                print(GREEN + f'=> Successfully deleted existing file "{existing_file["name"]}"!' + RESET)

            file_metadata = { "name": file_to_upload_name, "parents": [folder_id] }
            if not os.path.exists(pdf_path):
                 print(RED + f"=> Local file '{pdf_path}' not found. Cannot upload to Google Drive!" + RESET)
                 continue

            media = MediaFileUpload(pdf_path, mimetype='application/pdf')
            upload_file = service.files().create(body=file_metadata, media_body=media, fields="id").execute()
            media.stream().close()
            print(GREEN + f'=> Uploaded "{file_to_upload_name}" to Google Drive folder "{folder_name}"' + RESET)

    except HttpError as e:
        print(RED + f"=> An HTTP error occurred with Google Drive API: {e}" + RESET)
    except Exception as e:
        print(RED + f"=> An unexpected error occurred during Google Drive operations: {str(e)}" + RESET)


print(DARK_CYAN + "\n[One Drive Authentication]" + RESET)
one_drive_access_token = authenticate_onedrive()
if not one_drive_access_token:
    print(RED + "=> Failed to authenticate with One Drive. Exiting!" + RESET)
    exit(1)

onedrive_target_folder = "Training"
file_to_download_and_edit = "ThePRogram2025.docx"

print(DARK_CYAN + f'\n[Downloading "{file_to_download_and_edit}" from One Drive]' + RESET)
download_success = download_file_from_onedrive(one_drive_access_token, onedrive_target_folder, file_to_download_and_edit, docx_path)

if not download_success:
    if os.path.exists(docx_path):
        print(GREEN + f'=> Using existing local file: "{docx_path}"' + RESET)
    else:
        print(RED + f'=> Local file "{docx_path}" also not found. Exiting!' + RESET)
        exit(1)
else:
    print(DARK_CYAN + f'\n[Deleting "{file_to_download_and_edit}" from One Drive post-download]' + RESET)
    delete_success = delete_file_from_onedrive(one_drive_access_token, onedrive_target_folder, file_to_download_and_edit)
    if not delete_success:
        print(RED + f"=> Failed to delete '{file_to_download_and_edit}' from One Drive. Proceeding, but this might cause issues if the file was locked!" + RESET)


print(DARK_CYAN + f'\n[Opening "{file_to_download_and_edit}" for editing]' + RESET)
if not os.path.exists(docx_path):
    print(RED + f'=> "{docx_path}" not found before attempting to open. Exiting!' + RESET)
    exit(1)
try:
    if platform.system() == "Windows":
        os.startfile(docx_path)
    elif platform.system() == "Darwin":
        subprocess.run(['open', docx_path], check=False)
    else:
        subprocess.run(['xdg-open', docx_path], check=False)
except Exception as e:
    print(RED + f"=> Could not automatically open the file: {e}" + RESET)
    print(f"=> Please open '{docx_path}' manually from your file explorer!")

input(RED + f'\nEdit "{file_to_download_and_edit}", save the changes, and CLOSE THE DOCUMENT EDITOR!\n'
      'Once done, press Enter in this console window to continue...' + RESET)

print(DARK_CYAN + '\n[Converting ".docx" to ".pdf"]' + RESET)
if not os.path.exists(docx_path):
    print(RED + f"=> Document '{docx_path}' not found after editing step. Cannot convert to PDF. Exiting!" + RESET)
    exit(1)
try:
    convert(docx_path, pdf_path)
    print(GREEN + f'=> Converted "{os.path.basename(docx_path)}" to "{os.path.basename(pdf_path)}" successfully!' + RESET)
except Exception as e:
    print(RED + f"=> Error during DOCX to PDF conversion: {e}" + RESET)
    print("Attempting to proceed with uploads, but PDF-related parts might fail or use stale data!")

print(DARK_CYAN + "\n[Uploading to One Drive]" + RESET)
upload_to_onedrive(one_drive_access_token)

if os.path.exists(pdf_path):
    print(DARK_CYAN + "\n[Google Drive Authentication]" + RESET)
    google_drive_creds = authenticate_google_drive()
    if google_drive_creds:
        print(DARK_CYAN + "\n[Google Drive Upload]" + RESET)
        upload_to_google_drive(google_drive_creds)
    else:
        print(RED + "=> Skipping Google Drive upload due to authentication failure" + RESET)
else:
    print(f"=> Skipping Google Drive upload as '{os.path.basename(pdf_path)}' was not found!")

print(DARK_CYAN + "\n[Cleaning up local files]" + RESET)
clean_local_folder(training_pdf_path)
clean_local_folder(pdf_path)