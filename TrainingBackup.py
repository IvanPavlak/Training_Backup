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
from urllib.parse import urlparse, parse_qs, quote
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# --- Constants ---
RESET = "\033[0m"
DARK_CYAN = "\033[36m"
GREEN = "\033[92m"
RED = "\033[91m"

# --- Configuration Loading ---
hostname = socket.gethostname()
try:
    with open("configuration.json") as f:
        config = json.load(f)
except FileNotFoundError:
    print(RED + "=> Configuration file 'configuration.json' not found!" + RESET)
    exit("Exiting due to configuration error.")
except json.JSONDecodeError:
    print(RED + "=> Error decoding 'configuration.json'. Please ensure it's valid JSON." + RESET)
    exit("Exiting due to configuration error.")


if hostname in config:
    paths = config[hostname]
    training_folder = paths.get("training_folder")
    credentials_folder = paths.get("credentials_folder")
    if not training_folder or not credentials_folder:
        print(RED + '=> "training_folder" or "credentials_folder" missing in configuration for this host!' + RESET)
        exit("Exiting due to configuration error.")
else:
    print(RED + f"=> Hostname '{hostname}' not found in configuration file!" + RESET)
    exit("Exiting due to configuration error.")

# --- Global Variable Definitions ---
# File and Folder Names
ONEDRIVE_TARGET_FOLDER = "Training"
FILE_TO_DOWNLOAD_AND_EDIT = "ThePRogram2025.docx"
PDF_OUTPUT_FILENAME = "ThePRogram2025.pdf"
ONEDRIVE_TRAINING_PDF_FILENAME = "Training.pdf" # Last page of PDF_OUTPUT_FILENAME
GOOGLE_DRIVE_UPLOAD_FOLDER = "PRogram"
GOOGLE_DRIVE_UPLOAD_FILENAME = "ThePRogram2025.pdf"

# Local Paths
docx_path = os.path.join(training_folder, FILE_TO_DOWNLOAD_AND_EDIT)
pdf_path = os.path.join(training_folder, PDF_OUTPUT_FILENAME)
training_pdf_path = os.path.join(training_folder, ONEDRIVE_TRAINING_PDF_FILENAME)

# Credentials Paths
google_token_path = os.path.join(credentials_folder, "google_token.json")
google_credentials_path = os.path.join(credentials_folder, "google_credentials.json")

onedrive_token_path = os.path.join(credentials_folder, "onedrive_token.json")
one_drive_credentials_path = os.path.join(credentials_folder, "onedrive_credentials.json")

try:
    with open(one_drive_credentials_path, "r") as f:
        onedrive_creds_json = json.load(f)
except FileNotFoundError:
    print(RED + f"=> OneDrive credentials file '{one_drive_credentials_path}' not found!" + RESET)
    exit("Exiting due to configuration error.")
except json.JSONDecodeError:
    print(RED + f"=> Error decoding OneDrive credentials file '{one_drive_credentials_path}'." + RESET)
    exit("Exiting due to configuration error.")

onedrive_client_id = onedrive_creds_json.get("client_id")
onedrive_client_secret = onedrive_creds_json.get("client_secret")
if not onedrive_client_id or not onedrive_client_secret:
    print(RED + "=> 'client_id' or 'client_secret' missing in OneDrive credentials file." + RESET)
    exit("Exiting due to configuration error.")

# OneDrive OAuth Constants
REDIRECT_URI = "http://localhost:8080/"
TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
AUTH_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
ONEDRIVE_SCOPES = "files.readwrite offline_access"

# Google Drive OAuth Constants
GOOGLE_DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]


# --- Function Definitions ---

def exchange_code_for_tokens(code):
    """
    Exchanges an OAuth 2.0 authorization code for an access token, refresh token,
    and expiration time from Microsoft's token endpoint.

    This function is part of the OneDrive authentication flow when a new authorization
    is required.

    Parameters:
    code (str): The authorization code obtained from the user after they
                authorize the application via the authorization URL.

    Returns:
    tuple: A tuple containing (access_token, expires_at, refresh_token).
           - access_token (str): The access token for Microsoft Graph API.
           - expires_at (float): The Unix timestamp when the access token will expire.
           - refresh_token (str): The refresh token to obtain new access tokens.
           Returns (None, None, None) if the exchange fails.
    """
    token_params = {
        "client_id": onedrive_client_id,
        "client_secret": onedrive_client_secret,
        "code": code,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code"
    }
    try:
        response = requests.post(TOKEN_URL, data=token_params)
        response.raise_for_status()
        token_data = response.json()
        access_token = token_data.get("access_token")
        expires_in = token_data.get("expires_in", 3600) # Default to 1 hour
        expires_at = time.time() + expires_in
        refresh_token = token_data.get("refresh_token")
        if not all([access_token, refresh_token]):
            print(RED + f"=> Failed to retrieve all necessary tokens from response: {token_data}" + RESET)
            return None, None, None
        return access_token, expires_at, refresh_token
    except requests.exceptions.RequestException as e:
        print(RED + f"=> Error exchanging code for tokens: {e}" + RESET)
        if hasattr(e, 'response') and e.response is not None:
            print(RED + f"Response content: {e.response.text}" + RESET)
        return None, None, None
    except json.JSONDecodeError:
        print(RED + f"=> Error decoding JSON response from token endpoint." + RESET)
        return None, None, None

def refresh_access_token(refresh_token):
    """
    Refreshes an expired OneDrive access token using a provided refresh token.

    This function is called when an existing access token is found to be expired
    during OneDrive authentication.

    Parameters:
    refresh_token (str): The refresh token obtained during a previous authorization.

    Returns:
    tuple: A tuple containing (access_token, expires_at, new_refresh_token).
           - access_token (str): The new access token.
           - expires_at (float): The Unix timestamp when the new access token expires.
           - new_refresh_token (str): The new refresh token (often the same as input,
                                      but sometimes a new one is issued).
           Returns (None, None, None) if refreshing fails.
    """
    token_params = {
        "client_id": onedrive_client_id,
        "client_secret": onedrive_client_secret,
        "refresh_token": refresh_token,
        "grant_type": "refresh_token",
        "scope": ONEDRIVE_SCOPES
    }
    try:
        response = requests.post(TOKEN_URL, data=token_params)
        response.raise_for_status()
        token_data = response.json()
        access_token = token_data.get("access_token")
        expires_in = token_data.get("expires_in", 3600)
        expires_at = time.time() + expires_in
        new_refresh_token = token_data.get("refresh_token", refresh_token)
        if not all([access_token, new_refresh_token]):
            print(RED + f"=> Failed to retrieve all necessary tokens during refresh: {token_data}" + RESET)
            return None, None, None
        return access_token, expires_at, new_refresh_token
    except requests.exceptions.RequestException as e:
        print(RED + f"=> Error refreshing One Drive token: {e}" + RESET)
        if hasattr(e, 'response') and e.response is not None:
            print(RED + f"Response content: {e.response.text}" + RESET)
        return None, None, None
    except json.JSONDecodeError:
        print(RED + f"=> Error decoding JSON response during token refresh." + RESET)
        return None, None, None

def authenticate_onedrive():
    """
    Authenticates the user with OneDrive using OAuth 2.0.

    It first checks for a locally stored, valid access token. If found and valid,
    it's returned. If expired, it attempts to refresh it. If no token exists or
    refreshing fails, it initiates the full OAuth 2.0 authorization code grant flow,
    prompting the user to authorize the application in their browser.
    The obtained tokens (access and refresh) are saved locally for future use.

    Parameters:
    None

    Returns:
    str: The access token if authentication is successful, otherwise None.
         This token is used for subsequent API calls to OneDrive.
    """
    if os.path.exists(onedrive_token_path):
        try:
            with open(onedrive_token_path, "r") as token_file:
                token_data = json.load(token_file)
            access_token = token_data.get("access_token")
            expires_at = token_data.get("expires_at")
            stored_refresh_token = token_data.get("refresh_token")

            if access_token and expires_at and expires_at > time.time():
                print(GREEN + "=> OneDrive token authenticated successfully from stored file!" + RESET)
                return access_token

            if stored_refresh_token:
                print("OneDrive token expired or invalid, attempting to refresh...")
                refreshed_access_token, new_expires_at, new_refresh_token = refresh_access_token(stored_refresh_token)
                if refreshed_access_token:
                    print(GREEN + "=> OneDrive token refreshed successfully!" + RESET)
                    with open(onedrive_token_path, "w") as new_token_file:
                        new_token_data = {
                            "access_token": refreshed_access_token,
                            "expires_at": new_expires_at,
                            "refresh_token": new_refresh_token
                        }
                        json.dump(new_token_data, new_token_file)
                    return refreshed_access_token
                else:
                    print(RED + "=> Failed to refresh OneDrive token. Proceeding to full authentication!" + RESET)
        except (FileNotFoundError, json.JSONDecodeError, KeyError) as e:
            print(RED + f"=> Error reading or parsing OneDrive token file ({onedrive_token_path}): {e}. Proceeding to full authentication." + RESET)


    auth_params = {
        "client_id": onedrive_client_id,
        "redirect_uri": REDIRECT_URI,
        "scope": ONEDRIVE_SCOPES,
        "response_type": "code"
    }

    full_auth_url = AUTH_URL + "?" + urllib.parse.urlencode(auth_params)
    print("\nClick on this link to authenticate with OneDrive:\n")
    print(DARK_CYAN + full_auth_url + RESET)
    url_input = input("\nCopy the redirected URL from your browser and paste it here:\n\n")

    try:
        code = parse_qs(urlparse(url_input).query).get('code', [''])[0]
        if not code:
            print(RED + "=> Could not extract authorization code from the redirected URL." + RESET)
            return None
    except Exception as e:
        print(RED + f"=> Error parsing the redirected URL: {e}" + RESET)
        return None

    access_token, expires_at, refresh_token = exchange_code_for_tokens(code)

    if access_token:
        print(GREEN + "=> OneDrive authenticated successfully via new authorization!" + RESET)
        with open(onedrive_token_path, "w") as token_file:
            token_data = {
                "access_token": access_token,
                "expires_at": expires_at,
                "refresh_token": refresh_token
            }
            json.dump(token_data, token_file)
        return access_token
    else:
        print(RED + "=> Failed to obtain OneDrive tokens after authorization." + RESET)
        return None

def download_file_from_onedrive(access_token, onedrive_folder, onedrive_filename, local_target_path):
    """
    Downloads a specific file from a specified folder in OneDrive to a local path.

    Constructs the Microsoft Graph API URL for the file content and makes a GET
    request. If successful, streams the file content to the specified local path.
    Creates parent directories for the local path if they don't exist.

    Parameters:
    access_token (str): The valid OneDrive access token.
    onedrive_folder (str): The name of the folder in OneDrive containing the file.
    onedrive_filename (str): The name of the file to download from OneDrive.
    local_target_path (str): The full local file path where the downloaded file
                             should be saved.

    Returns:
    bool: True if the download was successful, False otherwise.
    """
    local_dir = os.path.dirname(local_target_path)
    if local_dir and not os.path.exists(local_dir):
        try:
            os.makedirs(local_dir, exist_ok=True)
            print(f"Created directory: {local_dir}")
        except OSError as e:
            print(RED + f"Error creating directory {local_dir}: {e}" + RESET)
            return False


    item_path = f"/{onedrive_folder}/{onedrive_filename}"
    encoded_item_path = quote(item_path)

    download_url = f"https://graph.microsoft.com/v1.0/me/drive/root:{encoded_item_path}:/content"
    headers = {"Authorization": "Bearer " + access_token}

    try:
        response = requests.get(download_url, headers=headers, stream=True)
        if response.status_code == 200:
            with open(local_target_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            print(GREEN + f'=> Downloaded "{onedrive_filename}" to "{local_target_path}" successfully!' + RESET)
            return True
        else:
            error_message = f'=> Error downloading "{onedrive_filename}" from OneDrive: {response.status_code}'
            try:
                error_details = response.json()
                if error_details.get("error", {}).get("code") == "itemNotFound":
                    error_message += f' - File not found in OneDrive folder "{onedrive_folder}".'
                else:
                    error_message += f' - Response: {response.text}'
            except json.JSONDecodeError:
                 error_message += f' - Response: {response.text}'
            print(RED + error_message + RESET)
            return False
    except requests.exceptions.RequestException as e:
        print(RED + f"=> Network or request error during OneDrive download: {e}" + RESET)
        return False

def delete_file_from_onedrive(access_token, onedrive_folder, onedrive_filename):
    """
    Deletes a specific file from a specified folder in OneDrive.

    Constructs the Microsoft Graph API URL for the item and makes a DELETE request.
    A 404 (Not Found) status is treated as a successful deletion for the
    purpose of this script's workflow (i.e., the file is no longer there).

    Parameters:
    access_token (str): The valid OneDrive access token.
    onedrive_folder (str): The name of the folder in OneDrive from which to delete the file.
    onedrive_filename (str): The name of the file to delete from OneDrive.

    Returns:
    bool: True if the file was successfully deleted or if it was not found (which
          also means it's "gone"). False if an an unexpected error occurred during deletion.
    """
    item_path = f"/{onedrive_folder}/{onedrive_filename}"
    encoded_item_path = quote(item_path)
    delete_url = f"https://graph.microsoft.com/v1.0/me/drive/root:{encoded_item_path}"
    headers = {"Authorization": "Bearer " + access_token}

    try:
        response = requests.delete(delete_url, headers=headers)
        if response.status_code == 204: # No Content - successful deletion
            print(GREEN + f'=> Successfully deleted "{onedrive_filename}" from OneDrive folder "{onedrive_folder}"!' + RESET)
            return True
        elif response.status_code == 404: # Not Found
            print(f'=> File "{onedrive_filename}" not found in OneDrive folder "{onedrive_folder}" for deletion (considered success for workflow).' + RESET)
            return True # File is not there, so effectively "deleted" for our purposes
        else:
            print(RED + f'=> Error deleting "{onedrive_filename}" from OneDrive: {response.status_code}' + RESET)
            print(RED + f'URL: {delete_url}' + RESET)
            print(RED + f'Response: {response.text}\n' + RESET)
            return False
    except requests.exceptions.RequestException as e:
        print(RED + f"=> Network or request error during OneDrive deletion: {e}" + RESET)
        return False

def upload_to_onedrive(access_token):
    """
    Uploads specified files to the 'Training' folder in OneDrive.

    Specifically, it attempts to upload:
    1. "ThePRogram2025.docx" (from `docx_path`).
    2. "Training.pdf" (which is created by extracting the last page of
       "ThePRogram2025.pdf" located at `pdf_path`).

    If "ThePRogram2025.pdf" doesn't exist or is empty, it will skip creating
    "Training.pdf" and might only upload the .docx file if available.
    It handles file existence checks before attempting uploads.

    Parameters:
    access_token (str): The valid OneDrive access token.

    Returns:
    None. Prints status messages to the console.
    """
    files_to_upload_map = {} # {onedrive_filename: local_path}

    # Prepare "Training.pdf" (last page of the main PDF)
    if os.path.exists(pdf_path):
        try:
            reader = PdfReader(pdf_path)
            if reader.pages:
                last_page = reader.pages[-1]
                writer = PdfWriter()
                writer.add_page(last_page)
                with open(training_pdf_path, "wb") as output_pdf:
                    writer.write(output_pdf)
                files_to_upload_map[ONEDRIVE_TRAINING_PDF_FILENAME] = training_pdf_path
            else:
                print(RED + f'=> "{PDF_OUTPUT_FILENAME}" contains no pages. Cannot create "{ONEDRIVE_TRAINING_PDF_FILENAME}".' + RESET)
        except Exception as e:
            print(RED + f"=> Error creating '{ONEDRIVE_TRAINING_PDF_FILENAME}' from '{PDF_OUTPUT_FILENAME}': {e}" + RESET)
    else:
        print(RED + f'=> "{PDF_OUTPUT_FILENAME}" not found. Cannot create "{ONEDRIVE_TRAINING_PDF_FILENAME}" for OneDrive upload.' + RESET)

    # Prepare "ThePRogram2025.docx"
    if os.path.exists(docx_path):
        files_to_upload_map[FILE_TO_DOWNLOAD_AND_EDIT] = docx_path
    else:
        print(RED + f'=> Main document "{FILE_TO_DOWNLOAD_AND_EDIT}" not found at "{docx_path}". Cannot upload to OneDrive.' + RESET)

    if not files_to_upload_map:
        print(RED + "=> No files are available or prepared for OneDrive upload. Skipping." + RESET)
        return

    for onedrive_filename, local_file_path_to_upload in files_to_upload_map.items():
        if not os.path.exists(local_file_path_to_upload):
            print(f'File "{local_file_path_to_upload}" for OneDrive upload as "{onedrive_filename}" not found. Skipping this file!')
            continue

        item_path = f"/{ONEDRIVE_TARGET_FOLDER}/{onedrive_filename}"
        encoded_item_path = quote(item_path)
        upload_url = f"https://graph.microsoft.com/v1.0/me/drive/root:{encoded_item_path}:/content"
        headers = {"Authorization": "Bearer " + access_token, "Content-Type": "application/octet-stream"}

        try:
            with open(local_file_path_to_upload, "rb") as file_content:
                response = requests.put(upload_url, headers=headers, data=file_content)

            if response.status_code == 200 or response.status_code == 201: # 200 OK (updated), 201 Created
                print(GREEN + f'=> Uploaded "{onedrive_filename}" to OneDrive folder "{ONEDRIVE_TARGET_FOLDER}" successfully!' + RESET)
            else:
                print(RED + f'Error uploading "{onedrive_filename}" to OneDrive: {response.status_code}' + RESET)
                print(RED + f'URL: {upload_url}\nResponse: {response.text}\n' + RESET)
        except requests.exceptions.RequestException as e:
            print(RED + f"Network or request error during OneDrive upload of '{onedrive_filename}': {e}" + RESET)
        except IOError as e:
            print(RED + f"IOError reading file '{local_file_path_to_upload}' for upload: {e}" + RESET)
        except Exception as e:
            print(RED + f"An unexpected error occurred during OneDrive upload of '{onedrive_filename}': {e}" + RESET)

def authenticate_google_drive():
    """
    Authenticates the user with Google Drive using OAuth 2.0.

    Similar to OneDrive authentication, it checks for a local token, refreshes if
    expired, or initiates a new authorization flow via the browser.
    The obtained credentials are saved locally.

    Parameters:
    None

    Returns:
    google.oauth2.credentials.Credentials: A Credentials object if authentication
                                           is successful, otherwise None. This object
                                           is used to build the Google Drive API service.
    """
    google_creds = None
    if os.path.exists(google_token_path):
        try:
            google_creds = Credentials.from_authorized_user_file(google_token_path, GOOGLE_DRIVE_SCOPES)
        except Exception as e:
            print(RED + f"=> Error loading Google credentials from token file '{google_token_path}': {e}" + RESET)
            google_creds = None

    if not google_creds or not google_creds.valid:
        if google_creds and google_creds.expired and google_creds.refresh_token:
            print("Google Drive token expired, attempting to refresh...")
            try:
                google_creds.refresh(Request())
                print(GREEN + "=> Google Drive token refreshed successfully!" + RESET)
            except Exception as e:
                print(RED + f"=> Error refreshing Google Drive token: {e}. Proceeding to full authentication." + RESET)
                google_creds = None
        else:
            print("No valid Google Drive token found or refresh failed. Performing new Google Drive authentication flow...")

        if not google_creds: # This block runs if creds are None (initial, load fail, or refresh fail)
            try:
                if not os.path.exists(google_credentials_path):
                    print(RED + f"Google API credentials file ('{google_credentials_path}') not found. Cannot proceed with Google auth." + RESET)
                    return None
                flow = InstalledAppFlow.from_client_secrets_file(google_credentials_path, GOOGLE_DRIVE_SCOPES)
                google_creds = flow.run_local_server(port=0)
                print(GREEN + "=> Google Drive authenticated successfully via new authorization!" + RESET)
            except Exception as e:
                print(RED + f"Error during new Google authentication flow: {e}" + RESET)
                return None

        # Save the credentials for the next run
        if google_creds and google_creds.valid:
            try:
                with open(google_token_path, "w") as token:
                    token.write(google_creds.to_json())
            except Exception as e:
                print(RED + f"Error saving Google token to '{google_token_path}': {e}" + RESET)

    if google_creds and google_creds.valid:
        print(GREEN + "=> Google Drive token authenticated successfully!" + RESET)
        return google_creds
    else:
        print(RED + "=> Google Drive authentication ultimately failed." + RESET)
        return None

def upload_to_google_drive(google_creds):
    """
    Uploads the "ThePRogram2025.pdf" (from `pdf_path`) to a specified folder
    (GOOGLE_DRIVE_UPLOAD_FOLDER) in Google Drive.

    It will:
    1. Check if the target folder exists in Google Drive. If not, it creates it.
    2. Check if a file with the target name (GOOGLE_DRIVE_UPLOAD_FILENAME) already
       exists in that folder. If so, it deletes the existing file.
    3. Upload the local PDF file to the target folder.

    Parameters:
    google_creds (google.oauth2.credentials.Credentials): The authenticated Google
                                                          Drive credentials object.

    Returns:
    None. Prints status messages to the console.
    """
    if not os.path.exists(pdf_path):
         print(RED + f"=> Local file '{pdf_path}' (for Google Drive upload as '{GOOGLE_DRIVE_UPLOAD_FILENAME}') not found. Skipping upload." + RESET)
         return

    try:
        service = build("drive", "v3", credentials=google_creds)

        # 1. Find or create the target folder
        folder_id = None
        response = service.files().list(
            q=f"name='{GOOGLE_DRIVE_UPLOAD_FOLDER}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
            spaces="drive",
            fields="files(id, name)"
        ).execute()

        if not response["files"]:
            print(f'Folder "{GOOGLE_DRIVE_UPLOAD_FOLDER}" not found in Google Drive, creating it...')
            folder_metadata = { "name": GOOGLE_DRIVE_UPLOAD_FOLDER, "mimeType": "application/vnd.google-apps.folder" }
            created_folder = service.files().create(body=folder_metadata, fields="id").execute()
            folder_id = created_folder.get("id")
            print(GREEN + f"Folder '{GOOGLE_DRIVE_UPLOAD_FOLDER}' created with ID: {folder_id}" + RESET)
        else:
            folder_id = response["files"][0]["id"]

        if not folder_id:
            print(RED + f"Could not obtain folder ID for '{GOOGLE_DRIVE_UPLOAD_FOLDER}'. Cannot upload." + RESET)
            return

        # 2. Delete existing file if it exists in the folder
        response = service.files().list(
            q=f"name='{GOOGLE_DRIVE_UPLOAD_FILENAME}' and parents in '{folder_id}' and trashed=false",
            spaces="drive", fields="files(id, name)"
        ).execute()
        for existing_file in response.get("files", []):
            existing_file_id = existing_file["id"]
            service.files().delete(fileId=existing_file_id).execute()
            print(GREEN + f'=> Successfully deleted existing file "{existing_file["name"]}"!' + RESET)

        # 3. Upload the new file
        file_metadata = { "name": GOOGLE_DRIVE_UPLOAD_FILENAME, "parents": [folder_id] }
        media = MediaFileUpload(pdf_path, mimetype='application/pdf', resumable=True)
        upload_file = service.files().create(body=file_metadata, media_body=media, fields="id").execute()
        print(GREEN + f'=> Uploaded "{GOOGLE_DRIVE_UPLOAD_FILENAME}" to Google Drive folder "{GOOGLE_DRIVE_UPLOAD_FOLDER}" successfully!' + RESET)

    except HttpError as e:
        error_content = e.content.decode() if e.content else "No additional error content."
        print(RED + f"=> An HTTP error occurred with Google Drive API: {e._get_reason()}\nDetails: {error_content}" + RESET)
    except Exception as e:
        print(RED + f"=> An unexpected error occurred during Google Drive operations: {str(e)}" + RESET)

def clean_local_folder(file_path_to_delete):
    """
    Deletes a specified file from the local filesystem.

    Checks if the file exists before attempting deletion. Prints a success message
    if deleted, or a message if the file does not exist.

    Parameters:
    file_path_to_delete (str): The absolute or relative path to the file
                               that should be deleted.

    Returns:
    None.
    """
    file_name_only = os.path.basename(file_path_to_delete)
    if os.path.exists(file_path_to_delete):
        try:
            os.remove(file_path_to_delete)
            print(GREEN + f'=> Local file "{file_name_only}" deleted successfully from "{os.path.dirname(file_path_to_delete)}"!' + RESET)
        except OSError as e:
            print(RED + f'=> Error deleting local file "{file_name_only}": {e}' + RESET)
    else:
        print(f'=> Local file "{file_name_only}" does not exist at "{file_path_to_delete}", no need to delete.')


# --- Main Script Execution ---
if __name__ == "__main__":

    # 1. OneDrive Authentication
    print(DARK_CYAN + "\n[OneDrive Authentication]" + RESET)
    one_drive_access_token = authenticate_onedrive()
    if not one_drive_access_token:
        print(RED + "=> Critical: Failed to authenticate with OneDrive. Exiting script." + RESET)
        exit(1)

    # 2. Download file from OneDrive (or use local if download fails but local exists)
    print(DARK_CYAN + f'\n[Download/Prepare "{FILE_TO_DOWNLOAD_AND_EDIT}"]' + RESET)
    download_success = download_file_from_onedrive(
        one_drive_access_token,
        ONEDRIVE_TARGET_FOLDER,
        FILE_TO_DOWNLOAD_AND_EDIT,
        docx_path
    )

    if not download_success:
        print(f"=> Download of '{FILE_TO_DOWNLOAD_AND_EDIT}' from OneDrive failed.")
        if os.path.exists(docx_path):
            print(GREEN + f"=> Using existing local file: '{docx_path}'." + RESET)
        else:
            print(RED + f"=> Critical: Local file '{docx_path}' also not found. Cannot proceed. Exiting script." + RESET)
            exit(1)
    else:
        # 2a. Delete from OneDrive only if download was successful
        print(DARK_CYAN + f'\n[Delete "{FILE_TO_DOWNLOAD_AND_EDIT}" from OneDrive post-download]' + RESET)
        delete_success = delete_file_from_onedrive(
            one_drive_access_token,
            ONEDRIVE_TARGET_FOLDER,
            FILE_TO_DOWNLOAD_AND_EDIT
        )
        if not delete_success:
            print(RED + f"=> Warning: Failed to delete '{FILE_TO_DOWNLOAD_AND_EDIT}' from OneDrive. "
                  "This might cause issues if the file was intended to be exclusively processed." + RESET)

    # 3. Open file for editing (Manual Step)
    print(DARK_CYAN + f'\n[Open "{FILE_TO_DOWNLOAD_AND_EDIT}" for Editing]' + RESET)
    if not os.path.exists(docx_path): # Should be true if we reached here
        print(RED + f'=> Critical: Document "{docx_path}" not found before attempting to open. Exiting script.' + RESET)
        exit(1)
    try:
        os.startfile(docx_path)
    except Exception as e:
        print(RED + f"=> Could not automatically open the file '{docx_path}': {e}" + RESET)
        print(RED + f"=> Please open the file manually from your file explorer: {os.path.abspath(docx_path)}" + RESET)

    input(GREEN + f'=> ACTION REQUIRED\nEdit "{FILE_TO_DOWNLOAD_AND_EDIT}", save the changes, '
          'and CLOSE THE DOCUMENT EDITOR!\nOnce done, press Enter in this console window to continue...' + RESET)

    # 4. Convert DOCX to PDF
    print(DARK_CYAN + '\n[Convert ".docx" to ".pdf"]' + RESET)
    if not os.path.exists(docx_path):
        print(RED + f"=> Critical: Document '{docx_path}' not found after editing step. Cannot convert to PDF. Exiting script." + RESET)
        exit(1)
    try:
        convert(docx_path, pdf_path)
        if os.path.exists(pdf_path):
            print(GREEN + f'=> Converted "{os.path.basename(docx_path)}" to "{os.path.basename(pdf_path)}" successfully!' + RESET)
        else:
            print(RED + f"=> Conversion reported success, but PDF file '{pdf_path}' was not found. Check conversion tool." + RESET)
    except Exception as e:
        print(RED + f"=> Error during DOCX to PDF conversion: {e}" + RESET)
        print(RED + "Attempting to proceed with uploads, but PDF-related parts might fail or use stale/missing data." + RESET)

    # 5. Upload to OneDrive
    print(DARK_CYAN + "\n[Upload to OneDrive]" + RESET)
    upload_to_onedrive(one_drive_access_token)

    # 6. Google Drive Operations (Auth and Upload)
    if os.path.exists(pdf_path):
        print(DARK_CYAN + "\n[Google Drive Authentication]" + RESET)
        google_drive_creds = authenticate_google_drive()
        if google_drive_creds:
            print(DARK_CYAN + "\n[Google Drive Upload]" + RESET)
            upload_to_google_drive(google_drive_creds)
        else:
            print(RED + "=> Skipping Google Drive upload due to authentication failure." + RESET)
    else:
        print(RED + f"=> Skipping Google Drive operations as '{os.path.basename(pdf_path)}' was not found or failed to convert." + RESET)

    # 7. Clean up local generated files
    print(DARK_CYAN + "\n[Clean up local temporary files]" + RESET)
    clean_local_folder(training_pdf_path) # The one-page PDF for OneDrive
    clean_local_folder(pdf_path)          # The full converted PDF