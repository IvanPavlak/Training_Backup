import os
import os.path
import time
import json
import urllib
import requests
from getpass import getpass
from docx2pdf import convert
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# Convert .docx to .pdf
docx_path = r"E:\Accountability\Training\2024\ThePRogram2024.docx"
pdf_path = r"E:\Accountability\Training\2024\ThePRogram2024.pdf"
convert(docx_path, pdf_path)

# Delete local PDF file
def clean_local_folder(filename, path):
    file_path = os.path.join(path, filename)
    if os.path.exists(file_path):
        os.remove(file_path)
        print("---------------------------------------------------------------------------")
        print(f"'{filename}' deleted successfully from local folder.")
        print("---------------------------------------------------------------------------")
    else:
        print("---------------------------------------------------------------------------")
        print(f"'{filename}' does not exist in the specified path.")
        print("---------------------------------------------------------------------------")

# OneDrive Upload
# Authentication URL and credentials
URL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
client_id = "d8e6ca27-806b-4368-aa7f-6df4db14976b"
permissions = ["files.readwrite"]
response_type = "token"
redirect_uri = "http://localhost:8080/"
scope = ""
for items in range(len(permissions)):
    scope = scope + permissions[items]
    if items < len(permissions)-1:
        scope = scope + "+"

# Instructions for obtaining authentication code
print("---------------------------------------------------------------------------")
print("Click over this link:\n")
print(URL + "?client_id=" + client_id + "&scope=" + scope + "&response_type=" + response_type+\
     "&redirect_uri=" + urllib.parse.quote(redirect_uri))
print("---------------------------------------------------------------------------")
print("Sign in to your account, copy the whole redirected URL!")
code = input("Paste the URL here:\n")
print("---------------------------------------------------------------------------")
token = code[(code.find('access_token') + len('access_token') + 1) : (code.find('&token_type'))]

# Authorization header for API requests
URL = 'https://graph.microsoft.com/v1.0/'
HEADERS = {'Authorization': 'Bearer ' + token}

# Check response for authentication success
response = requests.get(URL + 'me/drive/', headers = HEADERS)
if (response.status_code == 200):
    response = json.loads(response.text)
    print('Connected to the OneDrive of', response['owner']['user']['displayName']+' (',response['driveType']+' ).', \
         '\nConnection valid for one hour. Reauthenticate if required.')
elif (response.status_code == 401):
    response = json.loads(response.text)
    print('API Error! : ', response['error']['code'],\
         '\nSee response for more details.')
else:
    response = json.loads(response.text)
    print('Unknown error! See response for more details.')

# Upload file to OneDrive
url = 'me/drive/root:/ThePRogram2024.pdf:/content'
url = URL + url
content = open(pdf_path, 'rb')
response = json.loads(requests.put(url, headers=HEADERS, data = content).text)
content.close()

# Google Drive Upload
# Authorization and Authentication
SCOPES = ["https://www.googleapis.com/auth/drive"]

# Check if token file exists and if credentials are valid, otherwise authenticate
credentials = None
if os.path.exists("token.json"):
    credentials = Credentials.from_authorized_user_file("token.json", SCOPES)

if not credentials or not credentials.valid:
    if credentials and credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(r"E:\Accountability\Training\2024\credentials.json", SCOPES)
        credentials = flow.run_local_server(port=0)

    # Save new token to file
    with open("token.json", "w") as token:
        token.write(credentials.to_json())

# Upload file to Google Drive
try:
    service = build("drive", "v3", credentials=credentials)

    # Check if PRogram1 folder exists, create if it doesn't
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

    # Check if ThePRogram2024.pdf already exists in PRogram1 folder, delete if it does
    response = service.files().list(
        q="name='ThePRogram2024.pdf' and '" + folder_id + "' in parents",
        spaces="drive"
    ).execute()

    if response["files"]:
        file_id = response["files"][0]["id"]
        service.files().delete(fileId=file_id).execute()
        print("---------------------------------------------------------------------------")
        print("Deleted existing file from Drive: ThePRogram2024.pdf")

    # Upload the new version of ThePRogram2024.pdf
    file_name = "ThePRogram2024.pdf"
    file_metadata = {
        "name": file_name,
        "parents": [folder_id]
    }
    
    media = MediaFileUpload(r"E:\Accountability\Training\2024\ThePRogram2024.pdf")
    upload_file = service.files().create(body=file_metadata,
                                         media_body=media,
                                         fields="id").execute()
    
    media.stream().close()
    print("---------------------------------------------------------------------------")
    print("Uploaded file: " + file_name)

except HttpError as e:
    print("Error: " + str(e))

clean_local_folder(r"ThePRogram2024.pdf", r"E:\Accountability\Training\2024")

time.sleep(5)