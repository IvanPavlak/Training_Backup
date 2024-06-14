# Training Backup

## Overview

This script automates converting a .docx file to .pdf, uploading both the resulting PDF and the original .docx to OneDrive, and distributing the PDF to multiple Google Drives. It also handles local file clean-up, all in a single PowerShell command.
____
# Usage

1. Setup One Drive and Google Drive (see Requirements).
2. Clone or download this repository to the local machine.
3. Copy the `TrainingBackupCredentials` folder into the repository.
4. Create a conda environment with the necessary dependencies (see Requirements).
5. Modify the paths in the `TrainingBackup.bat` file to call the right `activate.bat` and `miniconda3` (or Anaconda) on your machine.
6. Modify the paths of the python call in the `TrainingBackup.bat` to the `TrainingBackup.py` in cloned/downloaded repository (supports multiple machines).
7. Create a `configuration.json` with paths to the training and credentials folders.
8. Modify in `TrainingBackup.py`:
	- `docx_path` - modify to accommodate your naming.
	- `pdf_path` - modify to accommodate your naming.
	- `Convert .docx to .pdf` - modify to accommodate your naming.
	- `One Drive Upload` - modify to accommodate your needs.
		- Currently uploads the whole .docx file and the last page of the .pdf to the root folder in OneDrive.
	- `Google Drive Upload` - modify to accommodate your needs.
		- Currently uploads the .pdf to specified folders saved in the `folders` variable.
			- If the file already exists, it deletes it.
				- This prevents piling of duplicates in the drive.
9. Copy `TrainingBackupManualUpdate.ps1`'s contents into the `Microsoft.PowerShell_profile.ps1` and modify the `$backupDirectory` variable.
10. In PowerShell, type `Training-Backup` (PowerShell is case insensitive!).
___
# Requirements

## One Drive

- [Microsoft Azure](https://azure.microsoft.com/en-us)
	- [Azure App Signup step by step](https://github.com/pranabdas/Access-OneDrive-via-Microsoft-Graph-Python/blob/main/Azure_app_signup_step_by_step.md)
	- Client ID
	- Client Secret -> Created additionally in "Certificates & Secrets" section.
	- Both must be copied into a `onedrive_credentials.json`:
```
{
    "client_id": "your_client_id",
    "client_secret": "your_secret_value"
}
```

## Google Drive

- [Google Cloud](https://cloud.google.com/gcp?utm_source=google&utm_medium=cpc&utm_campaign=emea-emea-all-en-bkws-all-all-trial-e-gcp-1707574&utm_content=text-ad-none-any-DEV_c-CRE_683761846512-ADGP_Hybrid+%7C+BKWS+-+EXA+%7C+Txt+-+GCP+-+General+-+v2-KWID_43700078882258013-kwd-6458750523-userloc_1007612&utm_term=KW_google%20cloud-NET_g-PLAC_&&gad_source=1&gclid=CjwKCAiAivGuBhBEEiwAWiFmYSEVAU4nVtvqTjYCKbWC08C1ap_UukXjFhKNnvw9t3uknDf6DtumLBoCJTwQAvD_BwE&gclsrc=aw.ds)
	- Console -> APIs & Services -> OAuth consent screen -> Create a Project -> User Type: External -> Add yourself as Test User -> Credentials -> Create Credentials: OAuth client ID -> Desktop app -> Save the `json` file to `google_credentials.json`
 
## Training Backup Credentials Folder

- Create a `TrainingBackupCredentials` folder.
	- It should contain:
		- `onedrive_credentials.json`
		- `google_credentials.json`

## Anaconda/Conda Environment with Dependencies

- [Anaconda](https://docs.anaconda.com/free/anaconda/install/index.html) or [Miniconda](https://docs.anaconda.com/free/miniconda/index.html)
- Create an environment with:
	- `conda create --name TrainingBackup`
- Dependencies:
	- `pip install requests docx2pdf pypdf google-auth google-auth-oauthlib google-api-python-client`
___
# Functions

1. **`clean_local_folder(file_path)`**:
	- **Purpose**: Deletes a file from the local folder if it exists.
	- Arguments:
		- `file_path (str)`: Path of the file to be deleted.
2. **`authenticate_onedrive()`**:
	- **Purpose**: Authenticates with OneDrive using OAuth2. If tokens are expired or missing, it initiates the authentication process.
	- **Returns**:
		- str: Access token for OneDrive API.
1. **`exchange_code_for_tokens(code)`**:
	- **Purpose**: Exchanges the authorization code for access and refresh tokens.
	- Arguments:
		- code (str): Authorization code received from the OAuth2 flow.	
	- **Returns**:
		- tuple: Access token, expiry time, and refresh token.
1. **`refresh_access_token(refresh_token)`**:
	- **Purpose**: Refreshes the access token using the refresh token.
	- **Arguments**:
		- `refresh_token (str)`: Refresh token for obtaining a new access token.
	- **Returns**:
		- tuple: New access token, expiry time, and refresh token.
1. **`upload_to_onedrive(access_token)`**:
	- **Purpose**: Uploads the PDF and DOCX files to OneDrive.
	- **Arguments**:
		- 	`access_token (str)`: Access token for OneDrive API.
1. **`authenticate_google_drive()`**:
	- **Purpose**: Authenticates with Google Drive using OAuth2.
	- **Returns**: 
		- google.oauth2.credentials.Credentials: Google API credentials.
1. **`upload_to_google_drive(credentials)`**:
	- **Purpose**: Uploads a PDF file to specified folders in the Google Drive. If there already is a file with the same name, it is deleted before the upload.
		- this deletion is necessary because two files with the same name can be uploaded to Google Drive, which results in unnecessary clutter.
	- **Parameters**:
		- credentials: Google API credentials.
___
# Error Handling

- **Authentication Errors**:
	- Prompts reauthentication if credentials are invalid or expired.

- **Upload Errors**:
	- Retries upload on failure and notifies the user if unsuccessful.

- **File Path Errors**:
	- Informs the user of incorrect file paths and provides guidance on verification.
___
# Security Considerations:

- **Credential Management**:
	- Credentials are stored securely in JSON files and access is restricted.

- **Token Handling**:
	- Secure handling of tokens with proper rotation and expiration strategies.

- **HTTPS Usage**:
	- All communications are encrypted via HTTPS

- **Input Sanitization**:
	- Validates and sanitizes user inputs to prevent security risks.

- **Error Logging**:
	- Logs errors securely for regular reviews and potential issue identification.
___
# Helpful Sources

- These were studied, used and modified to accomplish the desired effect:
#### Google Drive Backup:
- [Automated Google Drive Backups in Python](https://www.youtube.com/watch?v=fkWM7A-MxR0)
#### One Drive Backup:
- [Access-OneDrive-via-Microsoft-Graph-Python](https://github.com/pranabdas/Access-OneDrive-via-Microsoft-Graph-Python)
___
