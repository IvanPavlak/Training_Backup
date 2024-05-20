# Training Backup

Overview
This script automates the process of converting a .docx file to .pdf, uploading the resulting PDF to OneDrive and Google Drive, and then cleaning up local files. It is designed to be used with Task Scheduler for automated backups.

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
# Usage

1. Clone or download this repository to the local machine.
2. Copy the `TrainingBackupCredentials` folder into the repository.
3. Create a conda environment with the necessary dependencies.
4. Modify the paths in the `TrainingBackup.bat` file to call the right `activate.bat` and `miniconda3` (or Anaconda) on your machine.
5. Modify the path of the python call in the `TrainingBackup.bat` to the `TrainingBackup.py` in cloned/downloaded repository.
6. Modify in `TrainingBackup.py`:
	- `training_folder` - provide a full path to the folder.
	- `credentials_folder` - provide a full path to the folder.
	- `docx_path` - modify to accommodate your naming.
	- `pdf_path` - modify to accommodate your naming.
	- `Convert .docx to .pdf` - modify to accommodate your naming.
	- `One Drive Upload` - modify to accommodate your needs.
		- Currently uploads the last page of the `pdf_path` to the root folder in OneDrive.
	- `Google Drive Upload` - modify to accommodate your needs.
		- Currently uploads the `pdf_path` to specified folders saved in the `folders` variable.
			- If the file already exists, it deletes it.
			- This prevents piling of duplicates in the drive.
7. `Task Scheduler Setup`
	- This script is intended for use with the Task Scheduler.
	- Open `Task Scheduler`.
		- Create a new task with the following settings:
		- General: Check "Run with highest privileges" and configure for your version of Windows.
		- Triggers: Set your preferred schedule.
		- Actions: Start a program and point to the .bat file created above.
		- Conditions and Settings: Adjust as needed.
___
# Functions

1. **`clean_local_folder(file_path)`**:
	- **Purpose**: This function deletes a file from the local folder.
	- **Parameters**:
		- `file_path`: The path of the file to be deleted.
2. **`authenticate_onedrive()`**:
	- **Purpose**: This function handles the authentication process for accessing OneDrive using the OAuth2.
	- **Returns**: An access token for OneDrive.
3. **`exchange_code_for_tokens(code)`**:
	- **Purpose**: This function exchanges an authorization code for access and refresh tokens.
	- **Parameters**:
		- `code`: The authorization code obtained during the authentication process.
	- **Returns**: Access token, expiration time, and refresh token.
4. **`refresh_access_token(refresh_token)`**:
	- **Purpose**: This function refreshes the access token for OneDrive authentication using the refresh token.
	- **Parameters**:
		- `refresh_token`: The refresh token used to obtain a new access token.
	- **Returns**: New access token, expiration time, and possibly a new refresh token.
5. **`upload_to_onedrive(access_token)`**:
	- **Purpose**: This function uploads a PDF file to OneDrive.
	- **Parameters**:
		- `access_token`: Access token required for authentication with OneDrive.
6. **`authenticate_google_drive()`**:
	- **Purpose**: This function handles the authentication process for accessing Google Drive using the OAuth2.
	- **Returns**: Google Drive credentials for authentication.
7. **`upload_to_google_drive(credentials)`**:
	- **Purpose**: This function uploads a PDF file to Google Drive.
	- **Parameters**:
		- `credentials`: Google Drive credentials obtained during authentication.
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
