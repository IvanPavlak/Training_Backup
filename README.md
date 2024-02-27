# Training Backup
# Requirements

## One Drive

- [Microsoft Azure](https://azure.microsoft.com/en-us)
	- [Azure App Signup step by step](https://github.com/pranabdas/Access-OneDrive-via-Microsoft-Graph-Python/blob/main/Azure_app_signup_step_by_step.md)
	- Client ID
	- Client Secret -> Created additionally in "Certificates & Secrets" section.
	- Both must be copied into a `onedrive_credentials.json`:
```
{
    "client_id": "client_id",
    "client_secret": "secret_value"
}
```

## Google Drive

- [Google Cloud](https://cloud.google.com/gcp?utm_source=google&utm_medium=cpc&utm_campaign=emea-emea-all-en-bkws-all-all-trial-e-gcp-1707574&utm_content=text-ad-none-any-DEV_c-CRE_683761846512-ADGP_Hybrid+%7C+BKWS+-+EXA+%7C+Txt+-+GCP+-+General+-+v2-KWID_43700078882258013-kwd-6458750523-userloc_1007612&utm_term=KW_google%20cloud-NET_g-PLAC_&&gad_source=1&gclid=CjwKCAiAivGuBhBEEiwAWiFmYSEVAU4nVtvqTjYCKbWC08C1ap_UukXjFhKNnvw9t3uknDf6DtumLBoCJTwQAvD_BwE&gclsrc=aw.ds)
	- Console -> APIs & Services -> OAuth consent screen -> Create a Project -> User Type: External -> Add yourself as Test User -> Credentials -> Create Credentials: OAuth client ID -> Desktop app -> save the `json` file to `google_credentials.json`
 
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
2. Copy the `TrainingBackupCredentials` folder in it.
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
		- Currently uploads the `pdf_path` to a drive named `PRogram1`.
			- If the file already exists, it deletes it.
			- This prevents piling of duplicates in the drive.
1. `Task Scheduler`
	- This script is intended for use with the Task Scheduler.
	- Example usage is:
		- Open `Task Scheduler` -> `Create Task` -> Check `Run with highest privileges` -> Configure for `Windows 10` -> Set trigger as preferred -> Set Action to `Start a program` with the path to the `.bat` file (`C:\folder\subfolder\TrainingBackup.bat`) -> Set Conditions and Settings as preferred.
___
# Functions

1. **`clean_local_folder(file_path)`**:
	- **Purpose**: This function deletes a file from the local folder.
	- **Parameters**:
		- `file_path`: The path of the file to be deleted.
2. **`authenticate_onedrive()`**:
	- **Purpose**: This function handles the authentication process for accessing OneDrive.
	- **Returns**: An access token for OneDrive authentication.
3. **`exchange_code_for_tokens(code)`**:
	- **Purpose**: This function exchanges an authorization code for access and refresh tokens during the OneDrive authentication process.
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
	- **Purpose**: This function handles the authentication process for accessing Google Drive.
	- **Returns**: Google Drive credentials for authentication.
7. **`upload_to_google_drive(credentials)`**:
	- **Purpose**: This function uploads a PDF file to Google Drive.
	- **Parameters**:
		- `credentials`: Google Drive credentials obtained during authentication.
___
# Error Handling

- **Authentication Errors**:
	- If authentication with OneDrive or Google Drive fails due to invalid credentials, expired tokens, or network issues, the script will catch these errors and prompt the user to reauthenticate. It will provide instructions on how to obtain new credentials or resolve authentication issues.

- **Upload Errors**:
	- If there are errors during the upload process to OneDrive or Google Drive, such as network interruptions or permission errors, the script will attempt to retry the upload process a few times before notifying the user of the failure and providing guidance on resolving the issue.

- **File Path Errors**:
	- If the specified file paths are incorrect or if the files are not found at the specified locations, the script will catch these errors and inform the user. It will provide guidance on verifying file paths and ensuring that the necessary files exist before proceeding with conversion and upload operations.
___
# Security Considerations:

- **Credential Management**:
	- The script securely manages credentials for accessing OneDrive and Google Drive by storing them in separate JSON files (`onedrive_credentials.json` and `google_credentials.json`). These files should be stored in a secure location and access restricted to authorized users only.

- **Token Handling**:
	- Tokens obtained during authentication, such as access tokens and refresh tokens, are handled securely within the script. The script ensures that tokens are not exposed in log files or error messages and implements token rotation and expiration strategies to minimize the risk of unauthorized access.

- **HTTPS Usage**:
	- All communications with external services, such as OneDrive and Google Drive APIs, are performed over HTTPS to encrypt data in transit and prevent eavesdropping or tampering by malicious actors.

- **Input Sanitization**:
	- User inputs, such as file paths and authentication codes, are validated and sanitized to prevent injection attacks or path traversal vulnerabilities. The script ensures that user-supplied data is properly sanitized before processing to mitigate security risks.

- **Error Logging**:
	- The script implements error logging to record any unexpected errors or exceptions that occur during execution. Log files are stored securely, and regular reviews are conducted to identify potential security issues or anomalies.
___
# Helpful Sources

- These were studied, used and modified to accomplish the desired effect:
#### Google Drive Backup:
- [Automated Google Drive Backups in Python](https://www.youtube.com/watch?v=fkWM7A-MxR0)
#### One Drive Backup:
- [Access-OneDrive-via-Microsoft-Graph-Python](https://github.com/pranabdas/Access-OneDrive-via-Microsoft-Graph-Python)
___
