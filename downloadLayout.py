import os
import re
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import requests

def authenticate_gdrive():
    SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
    creds = None

    # Token file stores the user's access and refresh tokens
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If no valid credentials, log in again
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Save credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('drive', 'v3', credentials=creds)
    return service

# Extract Google Drive ID and type
def extract_id(link):
    file_pattern = r"https:\/\/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)"
    folder_pattern = r"https:\/\/drive\.google\.com\/drive\/folders\/([a-zA-Z0-9_-]+)"
    
    match_file = re.match(file_pattern, link)
    match_folder = re.match(folder_pattern, link)
    
    if match_file:
        return match_file.group(1), "file"
    elif match_folder:
        return match_folder.group(1), "folder"
    else:
        return None, None

# Download PDF files directly from link
def download_pdf(pdf_url, folder_path):
    try:
        response = requests.get(pdf_url)
        response.raise_for_status()

        # Prepare the filename and ensure folder exists
        file_name = pdf_url.split("/")[-1]
        file_path = os.path.join(folder_path, file_name)
        os.makedirs(folder_path, exist_ok=True)

        with open(file_path, 'wb') as pdf_file:
            pdf_file.write(response.content)
        
        print(f"Downloaded PDF: {file_path}")
    except Exception as e:
        print(f"Failed to download PDF from {pdf_url}: {e}")

# Download file from Google Drive
def download_file_from_gdrive(service, file_id, file_name, folder_path):
    try:
        request = service.files().get_media(fileId=file_id)
        file_path = os.path.join(folder_path, file_name)
        os.makedirs(folder_path, exist_ok=True)

        with open(file_path, 'wb') as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
                print(f"Download {int(status.progress() * 100)}% for {file_name}.")
    except Exception as e:
        print(f"Failed to download {file_name}: {e}")

# List files in Google Drive folder and download them
def download_files_in_folder(service, folder_id, folder_path):
    try:
        query = f"'{folder_id}' in parents"
        results = service.files().list(q=query).execute()
        items = results.get('files', [])
        
        for item in items:
            file_name = item['name']
            file_id = item['id']
            download_file_from_gdrive(service, file_id, file_name, folder_path)
    
    except Exception as e:
        print(f"Error while accessing folder {folder_id}: {e}")

import os
import re
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import requests

def authenticate_gdrive():
    SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
    creds = None

    # Token file stores the user's access and refresh tokens
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If no valid credentials, log in again
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Save credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('drive', 'v3', credentials=creds)
    return service

# Extract Google Drive ID and type
def extract_id(link):
    file_pattern = r"https:\/\/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)"
    folder_pattern = r"https:\/\/drive\.google\.com\/drive\/folders\/([a-zA-Z0-9_-]+)"
    
    match_file = re.match(file_pattern, link)
    match_folder = re.match(folder_pattern, link)
    
    if match_file:
        return match_file.group(1), "file"
    elif match_folder:
        return match_folder.group(1), "folder"
    else:
        return None, None

# Download PDF files directly from link
def download_pdf(pdf_url, folder_path):
    try:
        response = requests.get(pdf_url)
        response.raise_for_status()

        # Prepare the filename and ensure folder exists
        # Only take the file ID part from the URL for naming, remove any query parameters
        file_name = pdf_url.split("/")[-2]  # This extracts the ID part from the URL and appends .pdf
        file_path = os.path.join(folder_path, file_name)
        os.makedirs(folder_path, exist_ok=True)

        with open(file_path, 'wb') as pdf_file:
            pdf_file.write(response.content)
        
        print(f"Downloaded PDF: {file_path}")
    except Exception as e:
        print(f"Failed to download PDF from {pdf_url}: {e}")

# Download file from Google Drive
def download_file_from_gdrive(service, file_id, file_name, folder_path):
    try:
        request = service.files().get_media(fileId=file_id)
        file_path = os.path.join(folder_path, file_name)
        os.makedirs(folder_path, exist_ok=True)

        with open(file_path, 'wb') as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
                print(f"Download {int(status.progress() * 100)}% for {file_name}.")
    except Exception as e:
        print(f"Failed to download {file_name}: {e}")

# List files in Google Drive folder and download them
def download_files_in_folder(service, folder_id, folder_path):
    try:
        query = f"'{folder_id}' in parents"
        results = service.files().list(q=query).execute()
        items = results.get('files', [])
        
        for item in items:
            file_name = item['name']
            file_id = item['id']
            download_file_from_gdrive(service, file_id, file_name, folder_path)
    
    except Exception as e:
        print(f"Error while accessing folder {folder_id}: {e}")

def main():
    # Load the Excel file
    df = pd.read_excel(r'C:\Users\yewyn\Documents\Verdant\BATCH 62.xlsx')

    # Initialize Google Drive service
    service = authenticate_gdrive()

    base_download_path = './layouts'
    os.makedirs(base_download_path, exist_ok=True)

    for index, row in df.iterrows():
        drawing_link = row['Layout Link']  # Assuming this is the column for PDFs/Google Drive links
        if drawing_link is None:
            print(f"Invalid Google Drive link: {drawing_link}")
            continue
        
        person_folder = os.path.join(base_download_path, row['Name'])  # Adjust folder structure as needed

        # Check the type of the Google Drive link
        gdrive_id, gdrive_type = extract_id(drawing_link)
        
        # if drawing_link.endswith('.pdf'):
        download_pdf(drawing_link, person_folder)
        # else:
        #     print(f"Unsupported link format: {drawing_link}")

if __name__ == "__main__":
    main()

