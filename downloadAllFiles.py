import os
import re
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
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

    # If there are no valid credentials, ask the user to log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('drive', 'v3', credentials=creds)
    return service

# Function to extract Google Drive ID and type
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

# Function to download PDF files directly
def download_pdf(pdf_url, folder_path):
    try:
        response = requests.get(pdf_url)
        response.raise_for_status()  # Raise an error for bad responses

        # Prepare the filename
        file_name = pdf_url.split("/")[-1]
        file_path = os.path.join(folder_path, file_name)

        # Ensure the folder exists
        os.makedirs(folder_path, exist_ok=True)

        # Write the PDF to file
        with open(file_path, 'wb') as pdf_file:
            pdf_file.write(response.content)
        
        print(f"Downloaded PDF: {file_path}")

    except Exception as e:
        print(f"Failed to download PDF from {pdf_url}: {e}")

# Function to download file from Google Drive
def download_file_from_gdrive(service, file_id, file_name, folder_path):
    try:
        request = service.files().get_media(fileId=file_id)
        file_path = os.path.join(folder_path, file_name)

        # Ensure the folder exists
        os.makedirs(folder_path, exist_ok=True)

        # Downloading the file
        with open(file_path, 'wb') as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
                print(f"Download {int(status.progress() * 100)}% for {file_name}.")

    except Exception as e:
        print(f"Failed to download {file_name}: {e}")

# Function to list files in a Google Drive folder and download them
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
    # Load your Excel file
    df = pd.read_excel(r'C:\Users\yewyn\Documents\Verdant\BATCH 71.xlsx')

    # Initialize Google Drive service (assuming you have authenticated)
    service = authenticate_gdrive()

    # Ensure a base download directory
    base_download_path = './downloads'
    os.makedirs(base_download_path, exist_ok=True)

    for index, row in df.iterrows():
        gdrive_link = row['G.Drive Link']
        drawing_link = row['Layout Link']  # Assume this is the column for PDFs

        # Handle Google Drive links
        gdrive_id, gdrive_type = extract_id(gdrive_link)
        
        if gdrive_id is None:
            print(f"Invalid Google Drive link: {gdrive_link}")
            continue
        
        person_folder = os.path.join(base_download_path, row['Name'])  # Adjust as necessary

        if gdrive_type == "file":
            download_file_from_gdrive(service, gdrive_id, row['File Name'], person_folder)
        elif gdrive_type == "folder":
            download_files_in_folder(service, gdrive_id, person_folder)
        
        # Handle PDF downloads
        if drawing_link and drawing_link.endswith('.pdf'):
            download_pdf(drawing_link, person_folder)

if __name__ == "__main__":
    main()

if __name__ == "__main__":
    main()
