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
    SCOPES = ['https://www.googleapis.com/auth/drive']
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

# Function to check if a file is an image based on extension or MIME type
def is_image_file(file_name, mime_type=None):
    # Common image extensions
    image_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif', 
                       '.webp', '.svg', '.ico', '.heic', '.heif', '.raw', '.cr2', 
                       '.nef', '.orf', '.sr2', '.dng'}
    
    # Check by file extension
    if file_name:
        _, ext = os.path.splitext(file_name.lower())
        if ext in image_extensions:
            return True
    
    # Check by MIME type (more reliable)
    if mime_type:
        return mime_type.startswith('image/')
    
    return False

# Function to download PDF files directly
def download_pdf(pdf_url, folder_path):
    try:
        response = requests.get(pdf_url)
        response.raise_for_status()

        file_name = pdf_url.split("/")[-1]
        file_path = os.path.join(folder_path, file_name)

        os.makedirs(folder_path, exist_ok=True)

        with open(file_path, 'wb') as pdf_file:
            pdf_file.write(response.content)
        
        print(f"Downloaded PDF: {file_path}")

    except Exception as e:
        print(f"Failed to download PDF from {pdf_url}: {e}")

# Function to download file from Google Drive (simplified output)
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
        
        # Only print when download is complete
        return True

    except Exception as e:
        raise Exception(f"Failed to download {file_name}: {e}")

def download_images_in_folder(service, folder_id, folder_path, person_name="Unknown"):
    try:
        query = f"'{folder_id}' in parents"
        results = service.files().list(
            q=query,
            fields="files(id, name, mimeType, size)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()

        items = results.get('files', [])

        print(f"Processing folder for {person_name}: {len(items)} total items found")

        successful_downloads = 0
        failed_downloads = []
        
        for item in items:
            file_name = item['name']
            file_id = item['id']
            mime_type = item.get('mimeType', '')

            try:
                if is_image_file(file_name, mime_type):
                    download_file_from_gdrive(service, file_id, file_name, folder_path)
                    successful_downloads += 1
                else:
                    print(f"Skipping non-image file (trying fallback): {file_name} (MIME: {mime_type})")
                    # Fallback: attempt to download anyway
                    # download_file_from_gdrive(service, file_id, file_name, folder_path)
                    # successful_downloads += 1  # Count fallback too
            except Exception as e:
                failed_downloads.append((person_name, file_name, str(e)))

        return successful_downloads, failed_downloads

    except Exception as e:
        return 0, [(person_name, "Folder access failed", str(e))]


# Alternative function that downloads images recursively from subfolders too
def download_images_recursively(service, folder_id, folder_path, max_depth=3, current_depth=0):
    if current_depth > max_depth:
        print(f"Maximum recursion depth ({max_depth}) reached.")
        return
    
    try:
        query = f"'{folder_id}' in parents"
        results = service.files().list(
            q=query,
            fields="files(id, name, mimeType, size)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()

        items = results.get('files', [])
        
        image_count = 0
        folder_count = 0
        
        for item in items:
            file_name = item['name']
            file_id = item['id']
            mime_type = item.get('mimeType', '')
            
            # If it's a folder, recurse into it
            if mime_type == 'application/vnd.google-apps.folder':
                subfolder_path = os.path.join(folder_path, file_name)
                print(f"Entering subfolder: {file_name}")
                download_images_recursively(service, file_id, subfolder_path, max_depth, current_depth + 1)
                folder_count += 1
            # If it's an image, download it
            elif is_image_file(file_name, mime_type):
                print(f"Found image: {file_name} (MIME: {mime_type})")
                download_file_from_gdrive(service, file_id, file_name, folder_path)
                image_count += 1
            else:
                print(f"Skipping non-image file: {file_name} (MIME: {mime_type})")
        
        print(f"In current folder: Downloaded {image_count} images, found {folder_count} subfolders.")
    
    except Exception as e:
        print(f"Error while accessing folder {folder_id}: {e}")

def main():
    # Load your Excel file
    df = pd.read_excel(r'C:\Users\yewyn\Documents\Verdant\BATCH 86.xlsx')

    # Initialize Google Drive service
    service = authenticate_gdrive()

    # Ensure a base download directory
    base_download_path = './downloads/Batch 86'
    os.makedirs(base_download_path, exist_ok=True)

    # Track overall statistics
    total_successful = 0
    all_failed_downloads = []
    processed_entries = 0

    for index, row in df.iterrows():
        gdrive_link = row['G.Drive Link']
        drawing_link = row['Layout Link']
        person_name = row['Name']

        # Handle Google Drive links
        gdrive_id, gdrive_type = extract_id(gdrive_link)
        
        if gdrive_id is None:
            all_failed_downloads.append((person_name, "Invalid Google Drive link", gdrive_link))
            continue
        
        person_folder = os.path.join(base_download_path, person_name)
        processed_entries += 1

        if gdrive_type == "file":
            # Check if the single file is an image before downloading
            try:
                file_metadata = service.files().get(fileId=gdrive_id, fields='name,mimeType').execute()
                file_name = file_metadata['name']
                mime_type = file_metadata.get('mimeType', '')
                
                if is_image_file(file_name, mime_type):
                    print(f"Processing folder for {person_name}: 1 images found")
                    try:
                        download_file_from_gdrive(service, gdrive_id, file_name, person_folder)
                        total_successful += 1
                    except Exception as e:
                        all_failed_downloads.append((person_name, file_name, str(e)))
                else:
                    print(f"Processing folder for {person_name}: 0 images found (file is not an image)")
            except Exception as e:
                all_failed_downloads.append((person_name, "Error checking file", str(e)))
                
        elif gdrive_type == "folder":
            successful, failed = download_images_in_folder(service, gdrive_id, person_folder, person_name)
            total_successful += successful
            all_failed_downloads.extend(failed)
        
        # Handle PDF downloads (keep existing functionality)
        if drawing_link and drawing_link.endswith('.pdf'):
            download_pdf(drawing_link, person_folder)

    # Print final summary
    print(f"\n" + "="*60)
    print("DOWNLOAD SUMMARY")
    print("="*60)
    print(f"Total entries processed: {processed_entries}")
    print(f"Total images successfully downloaded: {total_successful}")
    print(f"Total failed downloads: {len(all_failed_downloads)}")
    
    if all_failed_downloads:
        print(f"\nFAILED DOWNLOADS ({len(all_failed_downloads)}):")
        print("-" * 60)
        for person_name, file_name, error in all_failed_downloads:
            print(f"Person: {person_name}")
            print(f"File/Issue: {file_name}")
            print(f"Error: {error}")
            print("-" * 30)
    else:
        print("\n✅ All downloads were successful!")

if __name__ == "__main__":
    main()
