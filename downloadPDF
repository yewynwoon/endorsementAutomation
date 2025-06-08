import os
import re
import pandas as pd
import requests
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError
import pickle
import time

# Function to extract Google Drive file ID
def extract_drive_id(link):
    # Updated pattern to handle various Google Drive URL formats
    patterns = [
        r"https:\/\/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)\/view",  # /view format
        r"https:\/\/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)\?",      # with query params
        r"https:\/\/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)\/edit",  # /edit format
        r"https:\/\/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)$",       # basic format
        r"https:\/\/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)\/",      # with trailing slash
        r"https:\/\/drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)",        # /open?id= format
        r"https:\/\/docs\.google\.com\/.*[?&]id=([a-zA-Z0-9_-]+)",         # docs.google.com format
        r"https:\/\/drive\.google\.com\/.*[?&]id=([a-zA-Z0-9_-]+)"         # any drive.google.com with id parameter
    ]
    
    for pattern in patterns:
        match = re.search(pattern, link)
        if match:
            return match.group(1)
    
    return None

# Function to authenticate with Google Drive
def authenticate_gdrive():
    SCOPES = ['https://www.googleapis.com/auth/drive']  # Full access instead of readonly
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

def check_gdrive_file_exists(service, file_id):
    try:
        file_metadata = service.files().get(fileId=file_id, fields='name,permissions').execute()
        return True, file_metadata['name']
    except HttpError as e:
        if e.resp.status == 404:
            return False, "File not found (404) - File may have been deleted or moved"
        elif e.resp.status == 403:
            return False, "Permission denied (403) - Check sharing permissions"
        else:
            return False, f"HTTP Error {e.resp.status}: {e._get_reason()}"

# Function to download a Google Drive file using file ID
def download_file_from_gdrive(service, file_id, folder_path, max_retries=3):
    file_name = None
    
    # Try to get the original filename from API first
    try:
        file_metadata = service.files().get(fileId=file_id, fields='name').execute()
        file_name = file_metadata['name']
        print(f"Found original filename: {file_name}")
    except HttpError as e:
        if e.resp.status == 404:
            print(f"API can't access file {file_id}, but will try direct download")
        elif e.resp.status == 403:
            print(f"API permission denied for {file_id}, but will try direct download")
        # Continue with direct download even if API fails
    except Exception as e:
        print(f"API error for {file_id}: {str(e)}, continuing with direct download")
    
    # Attempt direct download
    for attempt in range(max_retries):
        try:
            file_url = f"https://drive.google.com/uc?export=download&id={file_id}"
            response = requests.get(file_url, timeout=30)
            
            # Handle Google Drive's redirect for large files
            if response.status_code == 200:
                # If we didn't get filename from API, try to extract from response headers
                if not file_name:
                    content_disposition = response.headers.get('content-disposition', '')
                    if 'filename=' in content_disposition:
                        # Extract filename from Content-Disposition header
                        file_name = content_disposition.split('filename=')[1].strip('"\'')
                    else:
                        # Fallback to file_id with .pdf extension
                        file_name = f"{file_id}.pdf"
                
                # Ensure the filename is safe for the filesystem
                file_name = sanitize_filename(file_name)
                file_path = os.path.join(folder_path, file_name)
                
                with open(file_path, 'wb') as drive_file:
                    drive_file.write(response.content)
                
                print(f"Downloaded Google Drive file: {file_path}")
                return True, None
                
            elif response.status_code == 303:
                # Handle redirect - Google sometimes returns 303 for large files
                print(f"Received redirect (303) for {file_id}, following redirect...")
                redirect_url = response.headers.get('Location')
                if redirect_url:
                    redirect_response = requests.get(redirect_url, timeout=30)
                    if redirect_response.status_code == 200:
                        if not file_name:
                            file_name = f"{file_id}.pdf"
                        file_name = sanitize_filename(file_name)
                        file_path = os.path.join(folder_path, file_name)
                        
                        with open(file_path, 'wb') as drive_file:
                            drive_file.write(redirect_response.content)
                        
                        print(f"Downloaded Google Drive file: {file_path}")
                        return True, None
            else:
                error_msg = f"HTTP {response.status_code}: Could not download file"
                if attempt < max_retries - 1:
                    print(f"Download failed for {file_id}, retrying in 2 seconds...")
                    time.sleep(2)
                else:
                    return False, error_msg
                    
        except requests.exceptions.Timeout:
            error_msg = f"Timeout error (attempt {attempt + 1}/{max_retries})"
            if attempt < max_retries - 1:
                print(f"Timeout downloading {file_id}, retrying in 2 seconds...")
                time.sleep(2)
            else:
                return False, f"Failed after {max_retries} attempts: {error_msg}"
                
        except Exception as e:
            error_msg = f"Unexpected error: {str(e)}"
            if attempt < max_retries - 1:
                print(f"Error downloading {file_id}, retrying in 2 seconds...")
                time.sleep(2)
            else:
                return False, error_msg
    
    return False, "Max retries exceeded"

# Helper function to sanitize filenames for filesystem compatibility
def sanitize_filename(filename):
    # Remove or replace characters that are invalid for Windows filenames
    invalid_chars = r'[<>:"/\\|?*]'
    sanitized = re.sub(invalid_chars, '_', filename)
    # Remove leading/trailing spaces and dots
    sanitized = sanitized.strip('. ')
    # Ensure filename isn't empty
    if not sanitized:
        sanitized = "unnamed_file"
    return sanitized

# Function to download PDF files directly from URL
def download_pdf(pdf_url, folder_path, max_retries=3):
    for attempt in range(max_retries):
        try:
            response = requests.get(pdf_url, timeout=30)
            response.raise_for_status()

            # Prepare the filename
            file_name = pdf_url.split("/")[-1]
            if not file_name.endswith('.pdf'):
                file_name += '.pdf'
            file_path = os.path.join(folder_path, file_name)

            # Write the PDF to file
            with open(file_path, 'wb') as pdf_file:
                pdf_file.write(response.content)
            
            print(f"Downloaded PDF: {file_path}")
            return True, None

        except requests.exceptions.Timeout:
            error_msg = f"Timeout error (attempt {attempt + 1}/{max_retries})"
            if attempt < max_retries - 1:
                print(f"Timeout downloading {pdf_url}, retrying in 2 seconds...")
                time.sleep(2)
            else:
                return False, f"Failed after {max_retries} attempts: {error_msg}"
                
        except Exception as e:
            return False, f"Error downloading PDF: {str(e)}"
    
    return False, "Max retries exceeded"

# Function to sanitize folder names (removes invalid characters for Windows)
def sanitize_folder_name(name):
    # Remove or replace characters that are invalid for Windows folder names
    invalid_chars = r'[<>:"/\\|?*\'()]'
    sanitized_name = re.sub(invalid_chars, '', name)
    return re.sub(r'\s+', ' ', sanitized_name).strip()

# Function to save failed downloads to a CSV file for easy review
def save_failed_downloads_to_csv(failed_downloads, filename="failed_downloads.csv"):
    if failed_downloads:
        df_failed = pd.DataFrame(failed_downloads, columns=['No', 'Name', 'Link', 'Error'])
        df_failed.to_csv(filename, index=False)
        print(f"\nFailed downloads saved to: {filename}")

def main():
    # Load your Excel file
    df = pd.read_excel(r'C:\Users\yewyn\Documents\Verdant\Batch 86.xlsx')

    # Authenticate with Google Drive
    try:
        service = authenticate_gdrive()
        print("Successfully authenticated with Google Drive")
    except Exception as e:
        print(f"Failed to authenticate with Google Drive: {e}")
        return

    # Ensure a base download directory
    base_download_path = './Batch 86'
    os.makedirs(base_download_path, exist_ok=True)

    # Counters for summary
    total_files = 0
    successful_downloads = 0
    failed_downloads = []

    print(f"Starting download for {len(df)} entries...\n")

    for index, row in df.iterrows():
        drawing_links = row['Layout Link']
        folder_no = row['No']
        person_name = row['Name']

        # Replace any '/' in the name with a space and sanitize the name
        person_name = person_name.replace('/', ' ')
        person_name = sanitize_folder_name(person_name)

        # Create folder name by appending 'No' value before the name
        person_folder = os.path.join(base_download_path, f"{folder_no}_{person_name}")
        os.makedirs(person_folder, exist_ok=True)

        # Handle links (multiple links in "Layout Link")
        if pd.notna(drawing_links):
            pdf_urls = drawing_links.split(',')
            
            for link in pdf_urls:
                link = link.strip()
                total_files += 1

                # Check if the link is a Google Drive link
                drive_id = extract_drive_id(link)
                success = False
                error = None

                if drive_id:
                    print(f"Processing Google Drive file: {drive_id}")
                    success, error = download_file_from_gdrive(service, drive_id, person_folder)
                elif link.endswith('.pdf'):
                    print(f"Processing direct PDF: {link}")
                    success, error = download_pdf(link, person_folder)
                else:
                    print(f"Skipping non-PDF and non-Google Drive link: {link}")
                    continue

                if success:
                    successful_downloads += 1
                else:
                    failed_downloads.append((row['No'], row['Name'], link, error))

        # Progress indicator
        if (index + 1) % 5 == 0:
            print(f"Processed {index + 1}/{len(df)} entries...")

    # Print summary
    print(f"\n" + "="*50)
    print("DOWNLOAD SUMMARY")
    print("="*50)
    print(f"Total files processed: {total_files}")
    print(f"Successful downloads: {successful_downloads}")
    print(f"Failed downloads: {len(failed_downloads)}")
    print(f"Success rate: {(successful_downloads/total_files*100):.1f}%" if total_files > 0 else "No files processed")

    # Print and save failed downloads
    if failed_downloads:
        print(f"\nFailed Layout Downloads ({len(failed_downloads)}):")
        print("-" * 50)
        for failure in failed_downloads:
            print(f"No: {failure[0]}")
            print(f"Name: {failure[1]}")
            print(f"Link: {failure[2]}")
            print(f"Error: {failure[3]}")
            print("-" * 30)
        
        # Save failed downloads to CSV for review
        save_failed_downloads_to_csv(failed_downloads)
    else:
        print("\n✅ All layout downloads were successful!")

if __name__ == "__main__":
    main()
