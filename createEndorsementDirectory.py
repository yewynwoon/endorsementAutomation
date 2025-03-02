import os
import pandas as pd
import re

# Load the Excel file (this would be replaced by the actual file path)
file_path = r'C:\Users\yewyn\Documents\Verdant\BATCH 71.xlsx'
excel_data = pd.read_excel(file_path)

# Define the output directory where you want to create the folders
output_directory = r'C:\Users\yewyn\Documents\Verdant\Batch 71'  # You can change this path

# Ensure the output directory exists
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# Function to sanitize folder names
def sanitize_folder_name(name):
    # Replace slashes, tabs, and other problematic characters with underscores
    sanitized_name = re.sub(r'[<>:"/\\|?*\t\n\r]', '_', name)
    return sanitized_name.strip()

# Loop through the Excel data and create directories
for index, row in excel_data.iterrows():
    sanitized_name = sanitize_folder_name(row['Name'])
    folder_name = f"{index + 1}_{sanitized_name}"  # Create folder name with index and sanitized name
    folder_path = os.path.join(output_directory, folder_name)  # Full path of the folder
    
    # Create the directory
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"Created folder: {folder_name}")

print(f"All folders created in {output_directory}")


