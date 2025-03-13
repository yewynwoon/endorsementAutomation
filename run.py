import os
import shutil
import subprocess
import sys
from PyPDF2 import PdfReader, PdfWriter, PageObject
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, A4
from io import BytesIO
import time

def add_stamp_to_page_with_precise_dpi(page, stamp_image_path):
    # Create a new PDF to hold the stamp
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=page.mediabox.upper_right)
    
    # Load the stamp image
    stamp_image = Image.open(stamp_image_path)
    
    # Ensure the image is in a compatible mode
    if stamp_image.mode != 'RGBA':
        stamp_image = stamp_image.convert('RGBA')
    
    # Set the exact DPI values from the reference stamp
    dpi_x, dpi_y = 295, 301
    
    # Convert pixel dimensions to points (72 points per inch)
    width_inches = stamp_image.size[0] / dpi_x
    height_inches = stamp_image.size[1] / dpi_y
    
    # Convert to points (1 inch = 72 points)
    width_points = width_inches * 72
    height_points = height_inches * 72
    
    # Get the dimensions of the page
    page_width = float(page.mediabox.upper_right[0])
    page_height = float(page.mediabox.upper_right[1])
    
    # Fixed position for consistent placement
    x_position = page_width - width_points - 12
    y_position = page_height - height_points - 47
    
    # Save the stamp image temporarily in PNG format
    temp_stamp_path = f"temp_stamp_{int(time.time() * 1000)}.png"
    stamp_image.save(temp_stamp_path, format='PNG')
    
    # Draw the stamp with precise dimensions and transparency support
    c.drawImage(temp_stamp_path, x_position, y_position, width=width_points, height=height_points, mask='auto')
    c.save()
    
    # Clean up temporary file
    os.remove(temp_stamp_path)
    
    # Move to the beginning of the BytesIO buffer
    packet.seek(0)
    
    # Create new PDF with the stamp
    new_pdf = PdfReader(packet)
    new_page = new_pdf.pages[0]
    
    # Merge the stamped page with the original
    page.merge_page(new_page)
    
    return page

def convert_to_landscape_a4(pdf_path, output_pdf):
    reader = PdfReader(pdf_path)
    
    for page in reader.pages:
        # Convert page to landscape A4
        a4_landscape = landscape(A4)
        new_page = PageObject.create_blank_page(width=a4_landscape[0], height=a4_landscape[1])
        new_page.merge_page(page)
        output_pdf.add_page(new_page)
    
    return output_pdf

def process_pdf_with_stamp_and_images(layout_pdf, image_files, stamp_image_path, output_pdf_path):
    # Step 1: Add stamp to each page with the correct DPI
    reader = PdfReader(layout_pdf)
    stamped_pdf_writer = PdfWriter()

    for page in reader.pages:
        stamped_page = add_stamp_to_page_with_precise_dpi(page, stamp_image_path)
        stamped_pdf_writer.add_page(stamped_page)
    
    # Write the stamped PDF to a temporary file
    temp_stamped_pdf_path = layout_pdf.replace(".pdf", "_stamped.pdf")
    with open(temp_stamped_pdf_path, 'wb') as temp_file:
        stamped_pdf_writer.write(temp_file)

    # Step 2: Convert the stamped PDF to landscape A4
    converted_pdf_writer = PdfWriter()
    convert_to_landscape_a4(temp_stamped_pdf_path, converted_pdf_writer)

    # Step 3: Append images to the converted PDF
    for image_path in image_files:
        image_pdf_path = "image.pdf"
        c = canvas.Canvas(image_pdf_path, pagesize=landscape(A4))

        # Adjust the position and size to fit the entire page
        image = Image.open(image_path)
        image_width, image_height = image.size
        page_width, page_height = landscape(A4)

        # Calculate aspect ratio
        aspect_ratio = image_width / image_height
        page_aspect_ratio = page_width / page_height

        # Determine the dimensions for the image
        if aspect_ratio > page_aspect_ratio:
            # Image is wider than the page
            scaled_width = page_width
            scaled_height = page_width / aspect_ratio
        else:
            # Image is taller than the page
            scaled_height = page_height
            scaled_width = page_height * aspect_ratio

        # Center the image on the page
        x_position = (page_width - scaled_width) / 2
        y_position = (page_height - scaled_height) / 2

        # Draw the image on the canvas
        c.drawImage(image_path, x_position, y_position, width=scaled_width, height=scaled_height)
        c.save()

        image_reader = PdfReader(image_pdf_path)
        image_page = image_reader.pages[0]
        converted_pdf_writer.add_page(image_page)
        os.remove(image_pdf_path)

    # Write the final output PDF
    with open(output_pdf_path, 'wb') as output_file:
        converted_pdf_writer.write(output_file)

    # Clean up temporary files
    os.remove(temp_stamped_pdf_path)

def process_all_subfolders(batch_folder, output_folder, stamp_image_path):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    no_output_folders = []  # List to store subfolders with no output

    for subfolder_name in os.listdir(batch_folder):
        subfolder_path = os.path.join(batch_folder, subfolder_name)
        if os.path.isdir(subfolder_path):
            layout_pdf = None
            image_files = []
            
            for file_name in os.listdir(subfolder_path):
                file_path = os.path.join(subfolder_path, file_name)
                if file_name.lower().endswith(".pdf"):
                    layout_pdf = file_path
                elif file_name.lower().endswith((".png", ".jpg", ".jpeg")):
                    image_files.append(file_path)
            
            if layout_pdf and image_files:
                # Sort images by filename
                image_files.sort()

                base_name = os.path.splitext(os.path.basename(layout_pdf))[0]
                endorsed_pdf_name = f"{base_name} - Endorsed.pdf"
                output_pdf_path = os.path.join(output_folder, endorsed_pdf_name)
                
                # Process PDF with stamp and images
                process_pdf_with_stamp_and_images(layout_pdf, image_files, stamp_image_path, output_pdf_path)
                
                print(f"Processed {layout_pdf}")
            else:
                # If no PDF or no images, add the subfolder to the no_output_folders list
                no_output_folders.append(subfolder_name)

    # Print folders with no output at the end
    if no_output_folders:
        print("\nSubfolders with no output:")
        for folder in no_output_folders:
            print(folder)

# Example usage
if __name__ == "__main__":
    batch_folder = r'C:\Users\nb1633\Documents\Batch 76'
    output_folder = r'C:\Users\nb1633\Documents\Batch 76 Endorsed'
    stamp_image_path = r'C:\Users\nb1633\Documents\newStamp.png'

    process_all_subfolders(batch_folder, output_folder, stamp_image_path)
