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
        a4_landscape = landscape(A4)
        new_page = PageObject.create_blank_page(width=a4_landscape[0], height=a4_landscape[1])
        new_page.merge_page(page)
        output_pdf.add_page(new_page)
    
    return output_pdf

def convert_image_to_pdf_with_imagemagick(image_path, output_pdf_path):
    """
    Convert image to PDF using ImageMagick (if available)
    Returns True if conversion was successful
    """
    try:
        # Check if ImageMagick is installed
        result = subprocess.run(['magick', '--version'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        
        # Use ImageMagick to convert the image to PDF
        subprocess.run([
            'magick', 
            image_path, 
            '-page', 'A4', 
            '-gravity', 'center', 
            '-background', 'white',
            '-extent', '842x595',
            output_pdf_path
        ], check=True)
        
        # Check if output file exists and has content
        if os.path.exists(output_pdf_path) and os.path.getsize(output_pdf_path) > 1000:
            print(f"Successfully converted {image_path} to PDF using ImageMagick")
            return True
        else:
            print(f"ImageMagick conversion produced an invalid PDF")
            return False
    except Exception as e:
        print(f"ImageMagick conversion failed: {str(e)}")
        return False

def convert_image_to_pdf_simple(image_path, output_pdf_path):
    """
    Simple conversion of image to PDF using PIL
    """
    try:
        # Open image and convert to RGB
        img = Image.open(image_path)
        if img.mode != 'RGB':
            img = img.convert('RGB')
        
        # Calculate proper page size while maintaining aspect ratio
        img_width, img_height = img.size
        a4_width, a4_height = landscape(A4)
        
        # Save directly to PDF (PIL has this capability)
        img.save(output_pdf_path, 'PDF', resolution=100.0)
        
        # Check if output file exists and has content
        if os.path.exists(output_pdf_path) and os.path.getsize(output_pdf_path) > 1000:
            print(f"Successfully converted {image_path} to PDF using PIL")
            return True
        else:
            print(f"PIL conversion produced an invalid PDF")
            return False
    except Exception as e:
        print(f"PIL conversion failed: {str(e)}")
        return False

def process_pdf_with_stamp_and_images(layout_pdf, image_files, stamp_image_path, output_pdf_path):
    try:
        print(f"\nProcessing layout PDF: {layout_pdf}")
        
        # Step 1: Add stamp to each page of the layout PDF
        reader = PdfReader(layout_pdf)
        stamped_pdf_writer = PdfWriter()

        for i, page in enumerate(reader.pages):
            print(f"Adding stamp to page {i+1}")
            stamped_page = add_stamp_to_page_with_precise_dpi(page, stamp_image_path)
            stamped_pdf_writer.add_page(stamped_page)
        
        # Write the stamped PDF to a temporary file
        temp_stamped_pdf_path = layout_pdf.replace(".pdf", f"_stamped_{int(time.time() * 1000)}.pdf")
        with open(temp_stamped_pdf_path, 'wb') as temp_file:
            stamped_pdf_writer.write(temp_file)
        
        # Step 2: Convert the stamped PDF to landscape A4
        converted_pdf_writer = PdfWriter()
        convert_to_landscape_a4(temp_stamped_pdf_path, converted_pdf_writer)
        
        # Create a temporary file for the converted PDF
        temp_converted_pdf_path = layout_pdf.replace(".pdf", f"_converted_{int(time.time() * 1000)}.pdf")
        with open(temp_converted_pdf_path, 'wb') as temp_file:
            converted_pdf_writer.write(temp_file)
        
        # Step 3: Process each image to a separate PDF
        image_pdfs = []
        for i, image_path in enumerate(image_files):
            try:
                print(f"Processing image {i+1}/{len(image_files)}: {image_path}")
                
                # Create output filename for this image PDF
                image_pdf_path = f"image_{int(time.time() * 1000)}_{i}.pdf"
                
                # Try different methods to convert the image to PDF
                success = False
                
                # Method 1: Try ImageMagick first (if available)
                success = convert_image_to_pdf_with_imagemagick(image_path, image_pdf_path)
                
                # Method 2: If ImageMagick failed, try direct PIL conversion
                if not success:
                    success = convert_image_to_pdf_simple(image_path, image_pdf_path)
                
                # If conversion succeeded, add to our list
                if success and os.path.exists(image_pdf_path) and os.path.getsize(image_pdf_path) > 0:
                    image_pdfs.append(image_pdf_path)
                    print(f"Successfully added {os.path.basename(image_path)} to PDF list")
                else:
                    print(f"Failed to convert {os.path.basename(image_path)} to PDF")
                
            except Exception as e:
                print(f"Error processing image {image_path}: {str(e)}")
        
        # Step 4: Combine all PDFs (converted layout PDF + image PDFs)
        all_pdfs = [temp_converted_pdf_path] + image_pdfs
        
        # If no PDFs were created, return error
        if len(all_pdfs) == 0:
            print("Error: No PDFs were created!")
            return False
        
        # Combine all PDFs using PyPDF2
        final_pdf_writer = PdfWriter()
        
        # First add all pages from the converted layout PDF
        layout_reader = PdfReader(temp_converted_pdf_path)
        for page in layout_reader.pages:
            final_pdf_writer.add_page(page)
        
        # Then add all pages from each image PDF
        for image_pdf in image_pdfs:
            try:
                image_reader = PdfReader(image_pdf)
                for page in image_reader.pages:
                    final_pdf_writer.add_page(page)
                print(f"Added pages from {image_pdf}")
            except Exception as e:
                print(f"Error adding pages from {image_pdf}: {str(e)}")
        
        # Write the final combined PDF
        print(f"Writing final combined PDF to {output_pdf_path}")
        with open(output_pdf_path, 'wb') as final_file:
            final_pdf_writer.write(final_file)
        
        # Clean up temporary files
        for temp_file in [temp_stamped_pdf_path, temp_converted_pdf_path] + image_pdfs:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
            except Exception as e:
                print(f"Error removing temporary file {temp_file}: {str(e)}")
        
        print(f"Successfully completed processing: {layout_pdf}")
        return True
        
    except Exception as e:
        print(f"Error processing PDF {layout_pdf}: {str(e)}")
        return False

def process_all_subfolders(batch_folder, output_folder, stamp_image_path):
    print(f"Starting batch processing from {batch_folder} to {output_folder}")
    print(f"Using stamp: {stamp_image_path}")
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"Created output folder: {output_folder}")
    
    # Check if stamp image exists
    if not os.path.exists(stamp_image_path):
        print(f"Error: Stamp image does not exist: {stamp_image_path}")
        return
    
    subfolder_count = 0
    success_count = 0
    error_count = 0
    
    for subfolder_name in os.listdir(batch_folder):
        subfolder_path = os.path.join(batch_folder, subfolder_name)
        if os.path.isdir(subfolder_path):
            subfolder_count += 1
            print(f"\nProcessing subfolder {subfolder_count}: {subfolder_name}")
            
            layout_pdf = None
            image_files = []
            
            for file_name in os.listdir(subfolder_path):
                file_path = os.path.join(subfolder_path, file_name)
                if "Layout" in file_name and file_name.lower().endswith(".pdf"):
                    layout_pdf = file_path
                    print(f"Found layout PDF: {file_name}")
                elif file_name.lower().endswith((".png", ".jpg", ".jpeg")):
                    image_files.append(file_path)
                    print(f"Found image: {file_name}")
            
            if layout_pdf and image_files:
                # Sort images by filename
                image_files.sort()
                print(f"Total images found: {len(image_files)}")
                
                base_name = os.path.splitext(os.path.basename(layout_pdf))[0]
                endorsed_pdf_name = f"{base_name} - Endorsed.pdf"
                output_pdf_path = os.path.join(output_folder, endorsed_pdf_name)
                
                print(f"Processing: {base_name}")
                print(f"Output will be saved to: {output_pdf_path}")
                
                # Process PDF with stamp and images
                result = process_pdf_with_stamp_and_images(layout_pdf, image_files, stamp_image_path, output_pdf_path)
                
                if result:
                    success_count += 1
                    print(f"Successfully processed: {subfolder_name}")
                else:
                    error_count += 1
                    print(f"Error processing: {subfolder_name}")
            else:
                if not layout_pdf:
                    print(f"No layout PDF found in subfolder: {subfolder_name}")
                if not image_files:
                    print(f"No images found in subfolder: {subfolder_name}")
    
    print("\nBatch processing complete")
    print(f"Total subfolders processed: {subfolder_count}")
    print(f"Successful: {success_count}")
    print(f"Errors: {error_count}")

# Example usage
if __name__ == "__main__":
    batch_folder = r'C:\Users\yewyn\Documents\Verdant\Batch 71'
    output_folder = r'C:\Users\yewyn\Documents\Verdant\Batch 71 Retry'
    stamp_image_path = r'C:\Users\yewyn\Documents\Verdant\newStamp.png'

    process_all_subfolders(batch_folder, output_folder, stamp_image_path)