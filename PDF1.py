import fitz  # PyMuPDF
from PIL import Image
import os
import argparse
import sys
import traceback
import time
import logging

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("pdf_debug.log"),
        logging.StreamHandler()
    ]
)

# Updated base folder paths
BASE_FOLDER = 'C:/Users/john/Downloads/Folder_For_Uploader'
TEMPLATE_IMAGES_FOLDER = os.path.join(BASE_FOLDER, 'Templateimages')
SCRIPTS_FOLDER = os.path.join(BASE_FOLDER, 'Uptodatescripts')

# Footer path (specifically from the scripts folder as requested)
FOOTER_PATH = os.path.join(SCRIPTS_FOLDER, 'Footer.pdf')

def verify_directories():
    """Verify all necessary directories exist and are accessible."""
    directories = [BASE_FOLDER, TEMPLATE_IMAGES_FOLDER, SCRIPTS_FOLDER]
    for directory in directories:
        if not os.path.exists(directory):
            logging.error(f"Directory not found: {directory}")
            try:
                os.makedirs(directory, exist_ok=True)
                logging.info(f"Created directory: {directory}")
            except Exception as e:
                logging.error(f"Failed to create directory {directory}: {e}")
                return False
                
        # Test write permissions by creating a test file
        try:
            test_file = os.path.join(directory, "test_write.tmp")
            with open(test_file, 'w') as f:
                f.write("test")
            os.remove(test_file)
            logging.info(f"Directory is writable: {directory}")
        except Exception as e:
            logging.error(f"Directory is not writable: {directory}, Error: {e}")
            return False
            
    return True

def create_tiled_image_pdf(output_pdf_path, image_path, template_width, template_height, dpi=300, horizontal_repeats=6):
    try:
        logging.info(f"Creating tiled PDF for {image_path}")
        logging.info(f"Output path: {output_pdf_path}")
        
        # Ensure output directory exists
        output_dir = os.path.dirname(output_pdf_path)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
            logging.info(f"Created output directory: {output_dir}")
        
        # Verify image file exists
        if not os.path.exists(image_path):
            logging.error(f"ERROR: Image file not found at {image_path}")
            return False
        
        # Verify image is valid
        try:
            image = Image.open(image_path)
            image_format = image.format
            image_size = image.size
            logging.info(f"Image format: {image_format}, size: {image_size}")
        except Exception as e:
            logging.error(f"ERROR: Could not open image {image_path}: {e}")
            return False

        # Create a new document
        try:
            doc = fitz.open()
            page = doc.new_page(width=template_width, height=template_height)
            logging.info(f"Created new document page: {template_width}x{template_height}")
        except Exception as e:
            logging.error(f"ERROR: Failed to create document: {e}")
            return False

        # Calculate tile dimensions
        tile_width = template_width / horizontal_repeats
        tile_height = tile_width / (image.width / image.height)  # Maintain aspect ratio
        logging.info(f"Calculated tile dimensions: {tile_width}x{tile_height}")

        # Convert to pixels for resizing
        tile_width_px = int(tile_width * (dpi / 72))
        tile_height_px = int(tile_height * (dpi / 72))
        logging.info(f"Pixel dimensions for {dpi} DPI: {tile_width_px}x{tile_height_px}")

        # Create a temp directory for intermediary files if it doesn't exist
        temp_dir = os.path.join(os.path.dirname(output_pdf_path), "temp")
        os.makedirs(temp_dir, exist_ok=True)
        
        scaled_image_path = os.path.join(temp_dir, f"scaled_image_{os.path.basename(image_path)}")
        logging.info(f"Temp scaled image path: {scaled_image_path}")
        
        try:
            image = image.resize((tile_width_px, tile_height_px), Image.Resampling.LANCZOS)
            image.save(scaled_image_path, "PNG")
            logging.info(f"Saved resized image to {scaled_image_path}")
        except Exception as e:
            logging.error(f"ERROR: Could not resize or save image: {e}")
            return False

        # Verify scaled image exists
        if not os.path.exists(scaled_image_path):
            logging.error(f"ERROR: Scaled image not found at {scaled_image_path}")
            return False

        # Insert tiles
        y = 0
        insertion_count = 0
        while y < template_height:
            x = 0
            while x + tile_width <= template_width:  # Ensure correct number of repetitions
                rect = fitz.Rect(x, y, x + tile_width, y + tile_height)
                try:
                    page.insert_image(rect, filename=scaled_image_path)
                    insertion_count += 1
                except Exception as e:
                    logging.error(f"ERROR: Failed to insert image at position {x},{y}: {e}")
                x += tile_width
            y += tile_height

        # Save the PDF
        try:
            doc.save(output_pdf_path)
            doc.close()
            logging.info(f"Tiled PDF saved successfully: {output_pdf_path} (inserted {insertion_count} tiles)")
            
            # Verify the PDF was created
            if os.path.exists(output_pdf_path):
                file_size = os.path.getsize(output_pdf_path)
                logging.info(f"Verified PDF file exists: {output_pdf_path}, size: {file_size} bytes")
                return True
            else:
                logging.error(f"ERROR: PDF file was not created at {output_pdf_path}")
                return False
        except Exception as e:
            logging.error(f"ERROR: Failed to save PDF {output_pdf_path}: {e}")
            return False
            
    except Exception as e:
        logging.error(f"ERROR creating tiled PDF for {image_path}: {e}")
        traceback.print_exc()
        return False

def overlay_footer_and_add_text(base_pdf_path, footer_path, output_pdf_path, template_width, template_height, pattern_name, roll_width, roll_length):
    try:
        logging.info(f"Adding footer and text to {base_pdf_path}")
        
        # Check if base PDF exists
        if not os.path.exists(base_pdf_path):
            logging.error(f"ERROR: Base PDF not found at {base_pdf_path}")
            return False
            
        # Check for footer PDF
        if not os.path.exists(footer_path):
            logging.error(f"ERROR: Footer file not found at {footer_path}")
            
            # Try fallback paths
            fallback_paths = [
                os.path.join(TEMPLATE_IMAGES_FOLDER, "Footer.pdf"),  # In template folder
                os.path.join(SCRIPTS_FOLDER, "Footer.pdf"),          # In scripts folder
                os.path.join(BASE_FOLDER, "Footer.pdf")              # In base folder
            ]
            
            for fallback_path in fallback_paths:
                if os.path.exists(fallback_path):
                    logging.info(f"Using fallback footer path: {fallback_path}")
                    footer_path = fallback_path
                    break
            else:
                # If no footer found, create a simple one
                logging.warning(f"No footer found in any location, creating a simple replacement footer")
                simple_footer_path = os.path.join(os.path.dirname(output_pdf_path), "simple_footer.pdf")
                try:
                    create_simple_footer(simple_footer_path, template_width, 100)
                    footer_path = simple_footer_path
                except Exception as e:
                    logging.error(f"Failed to create simple footer: {e}")
                    return False
        
        try:
            base_doc = fitz.open(base_pdf_path)
            if base_doc.page_count == 0:
                logging.error(f"ERROR: Base PDF has no pages")
                base_doc.close()
                return False
                
            base_page = base_doc[0]
            logging.info(f"Opened base PDF: {base_pdf_path}")

            footer_doc = fitz.open(footer_path)
            if footer_doc.page_count == 0:
                logging.error(f"ERROR: Footer PDF has no pages")
                base_doc.close()
                footer_doc.close()
                return False
                
            footer_page = footer_doc[0]
            logging.info(f"Opened footer PDF: {footer_path}")
        except Exception as e:
            logging.error(f"ERROR: Failed to open PDFs: {e}")
            return False
        
        try:
            footer_pixmap = footer_page.get_pixmap()
            logging.info(f"Got footer pixmap: {footer_pixmap.width}x{footer_pixmap.height}")
        except Exception as e:
            logging.error(f"ERROR: Could not get pixmap from footer: {e}")
            base_doc.close()
            footer_doc.close()
            return False

        # Calculate footer position
        footer_width = template_width
        footer_height = footer_pixmap.height / footer_pixmap.width * footer_width
        footer_rect = fitz.Rect(0, template_height - footer_height, footer_width, template_height)
        logging.info(f"Footer rectangle: {footer_rect}")
        
        try:
            base_page.insert_image(footer_rect, pixmap=footer_pixmap, keep_proportion=True)
            logging.info(f"Inserted footer image")
        except Exception as e:
            logging.error(f"ERROR: Could not insert footer image: {e}")
            base_doc.close()
            footer_doc.close()
            return False

        # Use more appropriate font settings with larger size
        font_name = "Helvetica"
        font_size = 14  # Increased from 10
        dark_grey_color = (0.2, 0.2, 0.2)  # Made slightly darker for better visibility
            
        try:
            # Add pattern name - moved further to the right and using larger font
            base_page.insert_text(
                point=(template_width - 160, template_height - footer_height + 22),  # X position moved from -195 to -160
                text=pattern_name,
                fontsize=font_size,
                fontname=font_name,
                color=dark_grey_color
            )
            logging.info(f"Added pattern name: {pattern_name}")
            
            # Add roll width - moved further to the right
            base_page.insert_text(
                point=(template_width - 135, template_height - footer_height + 37),  # X position moved from -170 to -135
                text=roll_width,
                fontsize=font_size,
                fontname=font_name,
                color=dark_grey_color
            )
            logging.info(f"Added width: {roll_width}")
            
            # Add roll length - moved further to the right
            base_page.insert_text(
                point=(template_width - 135, template_height - footer_height + 55),  # X position moved from -170 to -135
                text=roll_length,
                fontsize=font_size,
                fontname=font_name,
                color=dark_grey_color
            )
            logging.info(f"Added length: {roll_length}")
        except Exception as e:
            logging.warning(f"WARNING: Error adding text: {e}")
            # Try alternative text placement method with similar adjustments
            try:
                logging.info("Trying alternative text placement method")
                # Pattern name at adjusted position
                base_page.insert_text(
                    point=(template_width - 160, template_height - 78),  # Adjusted X position here too
                    text=pattern_name,
                    fontsize=font_size,
                    fontname=font_name,
                    color=dark_grey_color
                )
                
                # Width at adjusted position
                base_page.insert_text(
                    point=(template_width - 135, template_height - 58),  # Adjusted X position
                    text=roll_width,
                    fontsize=font_size,
                    fontname=font_name,
                    color=dark_grey_color
                )
                
                # Length at adjusted position
                base_page.insert_text(
                    point=(template_width - 135, template_height - 38),  # Adjusted X position
                    text=roll_length,
                    fontsize=font_size,
                    fontname=font_name,
                    color=dark_grey_color
                )
                logging.info("Alternative text placement succeeded")
            except Exception as alt_e:
                logging.error(f"Alternative text placement also failed: {alt_e}")
                # Continue despite text error - we'll still have the footer image

        # Save the final PDF
        try:
            temp_output_path = output_pdf_path.replace(".pdf", "_temp.pdf")
            base_doc.save(temp_output_path)
            logging.info(f"Saved temporary PDF: {temp_output_path}")
            
            base_doc.close()
            footer_doc.close()
            
            # Safely replace the original file
            if os.path.exists(output_pdf_path):
                try:
                    os.remove(output_pdf_path)
                    logging.info(f"Removed existing output file: {output_pdf_path}")
                except Exception as e:
                    logging.warning(f"WARNING: Could not remove existing output file: {e}")
                    # Try to use a different filename
                    output_pdf_path = output_pdf_path.replace(".pdf", "_new.pdf")
                    logging.info(f"Using alternative output path: {output_pdf_path}")
            
            os.replace(temp_output_path, output_pdf_path)
            logging.info(f"Final PDF with footer and text saved: {output_pdf_path}")
            
            # Verify file existence
            if os.path.exists(output_pdf_path):
                file_size = os.path.getsize(output_pdf_path)
                logging.info(f"Verified final PDF exists: {output_pdf_path}, size: {file_size} bytes")
                return True
            else:
                logging.error(f"ERROR: Final PDF was not created at {output_pdf_path}")
                return False
        except Exception as e:
            logging.error(f"ERROR: Failed to save final PDF: {e}")
            return False
            
    except Exception as e:
        logging.error(f"ERROR overlaying footer and adding text: {e}")
        traceback.print_exc()
        return False

def create_simple_footer(output_path, width, height):
    """Create a simple footer PDF as a fallback."""
    logging.info(f"Creating simple footer PDF: {output_path}")
    doc = fitz.open()
    page = doc.new_page(width=width, height=height)
    
    # Add a light grey background
    rect = fitz.Rect(0, 0, width, height)
    page.draw_rect(rect, color=(0.95, 0.95, 0.95), fill=(0.95, 0.95, 0.95))
    
    # Add some lines
    page.draw_line(fitz.Point(0, 0), fitz.Point(width, 0), color=(0.8, 0.8, 0.8), width=1)
    
    # Add placeholder text in the same positions as the actual footer
    # Left side text
    page.insert_text(fitz.Point(50, 20), "Celebrating the art of giving with love", fontsize=12)
    page.insert_text(fitz.Point(50, 40), "Aspen & Arlo Team", fontsize=12)
    
    # Center text
    page.insert_text(fitz.Point(width/2 - 100, 30), "ASPEN & ARLO", fontsize=16)
    
    # Right side labels - these are the fields we'll fill in with user data
    page.insert_text(fitz.Point(width - 200, 20), "Pattern:", fontsize=12)
    page.insert_text(fitz.Point(width - 200, 40), "Roll Width:", fontsize=12)
    page.insert_text(fitz.Point(width - 200, 60), "Roll Length:", fontsize=12)
    
    # Website
    page.insert_text(fitz.Point(width - 150, 80), "aspenandarlo.com", fontsize=10)
    
    doc.save(output_path)
    doc.close()
    logging.info(f"Simple footer created: {output_path}")
    return True

def main():
    try:
        parser = argparse.ArgumentParser(description="Generate PDFs with tiled images and overlay footers")
        parser.add_argument("image_paths", type=str, nargs="+", help="List of image paths to process")
        args = parser.parse_args()

        logging.info(f"Starting PDF generation with {len(args.image_paths)} images")
        
        # Verify directories and permissions
        if not verify_directories():
            logging.error("Directory verification failed, exiting")
            return 1

        template_6ft_width = 2171.53
        template_6ft_height = 5285.94
        template_15ft_width = 2171.53
        template_15ft_height = 13061.90

        # Use the new footer path from the scripts folder as specified
        logging.info(f"Using footer path: {FOOTER_PATH}")
        
        # Verify the footer exists or try to find a suitable file
        if not os.path.exists(FOOTER_PATH):
            logging.warning(f"WARNING: Footer file not found at {FOOTER_PATH}")
            # Try to find footer in other locations
            alternate_paths = [
                os.path.join(TEMPLATE_IMAGES_FOLDER, "Footer.pdf"),
                os.path.join(BASE_FOLDER, "Footer.pdf"),
                # Try common variations of the name
                os.path.join(SCRIPTS_FOLDER, "footer.pdf"),
                os.path.join(TEMPLATE_IMAGES_FOLDER, "footer.pdf"),
                os.path.join(BASE_FOLDER, "footer.pdf")
            ]
            
            for alt_path in alternate_paths:
                if os.path.exists(alt_path):
                    logging.info(f"Found alternate footer at: {alt_path}")
                    footer_path = alt_path
                    break
            else:
                logging.warning("No footer found, will create a simple one if needed")
                footer_path = FOOTER_PATH  # Use original path, will create if needed
        else:
            footer_path = FOOTER_PATH

        for image_path in args.image_paths:
            if not os.path.exists(image_path):
                logging.error(f"ERROR: Image file not found: {image_path}")
                continue

            base_name = os.path.splitext(os.path.basename(image_path))[0]
            logging.info(f"Processing {base_name} from {image_path}")

            # Get the current working directory for output
            current_dir = os.getcwd()
            logging.info(f"Current working directory: {current_dir}")

            # Create 6ft PDF
            tiled_pdf_6 = os.path.join(current_dir, f"{base_name}_6ft.pdf")
            logging.info(f"Generating 6ft tiled PDF at: {tiled_pdf_6}")
            
            # Use retry logic for both PDF creation steps
            max_attempts = 3
            success = False
            
            for attempt in range(max_attempts):
                try:
                    logging.info(f"6ft PDF attempt {attempt+1}/{max_attempts}")
                    if create_tiled_image_pdf(tiled_pdf_6, image_path, template_6ft_width, template_6ft_height, horizontal_repeats=6):
                        # Ensure correct foot marker (') is included
                        if overlay_footer_and_add_text(tiled_pdf_6, footer_path, tiled_pdf_6, template_6ft_width, template_6ft_height, base_name, "30'", "6'"):
                            success = True
                            logging.info(f"Successfully created 6ft PDF: {tiled_pdf_6}")
                            break
                except Exception as e:
                    logging.error(f"Error on attempt {attempt+1}/{max_attempts} for 6ft PDF: {e}")
                    traceback.print_exc()
                
                if attempt < max_attempts - 1:
                    logging.info(f"Retrying 6ft PDF generation (attempt {attempt+2}/{max_attempts})...")
                    time.sleep(3)  # Increased wait time before retry
            
            if not success:
                logging.error(f"Failed to generate 6ft PDF for {base_name} after {max_attempts} attempts")

            # Create 15ft PDF
            tiled_pdf_15 = os.path.join(current_dir, f"{base_name}_15ft.pdf")
            logging.info(f"Generating 15ft tiled PDF at: {tiled_pdf_15}")
            
            success = False
            for attempt in range(max_attempts):
                try:
                    logging.info(f"15ft PDF attempt {attempt+1}/{max_attempts}")
                    if create_tiled_image_pdf(tiled_pdf_15, image_path, template_15ft_width, template_15ft_height, horizontal_repeats=6):
                        # Ensure correct foot marker (') is included
                        if overlay_footer_and_add_text(tiled_pdf_15, footer_path, tiled_pdf_15, template_15ft_width, template_15ft_height, base_name, "30'", "15'"):
                            success = True
                            logging.info(f"Successfully created 15ft PDF: {tiled_pdf_15}")
                            break
                except Exception as e:
                    logging.error(f"Error on attempt {attempt+1}/{max_attempts} for 15ft PDF: {e}")
                    traceback.print_exc()
                
                if attempt < max_attempts - 1:
                    logging.info(f"Retrying 15ft PDF generation (attempt {attempt+2}/{max_attempts})...")
                    time.sleep(3)  # Increased wait time before retry
            
            if not success:
                logging.error(f"Failed to generate 15ft PDF for {base_name} after {max_attempts} attempts")

            # Verify files after creation
            if os.path.exists(tiled_pdf_6) and os.path.exists(tiled_pdf_15):
                logging.info(f"PDF generation successful for {base_name}")
            else:
                logging.error(f"PDF generation incomplete for {base_name}")

        logging.info("PDF generation complete.")
        
    except Exception as e:
        logging.error(f"Unexpected error in main function: {e}")
        traceback.print_exc()
        return 1
        
    return 0

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
