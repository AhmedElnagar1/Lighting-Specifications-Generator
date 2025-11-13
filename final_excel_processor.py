from typing import List, Optional, Union
import openpyxl
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import os
import sys
import win32com.client
import re
import shutil
from datetime import datetime
from openpyxl.styles import PatternFill
from PIL import Image as PILImage, ImageOps                
                


def add_page_footer(sheet, page_number: int, total_pages: int):
    """
    Add a footer with creation date on the left and page numbering on the right.
    
    Args:
        sheet: The openpyxl worksheet object
        page_number: Current page number (int)
        total_pages: Total number of pages (int)
    """
    try:
        # Get current date in DD/MM/YYYY format
        current_date = datetime.now().strftime("%d/%m/%Y")
        
        # Set the footer with date on left and page numbering on right
        date_text = current_date
        page_text = f"Page {page_number} of {total_pages}"
        
        sheet.oddFooter.left.text = date_text
        sheet.oddFooter.right.text = page_text
        sheet.evenFooter.left.text = date_text
        sheet.evenFooter.right.text = page_text
        sheet.firstFooter.left.text = date_text
        sheet.firstFooter.right.text = page_text
    except Exception as e:
        print(f"Error adding footer to sheet {sheet.title}: {e}")


def create_backup(excel_file_path: str) -> Optional[str]:
    """
    Create a backup copy of the original file in a Backup folder.
    
    Args:
        excel_file_path (str): Path to the Excel file to backup
    
    Returns:
        Optional[str]: Path to the backup file, or None if backup failed
    """
    try:
        # Get the directory of the Excel file
        file_dir = os.path.dirname(os.path.abspath(excel_file_path))
        
        # Create Backup folder path
        backup_dir = os.path.join(file_dir, "Backup")
        
        # Create Backup folder if it doesn't exist
        os.makedirs(backup_dir, exist_ok=True)
        
        # Get the base filename without extension
        base_name = os.path.splitext(os.path.basename(excel_file_path))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Create backup filename
        backup_filename = f"{base_name}_backup_{timestamp}.xlsx"
        backup_path = os.path.join(backup_dir, backup_filename)
        
        # Copy the file to the backup location
        shutil.copy2(excel_file_path, backup_path)
        print(f"Created backup: {backup_path}")
        return backup_path
    except Exception as e:
        print(f"Warning: Could not create backup: {e}")
        return None


def get_cell_dimensions_emu(sheet, cell_address: str) -> tuple:
    """
    Get cell dimensions in EMU (English Metric Units) for image positioning.
    Handles merged cells by calculating the total merged range dimensions.
    
    Args:
        sheet: The openpyxl worksheet object
        cell_address (str): Cell address (e.g., 'K7', 'AA10')
    
    Returns:
        tuple: (width_emu, height_emu) dimensions in EMU units
    """
    import re
    from openpyxl.utils import get_column_letter
    
    # Parse cell address using regex: column letters followed by row number
    match = re.match(r'^([A-Z]+)(\d+)$', cell_address.upper())
    if not match:
        raise ValueError(f"Invalid cell address: {cell_address}")
    
    col_letter = match.group(1)
    row_num = int(match.group(2))
    
    # Check if the cell is part of a merged range
    merged_range = None
    for merged_cell in sheet.merged_cells.ranges:
        if cell_address in merged_cell:
            merged_range = merged_cell
            break
    
    if merged_range:
        # Calculate dimensions of merged range
        # Get start and end columns and rows
        min_col = merged_range.min_col
        max_col = merged_range.max_col
        min_row = merged_range.min_row
        max_row = merged_range.max_row
        
        # Calculate total width (sum of all columns in merged range)
        total_width = 0
        for col_idx in range(min_col, max_col + 1):
            col_letter_merged = get_column_letter(col_idx)
            column_width = sheet.column_dimensions[col_letter_merged].width
            if column_width is None:
                column_width = 8.43  # Default Excel column width
            total_width += column_width
        
        # Calculate total height (sum of all rows in merged range)
        total_height = 0
        for row_idx in range(min_row, max_row + 1):
            row_height = sheet.row_dimensions[row_idx].height
            if row_height is None:
                row_height = 15  # Default Excel row height
            total_height += row_height
    else:
        # Single cell - use original logic
        column_width = sheet.column_dimensions[col_letter].width
        if column_width is None:
            column_width = 8.43  # Default Excel column width
        total_width = column_width
        
        row_height = sheet.row_dimensions[row_num].height
        if row_height is None:
            row_height = 15  # Default Excel row height
        total_height = row_height
    
    # Convert to pixels and then to EMU
    # Column width: 1 character ≈ 7 pixels at 96 DPI
    # Row height: 1 point = 4/3 pixels at 96 DPI
    # 1 pixel = 9525 EMU
    width_pixels = total_width * 7
    width_emu = int(width_pixels * 9525)
    
    height_pixels = total_height * (4 / 3)
    height_emu = int(height_pixels * 9525)
    
    return (width_emu, height_emu)


def fix_image_orientation(image_path: str) -> str:
    """
    Fix image orientation based on EXIF data and return path to corrected image.
    Uses PIL's ImageOps.exif_transpose() to automatically handle all EXIF orientations.
    
    Args:
        image_path (str): Path to the original image file
    
    Returns:
        str: Path to the corrected image (may be original if no rotation needed)
    """
    try:
        # Open image with PIL
        pil_img = PILImage.open(image_path)
        
        # Check if image has EXIF orientation data that needs correction
        needs_correction = False
        try:
            exif = pil_img.getexif()
            if exif:
                orientation = exif.get(274)  # EXIF orientation tag
                # Orientation values: 1=normal, 2-8 need correction
                if orientation and orientation != 1:
                    needs_correction = True
        except (AttributeError, KeyError, TypeError):
            # No EXIF data or can't read it
            pass
        
        if needs_correction:
            # Apply EXIF orientation correction automatically
            # This handles all EXIF orientation tags (1-8)
            pil_img_corrected = ImageOps.exif_transpose(pil_img)
            
            # Save to temporary file with same extension as original
            file_ext = os.path.splitext(image_path)[1].lower()
            base_name = os.path.splitext(image_path)[0]
            temp_path = base_name + "_temp_orient" + file_ext
            
            # Convert to RGB if necessary (for JPEG)
            if pil_img_corrected.mode in ("RGBA", "P", "LA"):
                pil_img_corrected = pil_img_corrected.convert("RGB")
            # Preserve original format if possible
            if file_ext in ['.jpg', '.jpeg']:
                pil_img_corrected.save(temp_path, quality=95, format='JPEG')
            elif file_ext == '.png':
                pil_img_corrected.save(temp_path, format='PNG')
            else:
                pil_img_corrected.save(temp_path, quality=95)
            
            pil_img.close()
            pil_img_corrected.close()
            return temp_path
        
        pil_img.close()
        return image_path
        
    except Exception as e:
        print(f"Warning: Could not fix image orientation for {image_path}: {e}")
        return image_path


def add_image_to_cell(sheet, img: Image, cell_address: str) -> None:
    """
    Add an image to a cell, resizing it to fit inside the cell.
    Handles merged cells and allows scaling up to fill the cell better.
    
    Args:
        sheet: The openpyxl worksheet object
        img: The Image object to add
        cell_address (str): Cell address where to place the image (e.g., 'K7')
    """
    # Get cell dimensions in EMU (handles merged cells)
    cell_width_emu, cell_height_emu = get_cell_dimensions_emu(sheet, cell_address)
    
    # Use 95% of cell dimensions to leave a small margin
    usable_width_emu = int(cell_width_emu * 0.95)
    usable_height_emu = int(cell_height_emu * 0.95)
    
    # Get current image dimensions in EMU (openpyxl uses EMU internally)
    # Image dimensions are already in pixels, convert to EMU
    # 1 pixel = 9525 EMU
    img_width_emu = int(img.width * 9525)
    img_height_emu = int(img.height * 9525)
    
    # Calculate scaling factor to fit image inside cell
    # Allow scaling up if image is smaller than cell
    width_ratio = usable_width_emu / img_width_emu
    height_ratio = usable_height_emu / img_height_emu
    scale_factor = min(width_ratio, height_ratio)  # Use smaller ratio to maintain aspect ratio
    
    # Resize image to fit within cell (can scale up or down)
    img.width = int(img.width * scale_factor)
    img.height = int(img.height * scale_factor)
    print(f"Resized image for cell {cell_address}: {img.width}x{img.height} pixels (scale factor: {scale_factor:.2f})")
    
    # Set anchor to cell (this positions the image at the top-left of the cell)
    img.anchor = cell_address
    
    # Add image to sheet
    sheet.add_image(img)


def add_image_to_sheet(sheet, sheet_id: str, img_dir: str) -> tuple:
    """
    Add images to sheet based on sheet ID.
    
    Args:
        sheet: The openpyxl worksheet object
        sheet_id (str): The sheet ID to match with image folder
        img_dir (str): Path to the image directory
    
    Returns:
        tuple: (success: bool, temp_files: List[str]) - Success status and list of temporary files to clean up later
    """
    temp_files = []  # Track temporary files for cleanup
    try:
        images_path = os.path.join(img_dir, sheet_id)
        if not os.path.exists(images_path):
            print(f"Warning: Images not found for {sheet_id}: {images_path}")
            return (False, temp_files)
        images = os.listdir(images_path)
        site_images = []
        plan_images = []
        for image in images:
            if "site" in image.lower():
                site_images.append(image)
            elif "plan" in image.lower():
                plan_images.append(image)
        
        # Add site images
        for i, site_image in enumerate(site_images):
            image_path = os.path.join(images_path, site_image)
            try:
                # Fix orientation before loading
                corrected_path = fix_image_orientation(image_path)
                if corrected_path != image_path:
                    temp_files.append(corrected_path)
                
                if not os.path.exists(corrected_path):
                    print(f"Error: Image file not found: {corrected_path}")
                    continue
                
                img = Image(corrected_path)
                
                # Determine cell address based on image index
                if i == 0:
                    cell_address = 'K7'
                else:
                    cell_address = 'O7'
                
                # Add image to cell (will be resized to fit)
                add_image_to_cell(sheet, img, cell_address)
                print(f"Added site image {site_image} to cell {cell_address}")
            except Exception as e:
                print(f"Error adding site image {site_image}: {e}")
        
        # Add plan images
        for i, plan_image in enumerate(plan_images):
            image_path = os.path.join(images_path, plan_image)
            try:
                # Fix orientation before loading
                corrected_path = fix_image_orientation(image_path)
                if corrected_path != image_path:
                    temp_files.append(corrected_path)
                
                if not os.path.exists(corrected_path):
                    print(f"Error: Image file not found: {corrected_path}")
                    continue
                
                img = Image(corrected_path)
                
                # Determine cell address based on image index
                if i == 0:
                    cell_address = 'O16'
                else:
                    # If there's a second plan image, place it in a different cell
                    # Adjust this cell address as needed
                    cell_address = 'K16'
                
                # Add image to cell (will be resized to fit)
                add_image_to_cell(sheet, img, cell_address)
                print(f"Added plan image {plan_image} to cell {cell_address}")
            except Exception as e:
                print(f"Error adding plan image {plan_image}: {e}")
        
        print(f"Added images for {sheet_id}")
        return (True, temp_files)
        
    except Exception as e:
        print(f"Error adding image for {sheet_id}: {e}")
        import traceback
        traceback.print_exc()
        return (False, temp_files)

def find_sheets_containing(workbook, search_term: str) -> List[str]:
    """
    Find all sheet names containing the search term (case-insensitive).
    
    Args:
        workbook: The openpyxl workbook object
        search_term (str): The term to search for in sheet names
    
    Returns:
        List[str]: List of sheet names containing the search term
    """
    matching_sheets = []
    for sheet_name in workbook.sheetnames:
        if search_term.lower() in sheet_name.lower():
            matching_sheets.append(sheet_name)
    return matching_sheets


def find_decision_matrix_sheet(workbook) -> Optional[str]:
    """
    Find the sheet that contains a cell with the exact content "Decision Matrix" (case-insensitive).
    
    Args:
        workbook: The openpyxl workbook object
    
    Returns:
        Optional[str]: The name of the sheet containing "Decision Matrix" in a cell, or None if not found
    """
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        # Search through all cells in the sheet
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    cell_value = str(cell.value).strip()
                    if cell_value.lower() == "decision matrix":
                        return sheet_name
    return None


def find_template_sheets(workbook) -> List[str]:
    """
    Find all sheets containing "Template" in their name.
    
    Args:
        workbook: The openpyxl workbook object
    
    Returns:
        List[str]: List of sheet names containing "Template"
    """
    return find_sheets_containing(workbook, "Template")


def create_sheets(input_wb, img_dir, input_file, template_sheet_name: str, schedule_sheet_name: str):
    """
    Create sheets directly from Schedule sheet data using openpyxl.
    
    Args:
        input_wb: The openpyxl workbook object
        img_dir (str): Path to the image directory
        input_file (str): Path to the input Excel file
        template_sheet_name (str): Name of the template sheet to use
        schedule_sheet_name (str): Name of the Decision Matrix sheet to use
    
    Returns:
        List[str] or False: List of created sheet IDs, or False on error
    """
    if template_sheet_name not in input_wb.sheetnames:
        print(f"Template sheet '{template_sheet_name}' not found")
        return False
    
    if schedule_sheet_name not in input_wb.sheetnames:
        print(f"Schedule sheet '{schedule_sheet_name}' not found")
        return False
    
    try:
        # Get the template sheet
        template_sheet = input_wb[template_sheet_name]
        
        # Get the Schedule sheet
        schedule_sheet = input_wb[schedule_sheet_name]
        sheets_created = 0
        
        # Get column names from row 10
        column_names = {}
        for col_num in range(1, schedule_sheet.max_column + 1):
            cell_value = schedule_sheet.cell(row=10, column=col_num).value
            if cell_value and (cell_value != 'Eingebunden in Steuerung'):
                column_names[col_num] = str(cell_value).strip()
        
        print(f"Found columns: {list(column_names.values())}")
        
        # Process each row in the Schedule sheet starting from row 11
        base_rows = []
        option_1_rows = []
        option_2_rows = []
        sheet_ids = []
        for row_num in range(11, schedule_sheet.max_row + 1):
            # Get all values for this row
            row_data = {}
            for col_num in column_names.keys():
                cell_value = schedule_sheet.cell(row=row_num, column=col_num).value
                row_data[column_names[col_num]] = cell_value
            
            cell_value = row_data.get("Report Code")

            if cell_value is None or str(cell_value).strip() == '':
                # Skip row if Report Code is blank
                continue
            cell_value = cell_value.strip()

            if cell_value.endswith("X"):
                base_rows.append(row_data)
            elif cell_value.endswith("1"):
                option_1_rows.append(row_data)
            elif cell_value.endswith("2"):
                option_2_rows.append(row_data)
            
        for base_row, option_1_row, option_2_row in zip(base_rows, option_1_rows, option_2_rows):
            value_str = str(base_row.get("Report Code").strip())
            # Clean invalid characters
            base_sheet_id = re.sub(r'[\[\]*?/\\:;]', '', value_str).strip()
            
            # Check if sheet already exists and add suffix if needed
            sheet_id = base_sheet_id
            suffix = 1
            while sheet_id in input_wb.sheetnames:
                # Create unique sheet name with suffix
                suffix_sheet_id = f"{base_sheet_id}_{suffix}"
                if len(suffix_sheet_id) <= 31:
                    sheet_id = suffix_sheet_id
                else:
                    # If suffix makes name too long, truncate base name
                    max_base_length = 31 - len(f"_{suffix}") - 1
                    truncated_base = base_sheet_id[:max_base_length]
                    sheet_id = f"{truncated_base}_{suffix}"
                suffix += 1
            
            # Log if suffix was added
            if sheet_id != base_sheet_id:
                print(f"Sheet '{base_sheet_id}' already exists, using '{sheet_id}' instead")
            
            sheet_ids.append(sheet_id)
            if sheet_id and len(sheet_id) <= 31 and sheet_id not in input_wb.sheetnames:
                print(f"Creating sheet: {sheet_id}")
                print(f"Row data: {base_row}")
                
                # Copy the template sheet
                new_sheet = input_wb.copy_worksheet(template_sheet)
                new_sheet.title = sheet_id
                
                # Map data from Decision Matrix to template fields
                map_base_data_to_template(new_sheet, base_row, option_1_row, option_2_row)
                
                # Add image to the sheet
                success, temp_files = add_image_to_sheet(new_sheet, sheet_id, img_dir)
                if temp_files:
                    # Store temp files for cleanup after save
                    if not hasattr(input_wb, '_temp_image_files'):
                        input_wb._temp_image_files = []
                    input_wb._temp_image_files.extend(temp_files)
                
                sheets_created += 1
            elif sheet_id in input_wb.sheetnames:
                print(f"Sheet {sheet_id} already exists, deleting and recreating...")
                # Remove the existing sheet
                input_wb.remove(input_wb[sheet_id])
                print(f"Deleted existing sheet: {sheet_id}")
                
                # Create the new sheet
                print(f"Creating sheet: {sheet_id}")
                print(f"Row data: {base_row}")
                
                # Copy the template sheet
                new_sheet = input_wb.copy_worksheet(template_sheet)
                new_sheet.title = sheet_id
                
                # Map data from Decision Matrix to template fields
                map_base_data_to_template(new_sheet, base_row, option_1_row, option_2_row)
                
                # Add image to the sheet
                success, temp_files = add_image_to_sheet(new_sheet, sheet_id, img_dir)
                if temp_files:
                    # Store temp files for cleanup after save
                    if not hasattr(input_wb, '_temp_image_files'):
                        input_wb._temp_image_files = []
                    input_wb._temp_image_files.extend(temp_files)
                
                sheets_created += 1
        
        # Add footers only to created sheets
        # Calculate total pages: Cover + GenInfo+Contacts + created sheets
        total_pages = len(sheet_ids)  # Cover and GenInfo+Contacts are always included
        
        # Add footers to created sheets only (starting from page 3)
        for i, sheet_id in enumerate(sheet_ids):
            if sheet_id in input_wb.sheetnames:
                page_number = 1 + i
                if page_number == 36:
                    pass
                add_page_footer(input_wb[sheet_id], page_number, total_pages)
        
        # Save the workbook
        input_wb.save(input_file)
        print(f"Created {sheets_created} new sheets")
        print("Workbook saved successfully")
        
        # Clean up temporary image files after save
        if hasattr(input_wb, '_temp_image_files'):
            for temp_file in input_wb._temp_image_files:
                try:
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                        print(f"Cleaned up temporary file: {temp_file}")
                except Exception as e:
                    print(f"Warning: Could not delete temporary file {temp_file}: {e}")
            delattr(input_wb, '_temp_image_files')
        
        return sheet_ids
        
    except Exception as e:
        print(f"Error creating sheets: {e}")
        print("Trying to save with different name...")
        
        # Try saving with a different name
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        new_path = f"{os.path.splitext(input_file)[0]}_modified_{timestamp}.xlsx"
        
        try:
            input_wb.save(new_path)
            print(f"Workbook saved as: {new_path}")
            return new_path
        except Exception as e2:
            print(f"Error saving to new path: {e2}")
            return False

def map_base_data_to_template(sheet, base_row_data, option_1_row_data, option_2_row_data):
    """Map data from Decision Matrix to template fields"""
    try:
        # Map specific fields to exact cell locations as requested
        
        # System Description in A3
        if "System Description" in base_row_data and base_row_data["System Description"]:
            sheet.cell(row=3, column=1).value = base_row_data["System Description"]
        
        # Manufacturer/Type in C3
        if "Manufacturer / Type" in base_row_data and base_row_data["Manufacturer / Type"]:
            sheet.cell(row=3, column=3).value = base_row_data["Manufacturer / Type"]
        
        # Assessed Condition in D3
        if "Assessed Condition" in base_row_data and base_row_data["Assessed Condition"]:
            sheet.cell(row=3, column=4).value = base_row_data["Assessed Condition"]
            if "1" in base_row_data["Assessed Condition"]:
                # Fill cell with RGB(0,230,104) - Green
                fill = PatternFill(start_color="00E668", end_color="00E668", fill_type="solid")
                sheet.cell(row=3, column=4).fill = fill
            elif "2" in base_row_data["Assessed Condition"]:
                # Fill cell with RGB(186,225,143) - Light Green
                fill = PatternFill(start_color="BAE18F", end_color="BAE18F", fill_type="solid")
                sheet.cell(row=3, column=4).fill = fill
            elif "3" in base_row_data["Assessed Condition"]:
                # Fill cell with RGB(247,199,172) - Light Orange
                fill = PatternFill(start_color="F7C7AC", end_color="F7C7AC", fill_type="solid")
                sheet.cell(row=3, column=4).fill = fill
            elif "4" in base_row_data["Assessed Condition"]:
                # Fill cell with RGB(241,169,131) - Orange
                fill = PatternFill(start_color="F1A983", end_color="F1A983", fill_type="solid")
                sheet.cell(row=3, column=4).fill = fill
            elif "5" in base_row_data["Assessed Condition"]:
                # Fill cell with RGB(255,113,113) - Red
                fill = PatternFill(start_color="FF7171", end_color="FF7171", fill_type="solid")
                sheet.cell(row=3, column=4).fill = fill

        
        # Current Light Technology in F3
        if "Current Light Technology" in base_row_data and base_row_data["Current Light Technology"]:
            sheet.cell(row=3, column=6).value = base_row_data["Current Light Technology"]
        
        # Lamp Fitting in G3
        if "Lamp Fitting" in base_row_data and base_row_data["Lamp Fitting"]:
            sheet.cell(row=3, column=7).value = base_row_data["Lamp Fitting"]
        
        # Socket in H3
        if "Socket" in base_row_data and base_row_data["Socket"]:
            sheet.cell(row=3, column=8).value = base_row_data["Socket"]
        
        # Quantity in Space in I3
        if "System Quantity in Space" in base_row_data and base_row_data["System Quantity in Space"]:
            sheet.cell(row=3, column=9).value = base_row_data["System Quantity in Space"]
        
        # Level in K3
        if "Level" in base_row_data and base_row_data["Level"]:
            sheet.cell(row=3, column=11).value = base_row_data["Level"]
        
        # Room No. in L3
        if "Room No." in base_row_data and base_row_data["Room No."]:
            sheet.cell(row=3, column=12).value = base_row_data["Room No."]
        
        # Room Name in M3
        if "Room Name" in base_row_data and base_row_data["Room Name"]:
            sheet.cell(row=3, column=13).value = base_row_data["Room Name"]
        
        # Luminaire Type in P3
        if "Luminaire Type" in base_row_data and base_row_data["Luminaire Type"]:
            sheet.cell(row=3, column=16).value = base_row_data["Luminaire Type"]
        
        # Notes in K16
        if "Notes" in base_row_data and base_row_data["Notes"]:
            sheet.cell(row=16, column=11).value = base_row_data["Notes"]
        
        # Option 1 row data mapping (Column C)
        if option_1_row_data:
            # Technical Specification in C8
            if "Technical Specification" in option_1_row_data and option_1_row_data["Technical Specification"]:
                sheet.cell(row=8, column=3).value = option_1_row_data["Technical Specification"]
            
            # Recommended Additional Services in C9
            if "Recommended Additional services" in option_1_row_data and option_1_row_data["Recommended Additional services"]:
                sheet.cell(row=9, column=3).value = option_1_row_data["Recommended Additional services"]
            
            # Sustainability Considerations in C11
            if "Sustainability Considerations" in option_1_row_data and option_1_row_data["Sustainability Considerations"]:
                sheet.cell(row=11, column=3).value = option_1_row_data["Sustainability Considerations"]
            
            # Anticipated Risks in C12
            if "Anticipated Risks" in option_1_row_data and option_1_row_data["Anticipated Risks"]:
                sheet.cell(row=12, column=3).value = option_1_row_data["Anticipated Risks"]
            
            # System Power Consumption in C16
            if "System Power Consumption" in option_1_row_data and option_1_row_data["System Power Consumption"]:
                power_value = option_1_row_data["System Power Consumption"]
                # Format as "XX W" if it's a number, otherwise keep as is
                if isinstance(power_value, (int, float)) and power_value != 0:
                    sheet.cell(row=16, column=3).value = f"{power_value} W"
                else:
                    sheet.cell(row=16, column=3).value = power_value
            
            # Estimated energy saving in C17
            if "Estimated Energy Saving" in option_1_row_data and option_1_row_data["Estimated Energy Saving"]:
                sheet.cell(row=17, column=3).value = option_1_row_data["Estimated Energy Saving"]
            
            # System Rated Life Time in C18
            if "System Rated Life Time" in option_1_row_data and option_1_row_data["System Rated Life Time"]:
                sheet.cell(row=18, column=3).value = option_1_row_data["System Rated Life Time"]
            
            # Expected Maintenance Cycle in C19
            if "Expected Maintenance Cycle*" in option_1_row_data and option_1_row_data["Expected Maintenance Cycle*"]:
                sheet.cell(row=19, column=3).value = option_1_row_data["Expected Maintenance Cycle*"]
            
            # Warranty in C20
            if "Warranty" in option_1_row_data and option_1_row_data["Warranty"]:
                sheet.cell(row=20, column=3).value = option_1_row_data["Warranty"]
            
            # Delivery Time in C21
            if "Delivery Time" in option_1_row_data and option_1_row_data["Delivery Time"]:
                sheet.cell(row=21, column=3).value = option_1_row_data["Delivery Time"]
            
            # Cost per System in C22
            if "Cost per System" in option_1_row_data and option_1_row_data["Cost per System"]:
                sheet.cell(row=22, column=3).value = option_1_row_data["Cost per System"]
            
            # Estimated Total Systems Cost incl. Mounting per Room in C23
            if "Estimated Total Systems Cost incl. \nMounting per Room" in option_1_row_data and option_1_row_data["Estimated Total Systems Cost incl. \nMounting per Room"]:
                sheet.cell(row=23, column=3).value = option_1_row_data["Estimated Total Systems Cost incl. \nMounting per Room"]

            # Product Link in C24
            if "System link" in option_1_row_data and option_1_row_data["System link"]:
                sheet.cell(row=24, column=3).value = option_1_row_data["System link"]

        # Option 2 row data mapping (Column F)
        if option_2_row_data:
            # Technical Specification in F8
            if "Technical Specification" in option_2_row_data and option_2_row_data["Technical Specification"]:
                sheet.cell(row=8, column=6).value = option_2_row_data["Technical Specification"]
            
            # Recommended Additional Services in F9
            if "Recommended Additional services" in option_2_row_data and option_2_row_data["Recommended Additional services"]:
                sheet.cell(row=9, column=6).value = option_2_row_data["Recommended Additional services"]
            
            # Sustainability Considerations in F11
            if "Sustainability Considerations" in option_2_row_data and option_2_row_data["Sustainability Considerations"]:
                sheet.cell(row=11, column=6).value = option_2_row_data["Sustainability Considerations"]
            
            # Anticipated Risks in F12
            if "Anticipated Risks" in option_2_row_data and option_2_row_data["Anticipated Risks"]:
                sheet.cell(row=12, column=6).value = option_2_row_data["Anticipated Risks"]
            
            # System Power Consumption in F16
            if "System Power Consumption" in option_2_row_data and option_2_row_data["System Power Consumption"]:
                power_value = option_2_row_data["System Power Consumption"]
                # Format as "XX W" if it's a number, otherwise keep as is
                if isinstance(power_value, (int, float)) and power_value != 0:
                    sheet.cell(row=16, column=6).value = f"{power_value} W"
                else:
                    sheet.cell(row=16, column=6).value = power_value
            
            # Estimated energy saving in F17
            if "Estimated Energy Saving" in option_2_row_data and option_2_row_data["Estimated Energy Saving"]:
                sheet.cell(row=17, column=6).value = option_2_row_data["Estimated Energy Saving"]
            
            # System Rated Life Time in F18
            if "System Rated Life Time" in option_2_row_data and option_2_row_data["System Rated Life Time"]:
                sheet.cell(row=18, column=6).value = option_2_row_data["System Rated Life Time"]
            
            # Expected Maintenance Cycle in F19
            if "Expected Maintenance Cycle*" in option_2_row_data and option_2_row_data["Expected Maintenance Cycle*"]:
                sheet.cell(row=19, column=6).value = option_2_row_data["Expected Maintenance Cycle*"]
            
            # Warranty in F20
            if "Warranty" in option_2_row_data and option_2_row_data["Warranty"]:
                sheet.cell(row=20, column=6).value = option_2_row_data["Warranty"]
            
            # Delivery Time in F21
            if "Delivery Time" in option_2_row_data and option_2_row_data["Delivery Time"]:
                sheet.cell(row=21, column=6).value = option_2_row_data["Delivery Time"]
            
            # Cost per System in F22
            if "Cost per System" in option_2_row_data and option_2_row_data["Cost per System"]:
                sheet.cell(row=22, column=6).value = option_2_row_data["Cost per System"]
            
            # Estimated Total Systems Cost incl. Mounting per Room in F23
            if "Estimated Total Systems Cost incl. \nMounting per Room" in option_2_row_data and option_2_row_data["Estimated Total Systems Cost incl. \nMounting per Room"]:
                sheet.cell(row=23, column=6).value = option_2_row_data["Estimated Total Systems Cost incl. \nMounting per Room"]

            # Product Link in F24
            if "System link" in option_2_row_data and option_2_row_data["System link"]:
                sheet.cell(row=24, column=6).value = option_2_row_data["System link"]

        print(f"Mapped data for sheet: {sheet.title}")
        
    except Exception as e:
        print(f"Error mapping data to template: {e}")

def create_pdf(excel_file_path: str, sheet_ids: List[str]) -> Union[str, bool]:
    """
    Create PDF from all sheets using Excel COM automation.
    
    Args:
        excel_file_path (str): Path to the Excel file
        sheet_ids (List[str]): List of sheet IDs to include in PDF
    
    Returns:
        Union[str, bool]: Path to the generated PDF file on success, or False on error
    """
    output_pdf = os.path.splitext(excel_file_path)[0] + "_output.pdf"
    if os.path.exists(output_pdf):
        try:
            os.remove(output_pdf)
        except Exception as e:
            print(f"Warning: Could not remove existing PDF: {e}")
    
    print(f"Creating PDF: {output_pdf}")
    
    excel_app = None
    workbook = None
    
    try:
        # Use Excel automation
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        excel_app.ScreenUpdating = False
        
        # Get absolute path and ensure it exists
        abs_file_path = os.path.abspath(excel_file_path)
        if not os.path.exists(abs_file_path):
            raise FileNotFoundError(f"Excel file not found: {abs_file_path}")
        
        print(f"Opening workbook: {abs_file_path}")
        # Open workbook with ReadOnly=False to allow modifications
        workbook = excel_app.Workbooks.Open(
            abs_file_path,
            ReadOnly=False,
            UpdateLinks=0,
            CorruptLoad=0
        )
        
        if workbook is None:
            raise Exception("Failed to open workbook")
        
        # Get list of actual sheet names in workbook
        actual_sheet_names = []
        for sheet in workbook.Sheets:
            actual_sheet_names.append(sheet.Name)
        
        print(f"Actual sheets in workbook: {actual_sheet_names}")
        
        # Define sheets to include in PDF export
        sheets_to_include = ["Cover", "GenInfo+Contacts"]
        sheets_to_include = sheets_to_include + sheet_ids
        
        # Filter to only include sheets that actually exist
        existing_sheets_to_include = [s for s in sheets_to_include if s in actual_sheet_names]
        missing_sheets = [s for s in sheets_to_include if s not in actual_sheet_names]
        
        if missing_sheets:
            print(f"Warning: Some sheets not found in workbook: {missing_sheets}")
        
        if not existing_sheets_to_include:
            raise Exception("No valid sheets found to include in PDF")
        
        print(f"Sheets to include in PDF: {existing_sheets_to_include}")
        
        # Hide sheets that should not be included in PDF
        # Excel constants: xlSheetHidden = -1, xlSheetVisible = -1 (but True/False also works)
        for sheet in workbook.Sheets:
            try:
                if sheet.Name not in existing_sheets_to_include:
                    sheet.Visible = False  # False = hidden
                    print(f"Hidden sheet: {sheet.Name}")
            except Exception as e:
                print(f"Warning: Could not hide sheet {sheet.Name}: {e}")
        
        # Ensure at least one sheet is visible
        visible_count = 0
        for sheet in workbook.Sheets:
            try:
                if sheet.Visible != False:  # Check if visible (not False)
                    visible_count += 1
            except:
                pass
        
        if visible_count == 0:
            raise Exception("No visible sheets to export to PDF")
        
        print(f"Exporting {visible_count} visible sheet(s) to PDF...")
        
        # Get absolute path for output PDF
        abs_output_pdf = os.path.abspath(output_pdf)
        output_dir = os.path.dirname(abs_output_pdf)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Export to PDF (only visible sheets will be included)
        workbook.ExportAsFixedFormat(
            Type=0,  # xlTypePDF = 0
            Filename=abs_output_pdf,
            Quality=0,  # xlQualityStandard = 0
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False
        )
        
        # Make all sheets visible again
        for sheet in workbook.Sheets:
            try:
                sheet.Visible = True  # True = visible
            except Exception as e:
                print(f"Warning: Could not make sheet {sheet.Name} visible: {e}")
        
        # Clean up
        workbook.Close(SaveChanges=False)
        workbook = None
        excel_app.Quit()
        excel_app = None
        
        # Verify PDF was created
        if os.path.exists(abs_output_pdf):
            print(f"PDF created successfully: {abs_output_pdf}")
            return abs_output_pdf
        else:
            raise Exception(f"PDF file was not created: {abs_output_pdf}")
        
    except Exception as e:
        error_msg = f"Error creating PDF: {e}"
        print(error_msg)
        import traceback
        traceback.print_exc()
        
        # Clean up in case of error
        try:
            if workbook is not None:
                try:
                    # Make all sheets visible before closing
                    for sheet in workbook.Sheets:
                        try:
                            sheet.Visible = True
                        except:
                            pass
                    workbook.Close(SaveChanges=False)
                except:
                    pass
        except:
            pass
        
        try:
            if excel_app is not None:
                excel_app.Quit()
        except:
            pass
        
        return False

def process_excel_file(input_file: str, img_dir: str, template_sheet_name: Optional[str] = None) -> Union[str, bool]:
    """
    Main processing function.
    
    Args:
        input_file (str): Path to the input Excel file
        img_dir (str): Path to the image directory
        template_sheet_name (Optional[str]): Name of the template sheet to use. If None, will be auto-detected.
    
    Returns:
        Union[str, bool]: Path to the generated PDF file on success, or False on error
    """
    print("Starting Excel processing and PDF creation...")
    print("=" * 50)
    
    # Create backup
    backup_path = create_backup(input_file)
    
    # Load workbook
    try:
        input_wb = load_workbook(input_file, data_only=True)  # Read only values, not formulas
        print(f"Loaded workbook: {input_file}")
    except Exception as e:
        print(f"Error loading workbook: {e}")
        return False

    if input_wb is None:
        return False
    
    # Find Decision Matrix sheet (sheet containing a cell with "Decision Matrix")
    decision_matrix_sheet = find_decision_matrix_sheet(input_wb)
    if not decision_matrix_sheet:
        print("Error: No sheet containing a cell with 'Decision Matrix' found")
        return False
    print(f"Found Decision Matrix sheet: {decision_matrix_sheet}")
    
    # Find or use provided template sheet
    if template_sheet_name is None:
        template_sheets = find_template_sheets(input_wb)
        if not template_sheets:
            print("Error: No sheet containing 'Template' found")
            return False
        elif len(template_sheets) == 1:
            template_sheet_name = template_sheets[0]
            print(f"Found Template sheet: {template_sheet_name}")
        else:
            # Multiple template sheets found - this should be handled by the GUI
            print(f"Error: Multiple template sheets found: {template_sheets}")
            print("Please select a template sheet in the GUI")
            return False
    else:
        if template_sheet_name not in input_wb.sheetnames:
            print(f"Error: Specified template sheet '{template_sheet_name}' not found")
            return False
        print(f"Using Template sheet: {template_sheet_name}")
    
    # Determine sheets to keep
    sheets_to_keep = ["Cover", "GenInfo+Contacts", template_sheet_name, decision_matrix_sheet]
    for sheet in input_wb.sheetnames:
        if sheet not in sheets_to_keep:
            input_wb.remove(input_wb[sheet])
    
    # Create sheets
    sheet_ids = create_sheets(input_wb, img_dir, input_file, template_sheet_name, decision_matrix_sheet)
    if not sheet_ids:
        return False
    
    # Create PDF
    pdf_path = create_pdf(input_file, sheet_ids)
    if not pdf_path:
        return False
    
    print("=" * 50)
    print("Processing completed successfully!")
    print(f"Modified Excel file: {input_file}")
    print(f"PDF output: {os.path.splitext(input_file)[0]}_output.pdf")
    if backup_path:
        print(f"Backup created: {backup_path}")
    return pdf_path


if __name__ == "__main__":
    input_file = r"C:\Users\aelnagar\Documents\GitHub\Lighting-Specifications-Generator\0062646_Microsoft Berlin_UdL_Decision Matrix_v.02.xlsx"
    img_dir = r"C:\Users\aelnagar\Buro Happold\0062646 Microsoft Berlin - 05_Presentations\Specifications"
    #excel_file = r"C:\Users\aelnagar\Downloads\Lighting Computational Development\Bauphase.xlsm"
    process_excel_file(input_file, img_dir) 