from typing import List
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


def create_backup(excel_file_path):
    """Create a backup copy of the original file"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{os.path.splitext(excel_file_path)[0]}_backup_{timestamp}.xlsx"
    
    try:
        shutil.copy2(excel_file_path, backup_path)
        print(f"Created backup: {backup_path}")
        return backup_path
    except Exception as e:
        print(f"Warning: Could not create backup: {e}")
        return None


def add_image_to_sheet(sheet, sheet_id, img_dir):
    """Add image to sheet based on sheet ID"""
    try:
        images_path = os.path.join(img_dir, sheet_id)
        if not os.path.exists(images_path):
            print(f"Warning: Images not found for {sheet_id}: {images_path}")
            return False
        images = os.listdir(images_path)
        site_images = []
        plan_images = []
        for image in images:
            if "site" in image.lower():
                site_images.append(image)
            elif "plan" in image.lower():
                plan_images.append(image)
        
        for i, site_image in enumerate(site_images):
            image_path = os.path.join(images_path, site_image)
            img = Image(image_path)
            # Resize image while maintaining aspect ratio
            # Set maximum dimensions (adjust these values as needed)
            max_width = 300
            max_height = 400
            
            # Calculate new dimensions while maintaining aspect ratio
            original_width = img.width
            original_height = img.height
            # Calculate scaling factors
            width_ratio = max_width / original_width
            height_ratio = max_height / original_height
            
            # Use the smaller ratio to ensure image fits within bounds
            scale_factor = min(width_ratio, height_ratio)
            
            # Only resize if the image is larger than max dimensions
            if scale_factor < 1:
                img.width = int(original_width * scale_factor)
                img.height = int(original_height * scale_factor)
                print(f"Resized image from {original_width}x{original_height} to {img.width}x{img.height}")
            else:
                print(f"Image size {original_width}x{original_height} is within limits, no resizing needed")
            if i == 0:
                sheet.add_image(img, 'K7')
            else:
                sheet.add_image(img, 'O7')
        
        for i, plan_image in enumerate(plan_images):
            image_path = os.path.join(images_path, plan_image)
            img = Image(image_path)
            # Resize image while maintaining aspect ratio
            # Set maximum dimensions (adjust these values as needed)
            max_width = 300
            max_height = 400
            
            # Calculate new dimensions while maintaining aspect ratio
            original_width = img.width
            original_height = img.height
            # Calculate scaling factors
            width_ratio = max_width / original_width
            height_ratio = max_height / original_height
            
            # Use the smaller ratio to ensure image fits within bounds
            scale_factor = min(width_ratio, height_ratio)
            
            # Only resize if the image is larger than max dimensions
            if scale_factor < 1:
                img.width = int(original_width * scale_factor)
                img.height = int(original_height * scale_factor)
                print(f"Resized image from {original_width}x{original_height} to {img.width}x{img.height}")
            else:
                print(f"Image size {original_width}x{original_height} is within limits, no resizing needed")
            
            # Center the image in the cell
            if i == 0:
                # Get cell dimensions for K16
                cell = sheet['O16']
                # Calculate cell width and height (approximate)
                cell_width = 64  # Default column width in pixels
                cell_height = 20  # Default row height in pixels
                
                # Calculate offset to center the image
                x_offset = (cell_width - img.width) / 2
                y_offset = (cell_height - img.height) / 2
                
                # Position image with offset
                img.anchor = 'O16'
                img.left = x_offset
                img.top = y_offset
                sheet.add_image(img)
            # else:
            #     # Get cell dimensions for D16
            #     cell = sheet['D16']
            #     # Calculate cell width and height (approximate)
            #     cell_width = 64  # Default column width in pixels
            #     cell_height = 20  # Default row height in pixels
                
            #     # Calculate offset to center the image
            #     x_offset = (cell_width - img.width) / 2
            #     y_offset = (cell_height - img.height) / 2
                
            #     # Position image with offset
            #     img.anchor = 'D16'
            #     img.left = x_offset
            #     img.top = y_offset
            #     sheet.add_image(img)
        
        print(f"Added images for {sheet_id}")
        return True
        
    except Exception as e:
        print(f"Error adding image for {sheet_id}: {e}")
        return False

def create_sheets(input_wb, img_dir, input_file):
    """Create sheets directly from Schedule sheet data using openpyxl"""
    template_sheet_name = "Template option 6 MM"
    schedule_sheet_name = "Decision Matrix v.02"

    if template_sheet_name not in input_wb.sheetnames:
        print("Template sheet not found")
        return False
    
    if schedule_sheet_name not in input_wb.sheetnames:
        print("Schedule sheet not found")
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
                add_image_to_sheet(new_sheet, sheet_id, img_dir)
                
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
                add_image_to_sheet(new_sheet, sheet_id, img_dir)
                
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

def create_pdf(excel_file_path, sheet_ids: List[str]):
    """Create PDF from all sheets"""
    output_pdf = os.path.splitext(excel_file_path)[0] + "_output.pdf"
    if os.path.exists(output_pdf):
        os.remove(output_pdf)
    print(f"Creating PDF: {output_pdf}")
    
    try:
        # Use Excel automation
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        
        # Open workbook
        workbook = excel_app.Workbooks.Open(os.path.abspath(excel_file_path))

        # Define sheets to include in PDF export
        sheets_to_include = ["Cover", "GenInfo+Contacts"]
        sheets_to_include = sheets_to_include + sheet_ids
        
        print(f"Sheets to include in PDF: {sheets_to_include}")
        
        # Hide sheets that should not be included in PDF
        for sheet in workbook.Sheets:
            if sheet.Name not in sheets_to_include:
                sheet.Visible = False
                print(f"Hidden sheet: {sheet.Name}")
        
        # Export to PDF (only visible sheets will be included)
        workbook.ExportAsFixedFormat(
            Type=0,  # PDF
            Filename=os.path.abspath(output_pdf),
            Quality=0,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False
        )
        
        # Make all sheets visible again
        for sheet in workbook.Sheets:
            sheet.Visible = True
        
        # Clean up
        workbook.Close(SaveChanges=False)
        excel_app.Quit()
        
        print(f"PDF created successfully: {output_pdf}")
        return output_pdf
        
    except Exception as e:
        print(f"Error creating PDF: {e}")
        try:
            if 'workbook' in locals():
                workbook.Close(SaveChanges=False)
            if 'excel_app' in locals():
                excel_app.Quit()
        except:
            pass
        return False

def process_excel_file(input_file, img_dir):
    """Main processing function"""
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
    
    sheets_to_keep = ["Cover", "GenInfo+Contacts", "Template option 6 MM", "Decision Matrix v.02"]
    for sheet in input_wb.sheetnames:
        if sheet not in sheets_to_keep:
            input_wb.remove(input_wb[sheet])
    
    # Create sheets
    sheet_ids = create_sheets(input_wb, img_dir, input_file)
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