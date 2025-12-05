"""
Sheet creator module for creating Excel sheets from template.
"""

from typing import List, Union, Dict, Any, Optional
import os
import re
from datetime import datetime
from openpyxl import Workbook
from excel_processor.image_handler import ImageHandler
from excel_processor.data_mapper import DataMapper
from excel_processor.debug_logger import DebugLogger


class SheetCreator:
    """
    Handles creating Excel sheets from templates.
    
    This class provides methods for creating new sheets based on template
    and decision matrix data.
    """
    
    def __init__(
        self,
        image_handler: Optional[ImageHandler] = None,
        data_mapper: Optional[DataMapper] = None,
        debug_logger: Optional[DebugLogger] = None
    ) -> None:
        """
        Initialize the sheet creator.
        
        Args:
            image_handler (Optional[ImageHandler]): Image handler instance. If None, creates a new one.
            data_mapper (Optional[DataMapper]): Data mapper instance. If None, creates a new one.
            debug_logger (Optional[DebugLogger]): Debug logger instance. If None, creates a new one.
        """
        self.image_handler = image_handler if image_handler is not None else ImageHandler()
        self.debug_logger = debug_logger if debug_logger is not None else DebugLogger()
        self.data_mapper = data_mapper if data_mapper is not None else DataMapper(self.debug_logger)
    
    def extract_bold_words_from_template(self, template_sheet: object) -> List[str]:
        """
        Extract all words from cells with bold formatting in the template sheet.
        
        Args:
            template_sheet: The openpyxl worksheet object for the template
        
        Returns:
            List[str]: List of bold words found in the template sheet
        """
        bold_words: List[str] = []
        
        for row in template_sheet.iter_rows():
            for cell in row:
                if cell.value is not None and cell.font and cell.font.bold:
                    cell_value = str(cell.value).strip()
                    if cell_value:
                        bold_words.append(cell_value)
        
        return bold_words
    
    def find_header_row_by_bold_words(
        self,
        schedule_sheet: object,
        bold_words: List[str],
        min_matches: int = 7
    ) -> Optional[int]:
        """
        Find the header row in the Decision Matrix by matching bold words from template.
        
        Searches each row in the Decision Matrix sheet and returns the row number
        that contains at least the specified minimum number of bold words from the template.
        
        Args:
            schedule_sheet: The openpyxl worksheet object for the Decision Matrix
            bold_words (List[str]): List of bold words extracted from the template
            min_matches (int): Minimum number of bold words that must match (default: 7)
        
        Returns:
            Optional[int]: The row number containing the header, or None if not found
        """
        # Normalize bold words for case-insensitive comparison
        normalized_bold_words: set[str] = {word.lower().strip() for word in bold_words}
        
        for row_num in range(1, schedule_sheet.max_row + 1):
            match_count: int = 0
            
            for col_num in range(1, schedule_sheet.max_column + 1):
                cell_value = schedule_sheet.cell(row=row_num, column=col_num).value
                if cell_value is not None:
                    cell_text = str(cell_value).strip().lower()
                    if cell_text in normalized_bold_words:
                        match_count += 1
            
            if match_count >= min_matches:
                print(f"Found header row at row {row_num} with {match_count} matching bold words")
                return row_num
        
        print(f"Warning: No row found with at least {min_matches} matching bold words")
        return None
    
    def add_page_footer(self, sheet: object, page_number: int, total_pages: int) -> None:
        """
        Add a footer with creation date on the left and page numbering on the right.
        
        Args:
            sheet: The openpyxl worksheet object
            page_number (int): Current page number
            total_pages (int): Total number of pages
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
    
    def create_sheets(
        self,
        input_wb: Workbook,
        img_dir: str,
        input_file: str,
        template_sheet_name: str,
        schedule_sheet_name: str
    ) -> Union[List[str], str]:
        """
        Create sheets directly from Schedule sheet data using openpyxl.
        
        Args:
            input_wb: The openpyxl workbook object
            img_dir (str): Path to the image directory
            input_file (str): Path to the input Excel file
            template_sheet_name (str): Name of the template sheet to use
            schedule_sheet_name (str): Name of the Decision Matrix sheet to use
        
        Returns:
            Union[List[str], str]: List of created sheet IDs on success, or error message string on error
        """
        if template_sheet_name not in input_wb.sheetnames:
            error_msg = f"Template sheet '{template_sheet_name}' not found in the workbook."
            print(error_msg)
            return error_msg
        
        if schedule_sheet_name not in input_wb.sheetnames:
            error_msg = f"Decision Matrix sheet '{schedule_sheet_name}' not found in the workbook."
            print(error_msg)
            return error_msg
        
        try:
            # Initialize debug file
            self.debug_logger.write("=" * 50)
            self.debug_logger.write(f"Starting sheet creation process at {datetime.now()}")
            self.debug_logger.write("=" * 50)
            
            # Get the template sheet
            template_sheet = input_wb[template_sheet_name]
            
            # Get the Schedule sheet
            schedule_sheet = input_wb[schedule_sheet_name]
            sheets_created = 0
            
            # Extract bold words from template sheet
            bold_words = self.extract_bold_words_from_template(template_sheet)
            print(f"Extracted {len(bold_words)} bold words from template")
            self.debug_logger.write(f"Bold words from template: {bold_words}")
            
            # Find the header row dynamically by matching bold words
            header_row = self.find_header_row_by_bold_words(schedule_sheet, bold_words, min_matches=7)
            if header_row is None:
                error_msg = "Could not find header row in Decision Matrix sheet (no row with at least 7 matching bold words)"
                print(error_msg)
                self.debug_logger.write(error_msg)
                return error_msg
            
            print(f"Using header row: {header_row}")
            self.debug_logger.write(f"Found header row at row {header_row}")
            
            # Get column names from the dynamically found header row
            column_names: Dict[int, str] = {}
            for col_num in range(1, schedule_sheet.max_column + 1):
                cell_value = schedule_sheet.cell(row=header_row, column=col_num).value
                if cell_value and (cell_value != 'Eingebunden in Steuerung'):
                    column_names[col_num] = str(cell_value).strip()
            
            print(f"Found columns: {list(column_names.values())}")
            
            # Process each row in the Schedule sheet starting from the row after header
            base_rows: List[Dict[str, Any]] = []
            option_1_rows: List[Dict[str, Any]] = []
            option_2_rows: List[Dict[str, Any]] = []
            sheet_ids: List[str] = []
            for row_num in range(header_row + 1, schedule_sheet.max_row + 1):
                # Get all values for this row
                row_data: Dict[str, Any] = {}
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
                    self.data_mapper.map_base_data_to_template(new_sheet, base_row, option_1_row, option_2_row, schedule_sheet)
                    
                    # Add image to the sheet
                    success, temp_files = self.image_handler.add_image_to_sheet(new_sheet, sheet_id, img_dir)
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
                    self.data_mapper.map_base_data_to_template(new_sheet, base_row, option_1_row, option_2_row, schedule_sheet)
                    
                    # Add image to the sheet
                    success, temp_files = self.image_handler.add_image_to_sheet(new_sheet, sheet_id, img_dir)
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
                    self.add_page_footer(input_wb[sheet_id], page_number, total_pages)
            
            # Save the workbook
            input_wb.save(input_file)
            print(f"Created {sheets_created} new sheets")
            print("Workbook saved successfully")
            
            # Close debug file
            self.debug_logger.close()
            
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
            error_msg = f"Error creating sheets: {str(e)}"
            print(error_msg)
            self.debug_logger.write(f"Error creating sheets: {error_msg}")
            self.debug_logger.close()
            print("Trying to save with different name...")
            
            # Try saving with a different name
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_path = f"{os.path.splitext(input_file)[0]}_modified_{timestamp}.xlsx"
            
            try:
                input_wb.save(new_path)
                print(f"Workbook saved as: {new_path}")
                return new_path
            except Exception as e2:
                error_msg = f"Error saving to new path: {str(e2)}"
                print(error_msg)
                return error_msg

