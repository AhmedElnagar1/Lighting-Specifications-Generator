"""
PDF exporter module for creating PDFs from Excel sheets.
"""

from typing import List, Union
import os
import win32com.client


class PDFExporter:
    """
    Handles PDF creation from Excel sheets.
    
    This class provides methods for exporting Excel sheets to PDF format
    using Excel COM automation.
    """
    
    def __init__(self) -> None:
        """
        Initialize the PDF exporter.
        """
        pass
    
    def create_pdf(self, excel_file_path: str, sheet_ids: List[str]) -> Union[str, str]:
        """
        Create PDF from all sheets using Excel COM automation.
        
        Args:
            excel_file_path (str): Path to the Excel file
            sheet_ids (List[str]): List of sheet IDs to include in PDF
        
        Returns:
            Union[str, str]: Path to the generated PDF file on success, or error message string on error
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
            actual_sheet_names: List[str] = []
            for sheet in workbook.Sheets:
                actual_sheet_names.append(sheet.Name)
            
            print(f"Actual sheets in workbook: {actual_sheet_names}")
            
            # Define sheets to include in PDF export
            sheets_to_include: List[str] = ["Cover", "GenInfo+Contacts"]
            sheets_to_include = sheets_to_include + sheet_ids
            
            # Filter to only include sheets that actually exist
            existing_sheets_to_include = [s for s in sheets_to_include if s in actual_sheet_names]
            missing_sheets = [s for s in sheets_to_include if s not in actual_sheet_names]
            
            if missing_sheets:
                error_msg = f"The following sheets are missing from the workbook and cannot be included in the PDF:\n"
                error_msg += "\n".join(f"  - {sheet}" for sheet in missing_sheets)
                print(f"Warning: {error_msg}")
                # Still continue if we have some sheets to include
            
            if not existing_sheets_to_include:
                error_msg = "No valid sheets found to include in PDF. Please ensure the required sheets exist."
                raise Exception(error_msg)
            
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
            error_msg = f"Error creating PDF: {str(e)}"
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
            
            return error_msg


