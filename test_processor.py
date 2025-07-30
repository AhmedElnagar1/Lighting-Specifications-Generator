import os
import shutil
from excel_processor import ExcelProcessor

def test_excel_processor():
    """Test the Excel processor functionality"""
    
    # Check if the Excel file exists
    excel_file = "Bauphase.xlsm"
    if not os.path.exists(excel_file):
        print(f"Error: {excel_file} not found")
        return False
    
    # Create a backup copy for testing
    test_file = "Bauphase_test.xlsm"
    shutil.copy2(excel_file, test_file)
    print(f"Created test copy: {test_file}")
    
    try:
        # Initialize processor with test file
        processor = ExcelProcessor(test_file)
        
        # Test loading workbook
        print("\n1. Testing workbook loading...")
        if not processor.load_workbook():
            print("Failed to load workbook")
            return False
        
        # Test getting ID list
        print("\n2. Testing ID list extraction...")
        ids = processor.get_id_list()
        print(f"Found IDs: {ids}")
        
        if not ids:
            print("No IDs found - this might be normal if the Schedule sheet doesn't contain ID patterns")
        
        # Test worksheet existence check
        print("\n3. Testing worksheet existence...")
        test_sheet = "Template"
        exists = processor.worksheet_exists(test_sheet)
        print(f"Template sheet exists: {exists}")
        
        # Test sheet creation (only if IDs were found)
        if ids:
            print("\n4. Testing sheet creation...")
            success = processor.create_all_sheets()
            print(f"Sheet creation successful: {success}")
            
            # Test PDF creation
            print("\n5. Testing PDF creation...")
            pdf_success = processor.create_pdf("test_output.pdf")
            print(f"PDF creation successful: {pdf_success}")
            
            if pdf_success and os.path.exists("test_output.pdf"):
                print("PDF file created successfully!")
        
        print("\nTest completed successfully!")
        return True
        
    except Exception as e:
        print(f"Test failed with error: {e}")
        return False
    
    finally:
        # Clean up test files
        try:
            if os.path.exists(test_file):
                os.remove(test_file)
                print(f"Cleaned up: {test_file}")
            if os.path.exists("test_output.pdf"):
                os.remove("test_output.pdf")
                print("Cleaned up: test_output.pdf")
        except Exception as e:
            print(f"Warning: Could not clean up test files: {e}")

if __name__ == "__main__":
    test_excel_processor() 