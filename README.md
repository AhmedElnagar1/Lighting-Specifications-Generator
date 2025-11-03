https://burohappold.sharepoint.com/:f:/r/sites/062646/Shared%20Documents/Lighting/++LED%20Lamp%20Replacement++/05_Presentations?csf=1&web=1&e=wYFkAY

# Lighting Specifications Generator

A Python application that processes Excel files to generate lighting specifications with automatic PDF output. The application provides both a command-line interface and a modern PyQt6 GUI for easy file selection and processing.

## Features

- **Automated Sheet Creation**: Creates separate sheets for each lighting system ID found in the Decision Matrix
- **Template-Based Processing**: Uses a template sheet to generate consistent specification sheets
- **Image Integration**: Automatically loads and inserts images based on system IDs from a specified directory
- **Data Mapping**: Maps data from Decision Matrix rows to template cells, including base configurations and two option variations
- **Automatic PDF Generation**: Creates a PDF output containing all generated sheets
- **GUI Interface**: Modern PyQt6 interface with file browser and progress tracking
- **Multi-threaded Processing**: Background processing to keep GUI responsive
- **Backup Creation**: Automatically creates timestamped backups before processing
- **Error Handling**: Comprehensive error handling and logging throughout
- **Page Footers**: Adds creation date and page numbering to generated sheets

## Requirements

- Python 3.10 or higher
- Microsoft Excel (required for PDF generation via COM automation)
- Windows OS (due to Excel automation dependency)

### Installation

Install the required dependencies using `uv` (recommended):

```bash
uv sync
```

Or using `pip`:

```bash
pip install openpyxl>=3.1.5 pandas>=2.3.1 pillow>=11.3.0 pywin32>=311 PyQt6>=6.5.0
```

### Building Executable

To create a standalone executable, first install the build dependencies:

```bash
uv sync --extra build
```

Then create a PyInstaller spec file and build:

```bash
pyinstaller --name Lighting_Specifications_Generator --windowed app.py
```

The executable will be created in the `dist` folder.

## Usage

### GUI Application (Recommended)

1. Run the GUI application:
```bash
python app.py
```

2. Use the interface to:
   - Browse and select your Excel file (`.xlsx`, `.xlsm`, or `.xls`)
   - Browse and select the image directory containing subdirectories named by system IDs
   - Click "Process Excel File" to start processing
   - Monitor progress in the status area
   - Open the generated PDF using the "Open PDF" button

### Command Line Interface

1. Ensure your Excel file contains:
   - A sheet named "Decision Matrix v.02" with the lighting system data
   - A sheet named "Template option 6 MM" that will be used as the template
   - Sheets named "Cover" and "GenInfo+Contacts" (will be included in PDF)

2. Ensure the image directory structure:
   - Images should be organized in subdirectories named by system ID (e.g., `LC-01/`, `LC-02/`)
   - Images can be named with "site" or "plan" in the filename for automatic categorization

3. Run the script with the file paths (or modify the paths in `final_excel_processor.py`):

```bash
python final_excel_processor.py
```

## How it works

1. **Creates Backup**: Automatically creates a timestamped backup of the input Excel file
2. **Loads Workbook**: Opens the Excel file and identifies required sheets:
   - "Decision Matrix v.02" - Contains the lighting system data
   - "Template option 6 MM" - Template used for generating new sheets
   - "Cover" and "GenInfo+Contacts" - Included in PDF output
3. **Cleans Existing Sheets**: Removes any previously generated sheets (keeps only essential sheets)
4. **Extracts Data**: Reads rows from Decision Matrix starting at row 11, grouping by Report Code:
   - Base rows (ending with "X")
   - Option 1 rows (ending with "1")
   - Option 2 rows (ending with "2")
5. **Creates New Sheets**: For each base row:
   - Copies the template sheet
   - Names it based on the Report Code (sanitized for Excel sheet name requirements)
   - Maps data from base, option 1, and option 2 rows to specific cells
   - Applies conditional formatting to "Assessed Condition" cells
6. **Adds Images**: For each sheet:
   - Looks for images in a subdirectory named by the system ID
   - Categorizes images as "site" or "plan" based on filename
   - Inserts site images at cells K7 and O7
   - Inserts plan images at cell O16
   - Automatically resizes images to fit within specified dimensions
7. **Adds Footers**: Adds page footers with creation date and page numbering to each generated sheet
8. **Generates PDF**: Uses Excel COM automation to:
   - Hide sheets that shouldn't be in the PDF
   - Export visible sheets to PDF
   - Restore all sheet visibility
   - Save the PDF next to the Excel file

## GUI Features

The PyQt6 GUI provides:
- **File Browser**: Easy selection of Excel files with file type filtering (`.xlsx`, `.xlsm`, `.xls`)
- **Image Directory Selection**: Browse and select the directory containing system ID subdirectories
- **Progress Tracking**: Real-time status updates and progress indication
- **PDF Output Display**: Shows the generated PDF path and provides a button to open it
- **Error Handling**: User-friendly error messages and warnings
- **Multi-threading**: Processing runs in background thread to keep GUI responsive
- **Timestamped Logs**: Status messages include timestamps for tracking

## Output

- **Modified Excel File**: The input Excel file is updated with new sheets for each lighting system ID found in the Decision Matrix
- **PDF Output**: A PDF file named `{input_filename}_output.pdf` containing:
  - Cover sheet
  - GenInfo+Contacts sheet
  - All generated specification sheets
- **Backup File**: A timestamped backup of the original Excel file (e.g., `filename_backup_20250103_143022.xlsx`)

## Error Handling

The script includes comprehensive error handling for:
- Missing Excel file or invalid file paths
- Missing required sheets ("Decision Matrix v.02", "Template option 6 MM")
- Missing or invalid image directories
- PDF generation failures (Excel COM automation issues)
- Excel file permission errors
- Image loading and resizing errors
- Sheet name conflicts (automatically handles duplicates with suffixes)

## Important Notes

- **Excel Required**: The script uses `win32com.client` for Excel automation, which requires Microsoft Excel to be installed on the system
- **Windows Only**: Due to Excel COM automation dependency, this application currently only works on Windows
- **File Permissions**: Ensure the Excel file is not open in another application during processing
- **Image Structure**: Images must be organized in subdirectories matching system IDs (e.g., `LC-01/`, `LW-02/`)
- **Sheet Naming**: Sheet names are sanitized to comply with Excel's 31-character limit and invalid character restrictions
- **Auto-Save**: The workbook is automatically saved after creating sheets
- **PDF Generation**: PDF generation uses Excel's built-in PDF export functionality via COM automation
- **Logging**: All operations are logged to the console for debugging purposes

## Troubleshooting

If you encounter issues:

1. **Excel not found**: 
   - Ensure Microsoft Excel is installed and properly configured
   - Try opening Excel manually to verify it's working
   - Check that Excel COM automation is available

2. **Permission errors**: 
   - Ensure the Excel file is not open in another application
   - Run as administrator if file access is restricted
   - Check that you have write permissions for the Excel file and output directory

3. **Missing dependencies**: 
   - Install using `uv sync` (recommended) or `pip install` with the packages from `pyproject.toml`
   - Ensure Python 3.10 or higher is being used

4. **PDF creation fails**: 
   - Verify Excel is installed and COM automation is working
   - Check if the output path is writable
   - Ensure Excel has permission to create files in the output directory
   - Try closing any other Excel instances

5. **GUI not starting**: 
   - Ensure PyQt6 is installed: `pip install PyQt6` or `uv sync`
   - Check Python version compatibility (3.10+)
   - Verify all dependencies are installed correctly

6. **Images not loading**: 
   - Verify the image directory structure matches system IDs
   - Check that image filenames contain "site" or "plan" keywords
   - Ensure image files are in supported formats (JPG, PNG, etc.)
   - Verify subdirectory names match the Report Codes from the Decision Matrix

7. **Sheet creation errors**: 
   - Verify the "Decision Matrix v.02" sheet exists and has data starting at row 11
   - Check that the "Template option 6 MM" sheet exists
   - Ensure column headers are in row 10 as expected

8. **Processing hangs**: 
   - Check that the Excel file is not open in another application
   - Verify Excel is responding (try opening it manually)
   - Look at console output for detailed error messages

   - Consider using the GUI which provides better progress feedback 
