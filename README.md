# Lighting Specifications Generator

A Python application that processes Excel files to generate lighting specifications with automatic PDF output. The application provides both a command-line interface and a modern PyQt6 GUI for easy file selection and language processing.

## Features

The application provides the following functionality:

1. **CreateAllSheets()** - Creates separate tabs for each ID found in the Schedule sheet
2. **ClearAllSheets()** - Removes all automatically created catalogue sheets
3. **WorksheetExists()** - Checks if a worksheet exists
4. **Image handling** - Loads images based on selection (equivalent to VBA image loading)
5. **GUI Interface** - Modern PyQt6 interface for easy file selection and language choice
6. **Multi-language support** - Process files in English or German
7. **Automatic PDF generation** - Creates PDF output from all processed sheets

**Additional Python Features:**
- Automatic PDF generation from all created sheets
- No user input required - fully automated process
- Better error handling and logging
- Modern GUI interface with file browser
- Multi-threaded processing to prevent GUI freezing

## Features

The Python script provides the same functionality as the original VBA code:

1. **CreateAllSheets()** - Creates separate tabs for each ID found in the Schedule sheet
2. **ClearAllSheets()** - Removes all automatically created catalogue sheets
3. **WorksheetExists()** - Checks if a worksheet exists
4. **Image handling** - Loads images based on selection (equivalent to VBA image loading)

**Additional Python Features:**
- Automatic PDF generation from all created sheets
- No user input required - fully automated process
- Better error handling and logging

## Requirements

Install the required dependencies:

```bash
pip install -r requirements.txt
```

Or using uv (recommended):

```bash
uv sync
```

### Building Executable

To create a standalone executable, install the build dependencies:

```bash
uv sync --extra build
```

Then build the executable:

```bash
pyinstaller app.spec
```

The executable will be created in the `dist` folder as `Lighting_Specifications_Generator.exe`.

## Usage

### GUI Application (Recommended)

1. Run the GUI application:
```bash
python app.py
```

2. Use the interface to:
   - Browse and select your Excel file (`.xlsx`, `.xlsm`, or `.xls`)
   - Choose between English or German language
   - Click "Process Excel File" to start processing
   - Monitor progress in the status area

### Command Line Interface

1. Place the `Bauphase.xlsm` file in the same directory as the script
2. Ensure the `img/` folder with images is present
3. Run the script:

```bash
python final_excel_processor.py
```

## How it works

1. **Loads the Excel workbook** and identifies the Schedule and Template sheets
2. **Extracts IDs** from the Schedule sheet (looks for patterns like LC-, LW-, LT-, LJ-)
3. **Creates new sheets** by copying the Template sheet for each ID
4. **Sets selection values** in each new sheet to the corresponding ID
5. **Adds images** to each sheet based on the ID (looks for corresponding image files)
6. **Generates PDF** automatically from all sheets using Excel automation

## GUI Features

The PyQt6 GUI provides:
- **File Browser**: Easy selection of Excel files with file type filtering
- **Language Selection**: Dropdown to choose between English and German processing
- **Progress Tracking**: Real-time status updates and progress indication
- **Error Handling**: User-friendly error messages and warnings
- **Multi-threading**: Processing runs in background thread to keep GUI responsive

## Output

- Modified `Bauphase.xlsm` with new sheets for each ID
- `Bauphase_output.pdf` containing all sheets

## Error Handling

The script includes comprehensive error handling for:
- Missing Excel file
- Missing required sheets (Schedule, Template)
- PDF generation failures
- Excel automation issues

## Notes

- The script uses `win32com.client` for Excel automation, which requires Excel to be installed on the system
- The script automatically saves the workbook after creating sheets
- PDF generation uses Excel's built-in PDF export functionality
- All operations are logged to the console for debugging

## Testing

Run the unit tests to verify functionality:

```bash
python -m pytest test_app.py -v
```

Or run with unittest:

```bash
python -m unittest test_app.py -v
```

## Troubleshooting

If you encounter issues:

1. **Excel not found**: Ensure Microsoft Excel is installed
2. **Permission errors**: Run as administrator if needed
3. **Missing dependencies**: Install requirements with `pip install -r requirements.txt` or `uv sync`
4. **PDF creation fails**: Check if the output path is writable and Excel has permission to create files
5. **GUI not starting**: Ensure PyQt6 is installed: `pip install PyQt6`
6. **Processing hangs**: Check that the Excel file is not open in another application 