# Excel Processor - VBA to Python Conversion

This Python script converts the functionality of the original VBA code to Python, with the added capability of automatically creating PDF output.

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

## Usage

1. Place the `Bauphase.xlsm` file in the same directory as the script
2. Ensure the `img/` folder with images is present
3. Run the script:

```bash
python excel_processor.py
```

## How it works

1. **Loads the Excel workbook** and identifies the Schedule and Template sheets
2. **Extracts IDs** from the Schedule sheet (looks for patterns like LC-, LW-, LT-, LJ-)
3. **Creates new sheets** by copying the Template sheet for each ID
4. **Sets selection values** in each new sheet to the corresponding ID
5. **Generates PDF** automatically from all sheets using Excel automation

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

## Troubleshooting

If you encounter issues:

1. **Excel not found**: Ensure Microsoft Excel is installed
2. **Permission errors**: Run as administrator if needed
3. **Missing dependencies**: Install requirements with `pip install -r requirements.txt`
4. **PDF creation fails**: Check if the output path is writable and Excel has permission to create files 