"""
Excel processor module - backward compatibility wrapper.

This module maintains backward compatibility by re-exporting functions
from the new excel_processor package structure.
"""

# Import all functions and classes from the new module structure
from excel_processor import (
    process_excel_file,
    find_template_sheets,
    find_decision_matrix_sheet,
    add_cover_image,
    create_backup,
    add_page_footer,
    ExcelProcessor,
    ImageHandler,
    SheetFinder,
    DataMapper,
    SheetCreator,
    PDFExporter,
    DebugLogger
)

# Re-export for backward compatibility
__all__ = [
    'process_excel_file',
    'find_template_sheets',
    'find_decision_matrix_sheet',
    'add_cover_image',
    'create_backup',
    'add_page_footer',
    'ExcelProcessor',
    'ImageHandler',
    'SheetFinder',
    'DataMapper',
    'SheetCreator',
    'PDFExporter',
    'DebugLogger',
]


if __name__ == "__main__":
    input_file = r"C:\Users\vmylavarapu\Buro Happold\Germany Computational Team - General\3 Development\5 Lighting\Lighting Decision Matrix Report Generator\0062646_Microsoft Berlin_UdL_Decision Matrix_v.03_backup_20251113_141141 - Copy - Copy.xlsx"
    img_dir = r"C:\Users\vmylavarapu\Buro Happold\Germany Computational Team - General\3 Development\5 Lighting\Lighting Decision Matrix Report Generator\Specifications"
    #excel_file = r"C:\Users\aelnagar\Downloads\Lighting Computational Development\Bauphase.xlsm"
    process_excel_file(input_file, img_dir) 
