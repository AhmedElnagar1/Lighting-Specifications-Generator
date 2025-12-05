"""
Excel processor module for handling Excel file operations.

This module provides classes and functions for processing Excel files,
creating sheets, mapping data, and generating PDFs.
"""

from excel_processor.excel_processor import (
    ExcelProcessor,
    process_excel_file,
    find_template_sheets,
    find_decision_matrix_sheet,
    add_cover_image,
    create_backup,
    add_page_footer
)
from excel_processor.image_handler import ImageHandler
from excel_processor.sheet_finder import SheetFinder
from excel_processor.data_mapper import DataMapper
from excel_processor.sheet_creator import SheetCreator
from excel_processor.pdf_exporter import PDFExporter
from excel_processor.debug_logger import DebugLogger

__all__ = [
    'ExcelProcessor',
    'ImageHandler',
    'SheetFinder',
    'DataMapper',
    'SheetCreator',
    'PDFExporter',
    'DebugLogger',
    'process_excel_file',
    'find_template_sheets',
    'find_decision_matrix_sheet',
    'add_cover_image',
    'create_backup',
    'add_page_footer',
]

