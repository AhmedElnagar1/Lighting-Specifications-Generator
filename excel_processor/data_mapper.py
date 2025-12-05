"""
Data mapper module for mapping data from decision matrix to template sheets.
"""

from typing import Any, Dict, List, Optional
import re
from difflib import SequenceMatcher
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font
from excel_processor.debug_logger import DebugLogger


class DataMapper:
    """
    Handles mapping data from decision matrix to template sheets.
    
    This class provides methods for matching keywords and placing values
    in the appropriate cells of template sheets.
    """
    
    def __init__(self, debug_logger: Optional[DebugLogger] = None) -> None:
        """
        Initialize the data mapper.
        
        Args:
            debug_logger (Optional[DebugLogger]): Debug logger instance. If None, creates a new one.
        """
        self.debug_logger = debug_logger if debug_logger is not None else DebugLogger()
    
    def _clean_keyword(self, keyword: str) -> str:
        """
        Clean a keyword by removing special characters like *, \n, and extra whitespace.
        
        Args:
            keyword (str): The keyword to clean
        
        Returns:
            str: The cleaned keyword
        """
        # Remove newlines and replace with space
        cleaned = keyword.replace('\n', ' ').replace('\r', ' ')
        # Remove asterisks and other special characters
        cleaned = re.sub(r'[*]', '', cleaned)
        # Remove extra whitespace and strip
        cleaned = re.sub(r'\s+', ' ', cleaned).strip()
        return cleaned
    
    def _find_exact_match(self, keyword: str, cell_value: str) -> bool:
        """
        Check if a keyword exactly matches a cell value (after cleaning).
        
        Args:
            keyword (str): The keyword to search for
            cell_value (str): The cell value to search in
        
        Returns:
            bool: True if exact match is found, False otherwise
        """
        # Clean both strings for comparison
        keyword_clean = self._clean_keyword(keyword).lower()
        cell_clean = self._clean_keyword(cell_value).lower()
        
        # Check exact match
        return keyword_clean == cell_clean
    
    def _find_closest_match(self, keyword: str, cell_value: str, threshold: float = 0.85) -> bool:
        """
        Find if a keyword closely matches a cell value using fuzzy matching.
        Only used when exact match is not found.
        
        Args:
            keyword (str): The keyword to search for
            cell_value (str): The cell value to search in
            threshold (float): Similarity threshold (0.0 to 1.0), default 0.85
        
        Returns:
            bool: True if a close match is found, False otherwise
        """
        # Clean both strings for comparison
        keyword_clean = self._clean_keyword(keyword).lower()
        cell_clean = self._clean_keyword(cell_value).lower()
        
        # Use SequenceMatcher for fuzzy matching
        similarity = SequenceMatcher(None, keyword_clean, cell_clean).ratio()
        
        # Also check word-by-word matching
        keyword_words = set(keyword_clean.split())
        cell_words = set(cell_clean.split())
        
        # If there are common significant words (more than 2 characters)
        significant_keyword_words = {w for w in keyword_words if len(w) > 2}
        significant_cell_words = {w for w in cell_words if len(w) > 2}
        
        if significant_keyword_words and significant_cell_words:
            common_words = significant_keyword_words.intersection(significant_cell_words)
            if common_words:
                # If significant words match, consider it a match
                word_match_ratio = len(common_words) / max(len(significant_keyword_words), len(significant_cell_words))
                if word_match_ratio >= 0.5:
                    return True
        
        # Return True if similarity is above threshold
        return similarity >= threshold
    
    def _add_hyperlink_to_cell(self, sheet: object, row: int, column: int, link_value: str) -> None:
        """
        Add a hyperlink to a cell with display text "Link".
        
        Args:
            sheet: The openpyxl worksheet object
            row (int): Row number (1-indexed)
            column (int): Column number (1-indexed)
            link_value (str): The URL or link value to add as hyperlink
        """
        cell = sheet.cell(row=row, column=column)
        
        # Check if the value looks like a URL
        link_str = str(link_value).strip()
        if link_str and (link_str.startswith('http://') or link_str.startswith('https://')):
            # Set the hyperlink
            cell.hyperlink = link_str
            cell.value = "Link"
            # Style the cell to look like a hyperlink (blue, underlined)
            cell.font = Font(color="0563C1", underline="single")
        else:
            # If it's not a URL, just set the value as is
            cell.value = link_value
    
    def _place_value_in_cell(self, sheet: object, row: int, col: int, keyword: str, value: Any) -> None:
        """
        Place a value in a cell with appropriate formatting based on the keyword.
        
        Args:
            sheet: The openpyxl worksheet object
            row (int): Row number (1-indexed)
            col (int): Column number (1-indexed)
            keyword (str): The keyword associated with this value
            value (Any): The value to place in the cell
        """
        # Handle special formatting for certain fields
        if keyword == "Assessed Condition" and isinstance(value, str):
            sheet.cell(row=row, column=col).value = value
            # Apply color coding
            if "1" in value:
                fill = PatternFill(start_color="00E668", end_color="00E668", fill_type="solid")
                sheet.cell(row=row, column=col).fill = fill
            elif "2" in value:
                fill = PatternFill(start_color="BAE18F", end_color="BAE18F", fill_type="solid")
                sheet.cell(row=row, column=col).fill = fill
            elif "3" in value:
                fill = PatternFill(start_color="F7C7AC", end_color="F7C7AC", fill_type="solid")
                sheet.cell(row=row, column=col).fill = fill
            elif "4" in value:
                fill = PatternFill(start_color="F1A983", end_color="F1A983", fill_type="solid")
                sheet.cell(row=row, column=col).fill = fill
            elif "5" in value:
                fill = PatternFill(start_color="FF7171", end_color="FF7171", fill_type="solid")
                sheet.cell(row=row, column=col).fill = fill
        elif keyword == "System Power Consumption":
            # Format as "XX W" if it's a number
            if isinstance(value, (int, float)) and value != 0:
                sheet.cell(row=row, column=col).value = f"{value} W"
            else:
                sheet.cell(row=row, column=col).value = value
        elif keyword == "System link":
            # Handle hyperlinks
            self._add_hyperlink_to_cell(sheet, row=row, column=col, link_value=value)
        else:
            sheet.cell(row=row, column=col).value = value
        self.debug_logger.write(f"Placed value for '{keyword}' at row {row}, column {col}")
    
    def map_base_data_to_template(
        self,
        sheet: object,
        base_row_data: Dict[str, Any],
        option_1_row_data: Dict[str, Any],
        option_2_row_data: Dict[str, Any],
        decision_matrix_sheet: object
    ) -> None:
        """
        Map data from Decision Matrix to template fields using dynamic keyword matching.
        
        In the decision matrix sheet, finds the cell containing "Report Code" in the first column.
        Extracts keywords from that row. Then:
        1. Finds "System Description" in template sheet and fills values below all keywords in that row
        2. Finds "Technical parameters" and fills values in Option 1/Option 2 columns based on keywords below it
        
        Args:
            sheet: The openpyxl worksheet object (template sheet)
            base_row_data (Dict[str, Any]): Dictionary containing base row data from decision matrix
            option_1_row_data (Dict[str, Any]): Dictionary containing option 1 row data from decision matrix
            option_2_row_data (Dict[str, Any]): Dictionary containing option 2 row data from decision matrix
            decision_matrix_sheet: The openpyxl worksheet object for the decision matrix sheet
        """
        self.debug_logger.write(f"Option 1 row data: {option_1_row_data}")
        self.debug_logger.write(f"Option 2 row data: {option_2_row_data}")
        try:
            # Step 1: Find "Report Code" in the first column of decision matrix sheet
            report_code_row: Optional[int] = None
            for row_num in range(1, decision_matrix_sheet.max_row + 1):
                cell_value = decision_matrix_sheet.cell(row=row_num, column=1).value
                if cell_value is not None:
                    cell_value_str = str(cell_value).strip()
                    if "Report Code" in cell_value_str:
                        report_code_row = row_num
                        break
            
            if report_code_row is None:
                self.debug_logger.write("Warning: Could not find 'Report Code' in first column of decision matrix sheet")
                return
            
            # Step 2: Extract keywords from that row (all cell values in that row)
            keywords: List[str] = []
            keywords_original: List[str] = []  # Keep original for data lookup
            for col_num in range(1, decision_matrix_sheet.max_column + 1):
                cell_value = decision_matrix_sheet.cell(row=report_code_row, column=col_num).value
                if cell_value is not None:
                    keyword_original = str(cell_value).strip()
                    if keyword_original:  # Only add non-empty keywords
                        # Clean the keyword for matching, but keep original for data lookup
                        keyword_cleaned = self._clean_keyword(keyword_original)
                        keywords.append(keyword_cleaned)
                        keywords_original.append(keyword_original)
            
            self.debug_logger.write(f"Found keywords from decision matrix row {report_code_row}: {keywords_original}")
            
            # Step 3: Find "System Description" in template sheet
            system_description_row: Optional[int] = None
            system_description_col: Optional[int] = None
            for row_num in range(1, sheet.max_row + 1):
                for col_num in range(1, sheet.max_column + 1):
                    cell_value = sheet.cell(row=row_num, column=col_num).value
                    if cell_value is not None:
                        cell_value_str = str(cell_value).strip()
                        if "System Description" in cell_value_str:
                            system_description_row = row_num
                            system_description_col = col_num
                            break
                if system_description_row is not None:
                    break
            
            # Step 4: For all keywords found in the System Description row, fill values below them
            if system_description_row is not None:
                # Find all keywords in that row
                for col_num in range(1, sheet.max_column + 1):
                    cell_value = sheet.cell(row=system_description_row, column=col_num).value
                    if cell_value is not None:
                        cell_value_str = str(cell_value).strip()
                        # Check if this cell contains any of our keywords - try exact match first, then fuzzy matching
                        matched_keyword: Optional[str] = None
                        matched_keyword_original: Optional[str] = None
                        
                        # First, try exact match
                        for i, keyword_cleaned in enumerate(keywords):
                            keyword_original = keywords_original[i]
                            if self._find_exact_match(keyword_cleaned, cell_value_str):
                                matched_keyword = keyword_cleaned
                                matched_keyword_original = keyword_original
                                self.debug_logger.write(f"Exact match found: '{keyword_original}' in cell at row {system_description_row}, column {col_num}")
                                break
                        
                        # If no exact match, try fuzzy matching
                        if matched_keyword is None:
                            best_match_score = 0.0
                            for i, keyword_cleaned in enumerate(keywords):
                                keyword_original = keywords_original[i]
                                if self._find_closest_match(keyword_cleaned, cell_value_str):
                                    # Calculate similarity score
                                    keyword_clean = self._clean_keyword(keyword_cleaned).lower()
                                    cell_clean = self._clean_keyword(cell_value_str).lower()
                                    similarity = SequenceMatcher(None, keyword_clean, cell_clean).ratio()
                                    
                                    if similarity > best_match_score:
                                        best_match_score = similarity
                                        matched_keyword = keyword_cleaned
                                        matched_keyword_original = keyword_original
                            
                            if matched_keyword is not None:
                                self.debug_logger.write(f"Fuzzy match found: '{matched_keyword_original}' (similarity: {best_match_score:.2f}) in cell at row {system_description_row}, column {col_num}")
                        
                        # If we found a match (exact or fuzzy), place the value
                        if matched_keyword_original is not None:
                            # Get value from base_row_data using original keyword (prioritize base data for System Description row)
                            value_to_place: Optional[Any] = None
                            if matched_keyword_original in base_row_data and base_row_data[matched_keyword_original] is not None:
                                value_to_place = base_row_data[matched_keyword_original]
                            elif matched_keyword_original in option_1_row_data and option_1_row_data[matched_keyword_original] is not None:
                                value_to_place = option_1_row_data[matched_keyword_original]
                            elif matched_keyword_original in option_2_row_data and option_2_row_data[matched_keyword_original] is not None:
                                value_to_place = option_2_row_data[matched_keyword_original]
                            
                            # Place value in cell below the keyword
                            if value_to_place is not None:
                                target_row = system_description_row + 1
                                self._place_value_in_cell(sheet, target_row, col_num, matched_keyword_original, value_to_place)
            
            # Step 5: Find "Technical parameters" (or "Technical paramaters") in template sheet
            technical_params_row: Optional[int] = None
            technical_params_col: Optional[int] = None
            for row_num in range(1, sheet.max_row + 1):
                for col_num in range(1, sheet.max_column + 1):
                    cell_value = sheet.cell(row=row_num, column=col_num).value
                    if cell_value is not None:
                        cell_value_str = str(cell_value).strip()
                        if "Technical Parameters" in cell_value_str or "Technical paramaters" in cell_value_str:
                            technical_params_row = row_num
                            technical_params_col = col_num
                            col_letter = get_column_letter(col_num)
                            self.debug_logger.write(f"Found 'Technical parameters' at cell {col_letter}{row_num} (row {row_num}, column {col_num})")
                            break
                if technical_params_row is not None:
                    break
            
            if technical_params_row is None:
                self.debug_logger.write("Warning - 'Technical parameters' cell not found")
            
            # Step 6: Find columns containing "Option 1" and "Option 2"
            option_1_col: Optional[int] = None
            option_2_col: Optional[int] = None
            option_1_row: Optional[int] = None
            option_2_row: Optional[int] = None
            for row_num in range(1, sheet.max_row + 1):
                for col_num in range(1, sheet.max_column + 1):
                    cell_value = sheet.cell(row=row_num, column=col_num).value
                    if cell_value is not None:
                        cell_value_str = str(cell_value).strip()
                        if "Option 1" in cell_value_str and option_1_col is None:
                            option_1_col = col_num
                            option_1_row = row_num
                            col_letter = get_column_letter(col_num)
                            self.debug_logger.write(f"Found 'Option 1' at cell {col_letter}{row_num} (row {row_num}, column {col_num})")
                            self.debug_logger.write(f"Option 1 column number: {col_num}, row number: {row_num}")
                        if "Option 2" in cell_value_str and option_2_col is None:
                            option_2_col = col_num
                            option_2_row = row_num
                            col_letter = get_column_letter(col_num)
                            self.debug_logger.write(f"Found 'Option 2' at cell {col_letter}{row_num} (row {row_num}, column {col_num})")
                            self.debug_logger.write(f"Option 2 column number: {col_num}, row number: {row_num}")
                    if option_1_col is not None and option_2_col is not None:
                        break
                if option_1_col is not None and option_2_col is not None:
                    break
            
            if option_1_col is None:
                self.debug_logger.write("Warning - 'Option 1' column not found")
            if option_2_col is None:
                self.debug_logger.write("Warning - 'Option 2' column not found")
            
            # Place "Option Name" values directly below "Option 1" and "Option 2" cells
            if option_1_row is not None and option_1_col is not None:
                if "Option Name" in option_1_row_data and option_1_row_data["Option Name"] is not None:
                    option_name_1 = option_1_row_data["Option Name"]
                    target_row = option_1_row + 1
                    self.debug_logger.write(f"Placing Option 1 Name '{option_name_1}' at row {target_row}, column {option_1_col}")
                    sheet.cell(row=target_row, column=option_1_col).value = option_name_1
                else:
                    self.debug_logger.write("Warning - 'Option Name' not found in option_1_row_data")
            
            if option_2_row is not None and option_2_col is not None:
                if "Option Name" in option_2_row_data and option_2_row_data["Option Name"] is not None:
                    option_name_2 = option_2_row_data["Option Name"]
                    target_row = option_2_row + 1
                    self.debug_logger.write(f"Placing Option 2 Name '{option_name_2}' at row {target_row}, column {option_2_col}")
                    sheet.cell(row=target_row, column=option_2_col).value = option_name_2
                else:
                    self.debug_logger.write("Warning - 'Option Name' not found in option_2_row_data")
            
            # Step 7: Search for all cells below "Technical parameters" in the same column
            # For each keyword found, fill value in corresponding row in Option 1 or Option 2 column
            if technical_params_row is not None and technical_params_col is not None:
                # Search all cells below Technical parameters in that column
                for row_num in range(technical_params_row + 1, sheet.max_row + 1):
                    cell_value = sheet.cell(row=row_num, column=technical_params_col).value
                    if cell_value is not None:
                        cell_value_str = str(cell_value).strip()
                        # Check if this cell contains any of our keywords - try exact match first, then fuzzy matching
                        matched_keyword: Optional[str] = None
                        matched_keyword_original: Optional[str] = None
                        
                        # First, try exact match
                        for i, keyword_cleaned in enumerate(keywords):
                            keyword_original = keywords_original[i]
                            if self._find_exact_match(keyword_cleaned, cell_value_str):
                                matched_keyword = keyword_cleaned
                                matched_keyword_original = keyword_original
                                self.debug_logger.write(f"Exact match found: '{keyword_original}' in Technical Parameters at row {row_num}, column {technical_params_col}")
                                break
                        
                        # If no exact match, try fuzzy matching
                        if matched_keyword is None:
                            best_match_score = 0.0
                            for i, keyword_cleaned in enumerate(keywords):
                                keyword_original = keywords_original[i]
                                if self._find_closest_match(keyword_cleaned, cell_value_str):
                                    # Calculate similarity score
                                    keyword_clean = self._clean_keyword(keyword_cleaned).lower()
                                    cell_clean = self._clean_keyword(cell_value_str).lower()
                                    similarity = SequenceMatcher(None, keyword_clean, cell_clean).ratio()
                                    
                                    if similarity > best_match_score:
                                        best_match_score = similarity
                                        matched_keyword = keyword_cleaned
                                        matched_keyword_original = keyword_original
                            
                            if matched_keyword is not None:
                                self.debug_logger.write(f"Fuzzy match found: '{matched_keyword_original}' (similarity: {best_match_score:.2f}) in Technical Parameters at row {row_num}, column {technical_params_col}")
                        
                        # If we found a match (exact or fuzzy), place the value
                        if matched_keyword_original is not None:
                            # Determine which option column to use
                            # Check if keyword exists in option_1_row_data or option_2_row_data using original keyword
                            self.debug_logger.write(f"Checking data sources for keyword: '{matched_keyword_original}'")
                            self.debug_logger.write(f"  - In option_1_row_data: {matched_keyword_original in option_1_row_data}")
                            self.debug_logger.write(f"  - In option_2_row_data: {matched_keyword_original in option_2_row_data}")
                            self.debug_logger.write(f"  - In base_row_data: {matched_keyword_original in base_row_data}")
                            
                            # Check Option 1 data
                            if matched_keyword_original in option_1_row_data and option_1_row_data[matched_keyword_original] is not None:
                                value_to_place = option_1_row_data[matched_keyword_original]
                                if option_1_col is not None:
                                    self.debug_logger.write(f"  - Placing Option 1 value in column {option_1_col}, row {row_num}")
                                    self._place_value_in_cell(sheet, row_num, option_1_col, matched_keyword_original, value_to_place)
                                else:
                                    self.debug_logger.write(f"  - Warning: Option 1 column is None, cannot place value")
                            
                            # Check Option 2 data (separate check, not elif, so both can be placed)
                            if matched_keyword_original in option_2_row_data and option_2_row_data[matched_keyword_original] is not None:
                                value_to_place = option_2_row_data[matched_keyword_original]
                                if option_2_col is not None:
                                    self.debug_logger.write(f"  - Placing Option 2 value in column {option_2_col}, row {row_num}")
                                    self._place_value_in_cell(sheet, row_num, option_2_col, matched_keyword_original, value_to_place)
                                else:
                                    self.debug_logger.write(f"  - Warning: Option 2 column is None, cannot place value")
                            
                            # If not in option data, try base data and use Option 1 column as default
                            if (matched_keyword_original not in option_1_row_data or option_1_row_data.get(matched_keyword_original) is None) and \
                               (matched_keyword_original not in option_2_row_data or option_2_row_data.get(matched_keyword_original) is None):
                                if matched_keyword_original in base_row_data and base_row_data[matched_keyword_original] is not None:
                                    value_to_place = base_row_data[matched_keyword_original]
                                    if option_1_col is not None:
                                        self.debug_logger.write(f"  - Placing base data value in Option 1 column {option_1_col}, row {row_num}")
                                        self._place_value_in_cell(sheet, row_num, option_1_col, matched_keyword_original, value_to_place)
                                    else:
                                        self.debug_logger.write(f"  - Warning: Option 1 column is None, cannot place base data value")
            
            self.debug_logger.write(f"Mapped data for sheet: {sheet.title}")
            
        except Exception as e:
            self.debug_logger.write(f"Error mapping data to template: {e}")
            import traceback
            error_trace = traceback.format_exc()
            self.debug_logger.write(error_trace)


