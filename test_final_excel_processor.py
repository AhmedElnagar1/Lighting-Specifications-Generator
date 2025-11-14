"""Unit tests for final_excel_processor module."""
from openpyxl import Workbook
from final_excel_processor import add_hyperlink_to_cell


def test_add_hyperlink_to_cell() -> None:
    """
    Test the add_hyperlink_to_cell function with valid URL and non-URL values.
    """
    # Test 1: Valid URL should create hyperlink with "Link" as display text
    wb = Workbook()
    ws = wb.active
    add_hyperlink_to_cell(ws, row=1, column=1, link_value="https://example.com")
    
    cell = ws.cell(row=1, column=1)
    assert cell.value == "Link", "Cell value should be 'Link'"
    assert cell.hyperlink == "https://example.com", "Cell should have hyperlink set"
    assert cell.font.underline == "single", "Cell should be underlined"
    assert cell.font.color is not None, "Cell should have a color set"
    
    # Test 2: Non-URL value should be set as plain text
    add_hyperlink_to_cell(ws, row=2, column=1, link_value="Not a URL")
    
    cell2 = ws.cell(row=2, column=1)
    assert cell2.value == "Not a URL", "Cell value should be the original text"
    assert cell2.hyperlink is None, "Cell should not have hyperlink"
    
    wb.close()


if __name__ == "__main__":
    test_add_hyperlink_to_cell()
    print("All tests passed!")

