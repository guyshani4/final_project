from electronic_sheet import *
import pytest
from unittest.mock import Mock, patch

def test_is_valid_cell_name():
    spreadsheet = Spreadsheet()
    # Test valid cell names
    assert spreadsheet.is_valid_cell_name("A1") == True
    assert spreadsheet.is_valid_cell_name("B2") == True
    assert spreadsheet.is_valid_cell_name("AZ10") == True
    # Test invalid cell names
    assert spreadsheet.is_valid_cell_name("1A") == False
    assert spreadsheet.is_valid_cell_name("a10") == False
    assert spreadsheet.is_valid_cell_name("") == False

def test_set_cell_valid_name():
    spread1 = Spreadsheet()
    spread2 = Spreadsheet()
    spread1.set_cell("A1", 10)
    assert "A1" in spread1.cells
    assert isinstance(spread1.cells["A1"], Cell)
    assert spread1.cells["A1"].value == 10
    spread2.set_cell('B2', formula="A1 * 2")
    assert 'B2' in spread2.cells
    assert spread2.cells['B2'].formula == "A1 * 2"
    spread3 = Spreadsheet()
    spread3.set_cell('A3', 200)
    spread3.set_cell('A3', 300)  # Update the value
    assert spread3.cells['A3'].value == 300
    spread3.set_cell('A4', 500)
    spread3.set_cell('A4', formula="A3 + 100")  # Change to formula
    assert spread3.get_cell_value('A4') == 400


valid_cell_names = ["A1", "B2", "AA10", "Z99", "AAA100"]
invalid_cell_names = ["1A", "B-2", "AA_10", "99Z", "100AAA", "", "A!1", "A B"]

def test_get_cell_with_valid_names():
    ss = Spreadsheet()
    for name in valid_cell_names:
        # Mock a Cell and add it to the spreadsheet for each valid name
        test_cell = Cell()
        ss.cells[name] = test_cell
        assert ss.get_cell(name) is test_cell, f"Failed to retrieve the correct Cell object for {name}."

def test_get_cell_with_nonexistent_names():
    ss = Spreadsheet()
    for name in valid_cell_names:
        # Ensure that validly formatted but non-existent cell names return None
        assert ss.get_cell(name) is None, f"Should return None for a non-existent cell {name} that is validly named."

@patch.object(Spreadsheet, 'is_valid_cell_name')
def test_get_cell_with_invalid_names(mock_is_valid):
    ss = Spreadsheet()
    mock_is_valid.return_value = False  # Assume all names in the list are invalid
    for name in invalid_cell_names:
        result = ss.get_cell(name)
        assert result is None, f"Should return None for an invalid cell name {name}."
        mock_is_valid.assert_called_with(name)  # Check if is_valid_cell_name was called with the invalid name

def setup_spreadsheet():
    spreadsheet = Spreadsheet()
    spreadsheet.set_cell('A1', 10)
    spreadsheet.set_cell('A2', 20, formula="A1 * 2")
    spreadsheet.set_cell('A3', formula="A1 + A2")
    return spreadsheet

def test_get_cell_value_valid_cell():
    ss = Spreadsheet()
    # Test with a numeric value
    mock_cell_numeric = Mock(spec=Cell)
    mock_cell_numeric.calculated_value.return_value = 123
    with patch.object(ss, 'get_cell', return_value=mock_cell_numeric), \
         patch.object(ss, 'is_valid_cell_name', return_value=True):
        assert ss.get_cell_value("A1") == 123, "Failed to retrieve numeric value"

    # Test with a string value
    mock_cell_string = Mock(spec=Cell)
    mock_cell_string.calculated_value.return_value = "Hello"
    with patch.object(ss, 'get_cell', return_value=mock_cell_string):
        assert ss.get_cell_value("B2") == "Hello", "Failed to retrieve string value"


def test_get_cell_value_invalid_cell_name(capsys):
    ss = Spreadsheet()
    with patch.object(ss, 'is_valid_cell_name', return_value=False):
        result = ss.get_cell_value("InvalidCell")
        captured = capsys.readouterr()  # Capture the print output
        assert result is None, "Should return None for invalid cell names"
        assert "Invalid cell name 'InvalidCell'." in captured.out, "Expected error message not printed"

def test_get_cell_value_cell_does_not_exist(capsys):
    ss = Spreadsheet()
    # Assuming get_cell returns None for a non-existing cell and the cell name is valid
    with patch.object(ss, 'get_cell', return_value=None), \
         patch.object(ss, 'is_valid_cell_name', return_value=True):
        result = ss.get_cell_value("Z99")
        assert result is None, "Should return None for non-existing cells"
        captured = capsys.readouterr()
        assert not captured.out, "No error message should be printed for non-existing but valid cells"

def test_regular_formula_valid():
    spreadsheet = setup_spreadsheet()
    assert spreadsheet.regular_formula("A1 + A2") == 30
    assert spreadsheet.regular_formula("20 / 4") == 5

def test_regular_formula_unsupported_operation():
    ss = Spreadsheet()
    with patch.object(ss, 'get_cell_value', side_effect=[6, 3]):
        result = ss.regular_formula("A1 ^ A2")
        assert result is None, "Unsupported operation should return None"
        result = ss.regular_formula("A1A2")
        assert result is None, "Invalid formula format should return None"

def test_evaluate_formula_functions():
    spreadsheet = setup_spreadsheet()
    spreadsheet.set_cell('B1', 30)
    spreadsheet.set_cell('B2', 40)
    spreadsheet.set_cell('B3', 50)
    assert spreadsheet.evaluate_formula("SUM(B1:B3)") == 120
    assert spreadsheet.evaluate_formula("AVERAGE(B1:B3)") == 40
    assert spreadsheet.evaluate_formula("MIN(B1:B3)") == 30
    assert spreadsheet.evaluate_formula("MAX(B1:B3)") == 50

def test_regular_formula_division_by_zero():
    ss = Spreadsheet()
    with patch.object(ss, 'get_cell_value', side_effect=[6, 0]):
        result = ss.regular_formula("A1 / A2")
        assert result is None, "Division by zero should return None"

def populate_spreadsheet(spreadsheet):
    spreadsheet.set_cell('A1', 100)
    spreadsheet.set_cell('B1', 200)
    spreadsheet.set_cell('A2',None, "A1 * 2")
    spreadsheet.set_cell('B2',None, "B1 * 2")
    return spreadsheet
    # Assuming that formulas "A1 * 2" and "B1 * 2" would be evaluated to 200 and 400, respectively

def test_str_empty_spreadsheet():
    ss = Spreadsheet()
    assert str(ss) == ""

def test_str_non_empty():
    ss = Spreadsheet()
    ss = populate_spreadsheet(ss)
    expected_output_simple = "{\n  A1: 100.0,\n  B1: 200.0,\n  A2: 200.0 (Formula: A1 * 2),\n  B2: 400.0 (Formula: B1 * 2)\n}"
    assert str(ss) == expected_output_simple


def test_str_with_various_cell_values():
    ss = Spreadsheet()
    ss.set_cell('A1', "Text")
    ss.set_cell('B1',3.14159)
    ss.set_cell('C1', "True")
    expected_part_of_output = "{\n  A1: Text,\n  B1: 3.14159,\n  C1: True\n}"
    assert expected_part_of_output in str(ss)

def test_str_with_large_range():
    ss = Spreadsheet()
    for i in range(1, 11):  # Populate A1:A10 with incrementing values
        ss.set_cell(f'A{i}', i * 10)
    assert "A10: 100" in str(ss)  # Check if the last cell is correctly represented

def test_table_string_empty_spreadsheet():
    ss = Spreadsheet()
    expected_output = "The spreadsheet is empty."
    assert ss.table_string() == expected_output

def test_table_string_filled_spreadsheet():
    ss = Spreadsheet()
    ss.set_cell('A1', 100)  # Assuming set_cell takes numeric values directly
    ss.set_cell('B2', 200)  # This should match with your setup; if expecting float, consider this in expected output
    ss.set_cell('C3', "Hello")

    expected_output = (
         'A          B          C         \n'
         '-------------------------------------\n'
         '1    100.0      -          -         \n'
         '2    -          200.0      -         \n'
         '3    -          -          Hello'
    )
    # Adjust expected_output based on actual implementation details
    assert ss.table_string().strip() == expected_output.strip()

def test_check_operations():
    """
    the AVERAGE/MIN/MAX/SUM operations will work only if al the cell's values
    in the range are integers or floats. for example, if one of the values are a string,
    the operation will return None
    :return:
    """
    spread1 = Spreadsheet()
    spread1 = populate_spreadsheet(spread1)
    spread1.set_cell('C5', None, 'AVERAGE(A1:B2)')
    spread1.set_cell('C6', None, 'MIN(A1:B2)')
    spread1.set_cell('C7', None, 'MAX(A1:B2)')
    spread1.set_cell('C8', None, 'SUM(A1:B2)')
    spread1.set_cell('B3', 'Hello')
    spread1.set_cell('D5', None, 'AVERAGE(A1:B3)')
    spread1.set_cell('D6', None, 'MIN(A1:B3)')
    spread1.set_cell('D7', None, 'MAX(A1:B3)')
    spread1.set_cell('D8', None, 'SUM(A1:B3)')
    assert spread1.get_cell_value('C5') == 225.0
    assert spread1.get_cell_value('C6') == 100.0
    assert spread1.get_cell_value('C7') == 400.0
    assert spread1.get_cell_value('C8') == 900.0
    assert spread1.get_cell_value('D5') == None
    assert spread1.get_cell_value('D6') == None
    assert spread1.get_cell_value('D7') == None
    assert spread1.get_cell_value('D8') == None




