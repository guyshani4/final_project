from electronic_sheet import *
import pytest

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

def test_set_cell_invalid_name():
    spreadsheet = Spreadsheet()
    with pytest.raises(ValueError):
        spreadsheet.set_cell("1A", 10)

def test_get_cell():
    spreadsheet = Spreadsheet()
    spreadsheet.set_cell("B2", 20)
    cell = spreadsheet.get_cell("B2")
    assert cell is not None
    assert cell.value == 20

    # Test getting a cell that doesn't exist
    cell = spreadsheet.get_cell("C3")
    assert cell is None

    # Test invalid cell name
    with pytest.raises(ValueError):
        spreadsheet.get_cell("2B")

def setup_spreadsheet():
    spreadsheet = Spreadsheet()
    spreadsheet.set_cell('A1', 10)
    spreadsheet.set_cell('A2', 20, formula="A1 * 2")
    spreadsheet.set_cell('A3', formula="A1 + A2")
    return spreadsheet

def test_get_cell_value_valid():
    spreadsheet = setup_spreadsheet()
    assert spreadsheet.get_cell_value('A1') == 10
    assert spreadsheet.get_cell_value('A2') == 20  # Assuming calculated_value works correctly
    assert spreadsheet.get_cell_value('A3') == 30

def test_get_cell_value_invalid():
    spreadsheet = setup_spreadsheet()
    with pytest.raises(ValueError):
        spreadsheet.get_cell_value('Invalid')

def test_regular_formula_valid():
    spreadsheet = setup_spreadsheet()
    assert spreadsheet.regular_formula("A1 + A2") == 30
    assert spreadsheet.regular_formula("20 / 4") == 5

def test_regular_formula_invalid():
    spreadsheet = setup_spreadsheet()
    with pytest.raises(ValueError) as e:
        spreadsheet.regular_formula("A1 ** A2")
    assert str(e.value) == "Unsupported operation"

def test_evaluate_formula_functions():
    spreadsheet = setup_spreadsheet()
    spreadsheet.set_cell('B1', 30)
    spreadsheet.set_cell('B2', 40)
    spreadsheet.set_cell('B3', 50)
    assert spreadsheet.evaluate_formula("SUM(B1:B3)") == 120
    assert spreadsheet.evaluate_formula("AVERAGE(B1:B3)") == 40
    assert spreadsheet.evaluate_formula("MIN(B1:B3)") == 30
    assert spreadsheet.evaluate_formula("MAX(B1:B3)") == 50

def test_evaluate_formula_division_by_zero():
    spreadsheet = setup_spreadsheet()
    with pytest.raises(ValueError) as e:
        spreadsheet.evaluate_formula("A1 / 0")
    assert str(e.value) == "Division by zero"

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




