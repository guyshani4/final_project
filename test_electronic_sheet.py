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
    spreadsheet = Spreadsheet()
    spreadsheet.set_cell("A1", 10)
    assert "A1" in spreadsheet.cells
    assert isinstance(spreadsheet.cells["A1"], Cell)
    assert spreadsheet.cells["A1"].value == 10

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
    assert ss.__str__() == ""

def test_str_simple_version_non_empty():
    ss = Spreadsheet()
    ss = populate_spreadsheet(ss)
    expected_output_simple = "{\n  A1: 100,\n  B1: 200,\n  A2: 200.0 (Formula: A1 * 2),\n  B2: 400.0 (Formula: B1 * 2)\n}"
    assert str(ss) == expected_output_simple

def test_str_detailed_version_non_empty():
    ss = Spreadsheet()
    populate_spreadsheet(ss)
    expected_output_detailed = "{\n  A1: 100,\n  B1: 200,\n  A2: 200.0 (Formula: A1 * 2),\n  B2: 400.0 (Formula: B1 * 2)\n}"
    assert str(ss) == expected_output_detailed

def test_str_with_various_cell_values():
    ss = Spreadsheet()
    ss.set_cell('A1', "Text")
    ss.set_cell('B1',3.14159)
    ss.set_cell('C1', True)
    expected_part_of_output = "{\n  A1: Text,\n  B1: 3.14159,\n  C1: True\n}"
    assert expected_part_of_output in str(ss)

def test_str_with_large_range():
    ss = Spreadsheet()
    for i in range(1, 11):  # Populate A1:A10 with incrementing values
        ss.set_cell(f'A{i}', i * 10)
    assert "A10: 100" in str(ss)  # Check if the last cell is correctly represented