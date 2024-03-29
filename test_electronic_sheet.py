from workbook import *
import matplotlib.pyplot as plt
from unittest.mock import patch


def test_set_value():
    cell = Cell()
    cell.set_value(10)
    assert cell.value == 10
    cell.set_value(20)
    assert cell.value == 20
    cell.set_value("Hello")
    assert cell.value == "Hello"
    cell.set_value(None)
    assert cell.value is None


def test_add_remove_dependent():
    cell = Cell()
    cell.add_dependent('A1')
    cell.add_dependent('B1')
    cell.add_dependent('C1')
    assert len(cell.dependents) == 3
    cell.remove_dependent('B1')
    assert len(cell.dependents) == 2
    assert 'B1' not in cell.dependents


def test_calculated_value():
    cell = Cell()
    spreadsheet = Spreadsheet()
    cell.set_value(10)
    assert cell.calculated_value(spreadsheet) == 10
    cell.set_value(20)
    assert cell.calculated_value(spreadsheet) == 20
    cell.set_value("Hello")
    assert cell.calculated_value(spreadsheet) == "Hello"
    cell.set_value(None)
    assert cell.calculated_value(spreadsheet) is None


def test_is_valid_cell_name():
    spreadsheet = Spreadsheet()
    assert spreadsheet.is_valid_cell_name('A1') == True
    assert spreadsheet.is_valid_cell_name('B2') == True
    assert spreadsheet.is_valid_cell_name('AZ10') == True
    assert spreadsheet.is_valid_cell_name('1A') == False
    assert spreadsheet.is_valid_cell_name('a10') == False
    assert spreadsheet.is_valid_cell_name('') == False
    assert spreadsheet.is_valid_cell_name('A 1') == False
    assert spreadsheet.is_valid_cell_name('A-1') == False


def test_set_get_cell():
    spreadsheet = Spreadsheet()
    spreadsheet.set_cell('A1', 10)
    assert spreadsheet.get_cell_value('A1') == 10
    spreadsheet.set_cell('B1', 20)
    assert spreadsheet.get_cell_value('B1') == 20
    spreadsheet.set_cell('A1', 30)
    assert spreadsheet.get_cell_value('A1') == 30
    spreadsheet.set_cell('C1', "Hello")
    assert spreadsheet.get_cell_value('C1') == "Hello"
    spreadsheet.set_cell('D1', None)
    assert spreadsheet.get_cell_value('D1') is None


def test_set_cell():
    spreadsheet = Spreadsheet()

    # Test setting a cell with a valid name and a value
    spreadsheet.set_cell('A1', 10)
    assert spreadsheet.get_cell_value('A1') == 10

    # Test setting a cell with a valid name and a formula
    spreadsheet.set_cell('B1', formula='A1+10')
    assert spreadsheet.get_cell_value('B1') == 20

    # Test setting a cell with an invalid name
    try:
        spreadsheet.set_cell('1A', 10)
    except Exception as e:
        assert str(e) == "Invalid cell name '1A'. Cell names must be in the format 'A1', 'B2', 'AZ10' etc."

    # Test setting a cell with a valid name and a formula that references an invalid cell
    try:
        spreadsheet.set_cell('C1', formula='1A+10')
    except Exception as e:
        assert str(e) == "Invalid cell name '1A'. Cell names must be in the format 'A1', 'B2', 'AZ10' etc."


def test_remove_cell():
    spreadsheet = Spreadsheet()
    spreadsheet.set_cell('A1', 10)
    spreadsheet.remove_cell('A1')
    assert spreadsheet.get_cell('A1').value is None
    spreadsheet1 = Spreadsheet()
    # Set a cell with a formula
    spreadsheet1.set_cell('A1', formula='A2+A3')
    assert spreadsheet1.get_cell('A1').formula == 'A2+A3'
    spreadsheet1.remove_cell('A1')
    assert spreadsheet1.get_cell('A1') is None


def test_max_row():
    spreadsheet = Spreadsheet()
    spreadsheet.set_cell('A1', 10)
    spreadsheet.set_cell('A2', 20)
    spreadsheet.set_cell('A3', 30)
    assert spreadsheet.max_row() == 3


def test_max_col_index():
    spreadsheet = Spreadsheet()
    spreadsheet.set_cell('A1', 10)
    spreadsheet.set_cell('B1', 20)
    spreadsheet.set_cell('C1', 30)
    assert spreadsheet.max_col_index() == 2


def test_set_cell_formula():
    spreadsheet = Spreadsheet()

    # Test setting a cell with a valid name and a formula
    spreadsheet.set_cell_formula(spreadsheet.get_cell('A1'), 'A1', 'B1+10')
    assert spreadsheet.get_cell_value('A1') == None  # B1 is not set yet

    # Set value of B1 and recheck A1
    spreadsheet.set_cell('B1', 10)
    assert spreadsheet.get_cell_value('A1') == 10

    # Test setting a cell with a valid name and a formula that references multiple cells
    spreadsheet.set_cell_formula(spreadsheet.get_cell('A2'), 'A2', 'A1+B1')
    assert spreadsheet.get_cell_value('A2') == 20

    # Test setting a cell with an invalid name
    try:
        spreadsheet.set_cell_formula(spreadsheet.get_cell('1A'), '1A', 'A1+10')
    except Exception as e:
        assert str(e) == "Invalid cell name '1A'. Cell names must be in the format 'A1', 'B2', 'AZ10' etc."

    # Test setting a cell with a valid name and a formula that references an invalid cell
    try:
        spreadsheet.set_cell_formula(spreadsheet.get_cell('A3'), 'A3', '1A+10')
    except Exception as e:
        assert str(e) == "Invalid cell name '1A'. Cell names must be in the format 'A1', 'B2', 'AZ10' etc."


def test_workbook():
    workbook = Workbook()
    assert workbook.get_sheet('Sheet2') is None
    workbook.add_sheet('Sheet2')
    assert workbook.get_sheet('Sheet2') is not None
    workbook.remove_sheet('Sheet2')
    assert workbook.get_sheet('Sheet2') is None
    workbook.add_sheet('Sheet2')
    assert workbook.get_sheet('Sheet2') is not None
    workbook.rename_sheet('Sheet2', 'Sheet3')
    assert workbook.get_sheet('Sheet2') is None
    assert workbook.get_sheet('Sheet3') is not None


def test_str():
    spreadsheet = Spreadsheet()
    spreadsheet.set_cell('A1', 10)
    spreadsheet.set_cell('B1', 20)
    spreadsheet.set_cell('A2', 30)
    assert str(spreadsheet) == ('     A          B         \n'
                                '--------------------------\n'
                                '1    10.0       20.0      \n'
                                '2    30.0       -         ')


def test_get_range_cells_basic_range():
    ss = Spreadsheet()
    expected = ["A1", "A2", "B1", "B2"]
    assert ss.get_range_cells("A1", "B2") == expected
    assert ss.get_range_cells("B1", "A2") is None


# Test for a single-cell range
def test_get_range_cells_single_cell():
    ss = Spreadsheet()
    expected = ["A1"]
    assert ss.get_range_cells("A1", "A1") == expected


# Test for a larger range
def test_get_range_cells_larger_range():
    ss = Spreadsheet()
    expected = ["A1", "A2", "A3", "B1", "B2", "B3", "C1", "C2", "C3"]
    assert ss.get_range_cells("A1", "C3") == expected


# Test for wide range
def test_get_range_cells_wide_range():
    ss = Spreadsheet()
    expected = ["A1", "A2", "A3", "A4", "A5",
                "B1", "B2", "B3", "B4", "B5",
                "C1", "C2", "C3", "C4", "C5",
                "D1", "D2", "D3", "D4", "D5",
                "E1", "E2", "E3", "E4", "E5"]
    assert ss.get_range_cells("A1", "E5") == expected


def test_calculate_average():
    spreadsheet = Spreadsheet()

    # Test calculating average in a range with valid cell names
    spreadsheet.set_cell('A1', 10)
    spreadsheet.set_cell('A2', 20)
    spreadsheet.set_cell('A3', 30)
    assert spreadsheet.calculate_average('A1', 'A3') == 20

    # Test calculating average in a range where some cells are not set
    assert spreadsheet.calculate_average('A1', 'A4') == 20

    # Test calculating average in a range where all cells are not set
    assert spreadsheet.calculate_average('A4', 'A5') is None

    # Test calculating average in a range with invalid cell names
    try:
        spreadsheet.calculate_average('1A', '2A')
    except Exception as e:
        assert str(e) == "Invalid cell name '1A'. Cell names must be in the format 'A1', 'B2', 'AZ10' etc."


def test_valid_cells_index():
    spreadsheet = Spreadsheet()

    # Test a formula with a valid range of cells
    start, end = spreadsheet.valid_cells_index('AVERAGE(A1:B2)')
    assert start == 'A1'
    assert end == 'B2'

    # Test a formula with an invalid range of cells
    try:
        spreadsheet.valid_cells_index('AVERAGE(A1:2B)')
    except Exception as e:
        assert str(e) == "the formula does not fit the requirements"

    # Test a formula with a valid range of cells but in reverse order
    start, end = spreadsheet.valid_cells_index('AVERAGE(B2:A1)')
    assert start == 'B2'
    assert end == 'A1'

    # Test a formula with a valid range of cells but with different columns
    start, end = spreadsheet.valid_cells_index('AVERAGE(A1:B1)')
    assert start == 'A1'
    assert end == 'B1'


def test_calculate_sum():
    spreadsheet = Spreadsheet()
    spreadsheet.set_cell('A1', 10)
    spreadsheet.set_cell('B1', 20)
    spreadsheet.set_cell('A2', 30)
    assert spreadsheet.calculate_sum('A1', 'B2') == 60
    spreadsheet = Spreadsheet()
    # Test calculating sum in a range with valid cell names
    spreadsheet.set_cell('A1', 10)
    spreadsheet.set_cell('A2', 20)
    spreadsheet.set_cell('A3', 30)
    assert spreadsheet.calculate_sum('A1', 'A3') == 60

    # Test calculating sum in a range where some cells are not set
    assert spreadsheet.calculate_sum('A1', 'A4') == 60

    # Test calculating sum in a range where all cells are not set
    assert spreadsheet.calculate_sum('A4', 'A5') == 0

    # Test calculating sum in a range with invalid cell names
    try:
        spreadsheet.calculate_sum('1A', '2A')
    except Exception as e:
        assert str(e) == ("Invalid cell name '1A'. "
                          "Cell names must be in the format 'A1', 'B2', 'AZ10' etc.")


def test_col_index_to_letter():
    ss = Spreadsheet()

    # Test single-letter columns
    assert ss.col_index_to_letter(0) == "A", "Index 0 should correspond to A"
    assert ss.col_index_to_letter(1) == "B", "Index 1 should correspond to B"
    assert ss.col_index_to_letter(25) == "Z", "Index 25 should correspond to Z"

    # Test double-letter columns (after Z, which is 25)
    assert ss.col_index_to_letter(26) == "AA", "Index 26 should correspond to AA"
    assert ss.col_index_to_letter(27) == "AB", "Index 27 should correspond to AB"
    assert ss.col_index_to_letter(51) == "AZ", "Index 51 should correspond to AZ"
    assert ss.col_index_to_letter(52) == "BA", "Index 52 should correspond to BA"
    assert ss.col_index_to_letter(701) == "ZZ", "Index 701 should correspond to ZZ"
    assert ss.col_index_to_letter(702) == "AAA", "Index 702 should correspond to AAA"


def test_get_cell_formula():
    spreadsheet = Spreadsheet()

    # Test setting a cell with a valid name and a formula that references another cell
    spreadsheet.set_cell('B1', 10)
    spreadsheet.set_cell_formula(spreadsheet.get_cell('A1'), 'A1', 'B1+10')
    assert spreadsheet.get_cell_value('A1') == 20

    # Test setting a cell with a valid name and a formula that references multiple cells
    spreadsheet.set_cell('B2', 20)
    spreadsheet.set_cell_formula(spreadsheet.get_cell('A2'), 'A2', 'A1+B2')
    assert spreadsheet.get_cell_value('A2') == 40

    # Test setting a cell with a valid name and a formula that references an invalid cell
    try:
        spreadsheet.set_cell_formula(spreadsheet.get_cell('A3'), 'A3', '1A+10')
    except Exception as e:
        assert str(e) == "Invalid cell name '1A'. Cell names must be in the format 'A1', 'B2', 'AZ10' etc."

    # Test setting a cell with a valid name and an invalid formula
    try:
        spreadsheet.set_cell_formula(spreadsheet.get_cell('A4'), 'A4', 'A1+1A')
    except Exception as e:
        assert str(e) == "Invalid cell name '1A'. Cell names must be in the format 'A1', 'B2', 'AZ10' etc."

    # Test setting a cell with an invalid name and a valid formula
    try:
        spreadsheet.set_cell_formula(spreadsheet.get_cell('1A'), '1A', 'A1+10')
    except Exception as e:
        assert str(e) == "Invalid cell name '1A'. Cell names must be in the format 'A1', 'B2', 'AZ10' etc."


def test_regular_formula():
    spreadsheet = Spreadsheet()

    # Test a formula with a valid cell name and a valid operation
    spreadsheet.set_cell('A1', 10)
    spreadsheet.set_cell('B1', 20)
    assert spreadsheet.regular_formula('A1+B1') == 30

    # Test a formula with a valid cell name and an invalid operation
    try:
        spreadsheet.regular_formula('A1^B1')
    except Exception as e:
        assert str(e) == "Invalid formula format."

    # Test a formula with an invalid cell name and a valid operation
    try:
        spreadsheet.regular_formula('1A+B1')
    except Exception as e:
        assert str(e) == "Invalid cell name '1A'. Cell names must be in the format 'A1', 'B2', 'AZ10' etc."

    # Test a formula with a valid cell name and a division by zero operation
    spreadsheet.set_cell('B2', 0)
    try:
        spreadsheet.regular_formula('A1/B2')
    except Exception as e:
        assert str(e) == "Error: Division by zero."

    # Test a formula with a valid cell name and a valid operation with multiple cells
    spreadsheet.set_cell('A2', 30)
    assert spreadsheet.regular_formula('A1+A2+B1') == 60


def test_find_min():
    spreadsheet = Spreadsheet()

    # Test finding minimum in a range with valid cell names
    spreadsheet.set_cell('A1', 10)
    spreadsheet.set_cell('A2', 20)
    spreadsheet.set_cell('A3', 30)
    assert spreadsheet.find_min('A1', 'A3') == 10

    # Test finding minimum in a range where some cells are not set
    assert spreadsheet.find_min('A1', 'A4') == 10

    # Test finding minimum in a range where all cells are not set
    assert spreadsheet.find_min('A4', 'A5') is None

    # Test finding minimum in a range with invalid cell names
    try:
        spreadsheet.find_min('1A', '2A')
    except Exception as e:
        assert str(e) == ("Invalid cell name '1A'. "
                          "Cell names must be in the format 'A1', 'B2', 'AZ10' etc.")


def test_find_max():
    spreadsheet = Spreadsheet()

    # Test finding maximum in a range with valid cell names
    spreadsheet.set_cell('A1', 10)
    spreadsheet.set_cell('A2', 20)
    spreadsheet.set_cell('A3', 30)
    assert spreadsheet.find_max('A1', 'A3') == 30

    # Test finding maximum in a range where some cells are not set
    assert spreadsheet.find_max('A1', 'A4') == 30

    # Test finding maximum in a range where all cells are not set
    assert spreadsheet.find_max('A4', 'A5') is None

    # Test finding maximum in a range with invalid cell names
    try:
        spreadsheet.find_max('1A', 'A2')
    except Exception as e:
        assert str(e) == ("Invalid cell name '1A'. "
                          "Cell names must be in the format 'A1', 'B2', 'AZ10' etc.")


def test_create_graph():
    spreadsheet = Spreadsheet()

    # Set some cells for the graph
    spreadsheet.set_cell('A1', 10)
    spreadsheet.set_cell('A2', 20)
    spreadsheet.set_cell('A3', 30)
    spreadsheet.set_cell('B1', 'Data1')
    spreadsheet.set_cell('B2', 'Data2')
    spreadsheet.set_cell('B3', 'Data3')

    # Test creating a bar graph with valid ranges
    with patch.object(plt, 'show'):
        spreadsheet.create_graph('bar', 'A1:A3', 'B1:B3')

    # Test creating a pie graph with valid ranges
    with patch.object(plt, 'show'):
        spreadsheet.create_graph('pie', 'A1:A3', 'B1:B3')

    # Test creating a graph with an invalid graph type
    try:
        spreadsheet.create_graph('invalid', 'A1:A3', 'B1:B3')
    except Exception as e:
        assert str(e) == "Invalid graph type: invalid"

    # Test creating a graph with invalid ranges
    try:
        spreadsheet.create_graph('bar', 'A1:2A', 'B1:B3')
    except Exception as e:
        assert str(e) == "Invalid cells range. '2A' comes after 'A1'"

    try:
        spreadsheet.create_graph('bar', 'A1:A3', 'B1:2B')
    except Exception as e:
        assert str(e) == "Invalid cells range. '2B' comes after 'B1'"
