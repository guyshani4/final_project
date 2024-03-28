from electronic_sheet import *
from workbook import *
import unittest
import pytest
from unittest.mock import Mock, patch


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


def test_remove_cell():
    spreadsheet = Spreadsheet()
    spreadsheet.set_cell('A1', 10)
    spreadsheet.remove_cell('A1')
    assert spreadsheet.get_cell('A1').value is None


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
    spreadsheet.set_cell('A1', 10)
    spreadsheet.set_cell('B1', 20)
    spreadsheet.set_cell('A2', 30)
    assert spreadsheet.calculate_average('A1', 'B2') == 20

def test_calculate_sum():
    spreadsheet = Spreadsheet()
    spreadsheet.set_cell('A1', 10)
    spreadsheet.set_cell('B1', 20)
    spreadsheet.set_cell('A2', 30)
    assert spreadsheet.calculate_sum('A1', 'B2') == 60

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

