import json
from typing import *

class Cell:
    """
    Represents a single cell in a Spreadsheet.
    """
    def __init__(self, value: Optional[Any] = None) -> None:
        """
        Initializes a new Cell instance.
        :param value: The initial value of the cell.
        if nothing was inserted as value, None as default.
        """
        self.value = value
        self.formula = None

    def calculated_value(self, spreadsheet: 'Spreadsheet') -> Any:
        """
        updates the cell's value. If the cell contains a formula, insert the formula,
        otherwise returns the cell's current value.
        :param spreadsheet: The Spreadsheet object containing this cell.
        :return: The updated value of the cell.
        """
        if self.formula:
            return spreadsheet.evaluate_formula(self.formula)
        return self.value


class Spreadsheet:
    """
    Represents a spreadsheet page, similar to an Excel page.
    """
    def __init__(self) -> None:
        """
        Initializes a new Spreadsheet instance with an empty dictionary of cells.
        """
        self.cells = {}

    def is_valid_cell_name(self, cell_name: str) -> bool:
        """
        Check if the provided cell name follows the Excel format.

        The function validates that the cell name consists of one or more uppercase letters
        followed by one or more digits - a common format for cell names in spreadsheet
        applications like Excel (for example, "A1", "B2", "AZ10").

        Parameters: cell_name: The cell name to validate.

        Returns: bool: True if the cell name is valid, False otherwise.
        """
        if not cell_name:
            return False

        # Check if the first part of the string is a string of uppercase letters
        letter_part = [letter for letter in cell_name if letter.isalpha()]
        if not letter_part or not all(letter.isupper() for letter in letter_part):
            return False

        # Check if the second part of the string consists of digits
        number_part = [digit for digit in cell_name if digit.isdigit()]
        if not number_part:
            return False

        # Check if concatenating the two parts gives the original string
        if ''.join(letter_part + number_part) == cell_name:
            return True

        return False

    def set_cell(self, cell_name: str, value: Optional[Any] = None, formula: Optional[str] = None) -> None:
        """
        Sets or updates a cell's value and/or formula.
        :param cell_name: A string identifier for the cell (for example: "A1").
        :param value: The value to set in the cell.
        :param formula: An optional formula for the cell.
        If provided, the cell's value will be determined by this formula.
        """
        if not self.is_valid_cell_name(cell_name):
            raise ValueError(f"Invalid cell name '{cell_name}'."
                             f"Cell names must be in the format 'A1', 'B2', 'AZ10' etc.")
        cell = Cell(value)
        cell.formula = formula
        self.cells[cell_name] = cell

    def get_cell(self, cell_name: str) -> Optional[Cell]:
        """
        retrieves a Cell object from the cell's dictionary.
        :param cell_name: The name of the cell to retrieve.
        :return: The Cell object if found, None otherwise.
        """
        if not self.is_valid_cell_name(cell_name):
            raise ValueError(f"Invalid cell name '{cell_name}'. Cell names must be in the format 'A1', 'B2', etc.")
        if cell_name in self.cells:
            return self.cells[cell_name]
        else:
            return None

    def get_cell_value(self, cell_name: str) -> Any:
        """
        Retrieves the evaluated value of a cell.
        :param cell_name: The name of the cell to evaluate.
        :return: The evaluated value of the cell or an error message if the cell does not exist.
        """
        if not self.is_valid_cell_name(cell_name):
            raise ValueError(f"Invalid cell name '{cell_name}'. Cell names must be in the format 'A1', 'B2', etc.")
        cell = self.get_cell(cell_name)
        if cell:
            try:
                return cell.calculated_value(self)
            except Exception as err:
                return f"Error: {str(err)}"
        return "Cell does not exist"

    def evaluate_formula(self, formula: str) -> Any:
        """
        calculates a given formula represented as a string.
        :param formula: A string formula, for example: "A1 + B1".
        :return: The result of the formula calculation.
        :raises ValueError: If the formula is invalid or contains unknown operations.
        """
        parts = formula.split()
        if len(parts) == 3:
            side1, operation, side2 = parts
            # Try converting the first operand to a number, if it fails, treat it as a cell reference
            try:
                value1 = float(side1)
            except ValueError:
                value1 = self.get_raw_value(side1)
            # Repeat the process for the second operand
            try:
                value2 = float(side2)
            except ValueError:
                value2 = self.get_raw_value(side2)

            # checks the operation
            if operation == '+':
                return value1 + value2
            elif operation == '-':
                return value1 - value2
            elif operation == '*':
                return value1 * value2
            elif operation == '/':
                if value2 != 0:
                    return value1 / value2
                else:
                    raise ValueError("Division by zero")
            else:
                raise ValueError("Unsupported operation")
        else:
            raise ValueError("Invalid formula format")

    def get_raw_value(self, cell_name: str) -> Any:
        """
        Retrieves the raw value of a cell without evaluating its formula.
        :param cell_name: The name of the cell to retrieve the value from.
        :return: The value of the cell.
        :raises ValueError: If the cell is empty, has an invalid value, or does not exist.
        """
        if not self.is_valid_cell_name(cell_name):
            raise ValueError(f"Invalid cell name '{cell_name}'. Cell names must be in the format 'A1', 'B2', etc.")
        cell = self.get_cell(cell_name)
        if cell:
            if cell.value is not None:
                return cell.value
            else:
                raise ValueError(f"Cell {cell_name} is empty or has an invalid value.")
        else:
            raise ValueError(f"Cell {cell_name} does not exist.")

    def save_as(self, filename: str) -> None:
        """
        Saves the current state of the spreadsheet to a file in JSON format.
        :param filename: The name of the file to save the spreadsheet to.
        """
        data_as_dict = {name: {'value': cell.value, 'formula': cell.formula} for name, cell in self.cells.items()}
        with open(filename, 'w') as f:
            json.dump(data_as_dict, f)

    def load(self, filename: str) -> None:
        """
        Loads a spreadsheet from a file saved in JSON format.
        it loads the json to a Spreadsheet file again, initializing every cell again.
        :param filename: The name of the file to load the spreadsheet from.
        """
        with open(filename, 'r') as f:
            loaded_dict = json.load(f)
        for name, data in loaded_dict.items():
            self.set_cell(name, data['value'], data['formula'])