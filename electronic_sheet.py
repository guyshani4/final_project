import json
from typing import *
LETTERS_NUM = 26
class Cell:
    """
    Represents a single cell in a Spreadsheet.
    """
    def __init__(self, value: Optional[Any] = None, formula: Optional[str] = None) -> None:
        """
        Initializes a new Cell instance.
        :param value: The initial value of the cell.
        if nothing was inserted as value, None as default.
        """
        self.value = value
        self.formula = formula

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

    def set_value(self, value):
        self.value = value



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
        if "".join(letter_part + number_part) == cell_name:
            return True

        return False

    def __str__(self) -> str:
        """
        set a string representation of the spreadsheet.
        For cells with formulas, both the formula and its evaluated value are displayed.
        For cells without formulas, only the value is displayed.
        If the spreadsheet is empty, a message indicating that is returned.
        :return str: The formatted string representation of the spreadsheet's contents
        """
        if not self.cells:
            print("the spreadsheet is empty.")
            return ""

        cell_strings = []
        for cell_name, cell in self.cells.items():
            if cell.formula:
                # Assuming a method exists to evaluate the formula and get its value
                evaluated_value = self.evaluate_formula(cell.formula)
                cell_info = f"{cell_name}: {evaluated_value} (Formula: {cell.formula})"
            else:
                cell_info = f"{cell_name}: {cell.value}"
            cell_strings.append(cell_info)

        return "{\n  " + ",\n  ".join(cell_strings) + "\n}"

    def table_string(self) -> str:
        """
        Generates a string representation of the spreadsheet in a table format.
        This method creates a visual table where each cell is aligned in columns and rows,
        similar to a traditional spreadsheet view.
        Cells without a value or formula are represented by a dash ("-").
        The representation aims to provide a clear overview of the spreadsheet's contents,
        making it useful for debugging or displaying the current state of the spreadsheet.
        example of a print:
             A          B          C
        -----------------------------------------
        1    100        -          -
        2    -          200        -
        3    -          -          Hello

        :return str: A string representing the spreadsheet in a structured table format.
        """
        if not self.cells:
            return "The spreadsheet is empty."

        # Identify the max column and row
        max_col_index = 0
        max_row = 0
        for cell_name in self.cells.keys():
            col, row = cell_name.rstrip('0123456789'), int(cell_name.lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZ"))
            col_index = self.col_letter_to_index(col)
            max_col_index = max(max_col_index, col_index)
            max_row = max(max_row, row)

        # Generate column headers
        col_headers = [self.col_index_to_letter(i) for i in range(max_col_index + 1)]
        header = '     ' + ' '.join(f'{col: <10}' for col in col_headers)
        separator = '-' * len(header)

        # Generate the table rows
        rows = [header, separator]
        for row_num in range(1, max_row + 1):
            row_cells = [f'{row_num: <4}']
            for i in range(max_col_index + 1):
                col_letter = self.col_index_to_letter(i)
                cell_name = f"{col_letter}{row_num}"
                cell_value = self.get_cell_value(cell_name) or "-"
                cell_str = f'{str(cell_value): <10}'
                row_cells.append(cell_str)
            rows.append(' '.join(row_cells))

        return '\n'.join(rows)

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
        cell = Cell(value, formula)
        if not value:
            cell.set_value(cell.calculated_value(self))
        try:
            cell.set_value(float(value))
            cell.set_value(cell.calculated_value(self))
        except:
            pass
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
                print(f"Error: {str(err)}")
                return
        return

    def regular_formula(self, formula: str) -> Any:
        """
        regular formula stands for formulas with one/two cells and regular operation (*,/,+,-)
        for example: "A1/A2", "A3*4"
        :param formula: a string with the formula
        :return: the answer of the formula. if the formula doesnt meet the string requirements, None.
        """
        parts = formula.split()
        if len(parts) == 3:
            side1, operation, side2 = parts
            # Try converting the first operand to a number, if it fails, treat it as a cell reference
            try:
                value1 = float(side1)
            except ValueError:
                value1 = self.get_cell_value(side1)
            # Repeat the process for the second operand
            try:
                value2 = float(side2)
            except ValueError:
                value2 = self.get_cell_value(side2)
            if value2 == None or value1 == None:
                print("formula cannot be calculated.. value is not exist")
                return
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


    def evaluate_formula(self, formula: str) -> Any:
        """
        calculates a given formula represented as a string.
        :param formula: A string formula, for example: "A1 + B1", "AVERAGE(A1:B2)", "MIN(B1:Z3)"
        :return: The result of the formula calculation.
        :raises ValueError: If the formula is invalid or contains unknown operations.
        """
        if formula.startswith("AVERAGE"):
            try:
                cells = self.valid_cells_index(formula)
                return self.calculate_average(cells[0], cells[1])
            except Exception as err:
                print(f"Error: {str(err)}")
                return
        if formula.startswith("SUM"):
            try:
                cells = self.valid_cells_index(formula)
                return self.calculate_sum(cells[0], cells[1])
            except Exception as err:
                print(f"Error: {str(err)}")
                return
        if formula.startswith("MIN"):
            try:
                cells = self.valid_cells_index(formula)
                return self.find_min(cells[0], cells[1])
            except Exception as err:
                print(f"Error: {str(err)}")
                return
        if formula.startswith("MAX"):
            try:
                cells = self.valid_cells_index(formula)
                return self.find_max(cells[0], cells[1])
            except Exception as err:
                print(f"Error: {str(err)}")
                return
        return self.regular_formula(formula)



    def cells_values_list(self, start: str, end: str) -> Any:
        """
        Retrieves a list with all the values in a given range
         :param start: The starting cell name of the range.
        :param end: The ending cell name of the range.
        :return: list of values of all the cells in the range
        """
        cell_names = self.get_range_cells(start, end)
        # Retrieve values and filtering out cells that do not exist or have None as their value.
        values = []
        for name in cell_names:
            if self.get_cell_value(name):
                if isinstance(self.get_cell_value(name), int) or isinstance(self.get_cell_value(name), float):
                    values.append(float(self.get_cell_value(name)))
                else:
                    return
        return values

    def find_min(self, start: str, end: str) -> Any:
        """
        finds the minimum cell value in a specific range that was given
        :param start: The starting cell name of the range.
        :param end: The ending cell name of the range.
        :return: float: the minimum value in the range.
        """
        try:
            cell_names = self.cells_values_list(start, end)
            return min(cell_names)
        except Exception as err:
            print(f"Error: {str(err)}")
            return


    def find_max(self, start: str, end: str) -> Any:
        """
        finds the maximum cell value in a specific range that was given
        :param start: The starting cell name of the range.
        :param end: The ending cell name of the range.
        :return: float: the maximum value in the range.
        """
        cell_names = self.cells_values_list(start, end)
        try:
            return max(cell_names)
        except Exception as err:
            print(f"Error: {str(err)}")
            return

    def calculate_sum(self, start: str, end: str) -> Any:
        """
        calculates the sum of cells values in a specific range that was given
        :param start: The starting cell name of the range.
        :param end: The ending cell name of the range.
        :return: float: the sum of all the values in the range.
        """
        values_list = self.cells_values_list(start, end)
        if values_list:
            try:
                return sum(values_list)
            except Exception as err:
                return f"Error: {str(err)}"
        return

    def valid_cells_index(self, formula: str) -> Tuple[str, str]:
        """
        gets a string with specific formula.
        Retrieves the specific 2 cells that in the formula.
        :param formula: a string of an operation and 2 specific cells to check the range between them.
        :return: the 2 cells that in the formula.
        for example:  ("AVERAGE(A1:B2)" -> ("A1", "B2")
        """
        for index in range(len(formula)):
            if formula[index] == "(":
                formula = formula[index+1:]
                break
        formula = formula.replace(")", "")
        cell_list = formula.split(":")
        if len(cell_list) != 2:
            raise ValueError("index can not be calculated")
        return cell_list[0], cell_list[1]

    def calculate_average(self, start: str, end: str) -> Any:
        """
        Calculates the average value of cells in a specified range, ignoring cells with no value.
        :param start: The starting cell name of the range.
        :param end: The ending cell name of the range.
        :return None if any cell was not found, else,
        The average value of the cells in the range, ignoring cells without a value.
        """
        values_list = self.cells_values_list(start, end)
        if values_list:
            try:
                return sum(values_list)/len(values_list)
            except Exception as err:
                print(f"Error: {str(err)}")
        return

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

    def col_letter_to_index(self, col: str) -> int:
        """
        Converts a column letter (LIKE "A") to an integer index (for example, "A" -> "0").
        for example: "A" -> 0, "B" -> 1, "AA" - 26
        :param col: a col as a string
        :return the col's index as integer
        """
        index = 0
        for char in col:
            index = index * LETTERS_NUM + (ord(char.upper()) - ord('A') + 1)
        return index - 1


    def col_index_to_letter(self, index: int) -> str:
        """
        Converts an integer index to a column letter (for example, 0 -> "A").
        for example: 0 -> "A", 1 -> "B", 26 - "AA"
        :param index: the col's index as integer
        :return the col's string
        """
        col = ''
        while index >= 0:
            col = chr(index % LETTERS_NUM + ord('A')) + col
            index = index // LETTERS_NUM - 1
        return col

    def get_range_cells(self, start: str, end: str) -> List[str]:
        """
        creates a list of cell names in a range that can span multiple columns and rows.
        for example: (A1, B2) -> ["A1", "A2", "B1", "B2"]

        :param start: The starting cell name of the range.
        :param end: The ending cell name of the range.
        :return List[str]: A list of cell names within the specified range.
        """
        # Separates the letters and digits in each cell
        start_col = [letter for letter in start if letter.isalpha()]
        end_col = [letter for letter in end if letter.isalpha()]
        start_row = [digit for digit in start if digit.isdigit()]
        end_row = [digit for digit in end if digit.isdigit()]
        # Translates the col index into an integer
        start_col_index = self.col_letter_to_index(start_col[0])
        end_col_index = self.col_letter_to_index(end_col[0])

        if start_col_index > end_col_index or start_row > end_row:
            raise ValueError(f"Invalid cells range. '{end}' comes after '{start}")
        # Creates the list of all the indexes as strings.
        cells = []
        for col in range(start_col_index, end_col_index + 1):
            for row in range(int(int(start_row[0])), int(int(end_row[0])) + 1):
                col_letter = self.col_index_to_letter(col)
                cell_name = f"{col_letter}{row}"
                cells.append(cell_name)
        return cells

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

"""
if __name__ == '__main__':
    spreadsheet = Spreadsheet()
    spreadsheet.set_cell('A1', 100)
    spreadsheet.set_cell('B1', 200)
    spreadsheet.set_cell('A2',None, "A1 * 2")
    spreadsheet.set_cell('B2',None, "B1 * 2")
    spreadsheet.set_cell('A3', "Hello")
    print(spreadsheet)
    spreadsheet.set_cell('B3', None, "MAX(A1:B2)")
    print(spreadsheet.get_cell_value('B3'))
"""