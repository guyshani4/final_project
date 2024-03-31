import re
import math
from typing import *
import matplotlib.pyplot as plt  # type: ignore

LETTERS_NUM = 26
ALL_LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
ALL_DIGITS = "0123456789"
GRAPH_ERROR = "Invalid range. Please use the format 'graph [type] [range1] [range2]'.\n" \
              "For example: 'graph line A1:A10 B1:B10'\n" \
              "Supported graph types: 'bar', 'pie'\n" \
              "the first range should be the x-axis - represent topics\n" \
              "the second range should be the y-axis - represent values.\n" \
              "the start/end range of both axis should be lengths equal."


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
        # Cells that depend on this cell
        self.dependents: Set[str] = set()

    def add_dependent(self, cell_name: str) -> None:
        """
        Adds a cell to the list of cells that depend on this cell.

        :param cell_name: The name of the cell to add.
        """
        self.dependents.add(cell_name)

    def remove_dependent(self, cell_name: str) -> None:
        """
        Removes a cell from the list of cells that depend on this cell.
        :param cell_name: The name of the cell to remove.
        """

        if cell_name in self.dependents:
            self.dependents.remove(cell_name)

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

    def set_value(self, value: Any) -> None:
        """
       Sets the value of the cell.

       :param value: The value to set.
       """

        self.value = value

    def to_dict(self) -> Dict[str, Any]:
        """
        Converts the cell to a dictionary.

        :return: A dictionary representation of the cell.
        """
        return {
            'value': self.value,
            'formula': self.formula,
            'dependents': list(self.dependents)
        }

    def update_dependents(self, dependents: List[str]) -> None:
        """
        Updates the list of cells that depend on this cell.

        :param dependents: The new list of dependents.
        """
        self.dependents = set(dependents)


class Spreadsheet:
    """
    Represents a spreadsheet page, similar to an Excel page.
    """

    def __init__(self, sheet_name=None) -> None:
        """
        Initializes a new Spreadsheet instance with an empty dictionary of cells.
        """
        self.cells: Dict[str, Cell] = {}
        self.name = sheet_name

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
        if not number_part or number_part == ['0']:
            return False

        # Check if concatenating the two parts gives the original string
        if "".join(letter_part + number_part) == cell_name:
            return True

        return False

    def __str__(self) -> Any:
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
            try:
                col, row = cell_name.rstrip(ALL_DIGITS), int(cell_name.lstrip(ALL_LETTERS))
                col_index = self.col_letter_to_index(col)
                max_col_index = max(max_col_index, col_index)
                max_row = max(max_row, row)
            except:
                continue

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
                cell_value = self.get_cell_value(cell_name)
                cell_value = "-" if cell_value is None else cell_value
                cell_str = f'{str(cell_value): <10}'
                row_cells.append(cell_str)
            rows.append(' '.join(row_cells))

        return '\n'.join(rows)

    def set_cell(self, cell_name: str, value: Optional[Any] = None, formula: Optional[str] = None) -> None:
        """
        Sets the value or formula of a cell in the spreadsheet.
        If the cell does not exist, it is created.
        If a value is provided, the cell's formula (if any) is removed.
        If a formula is provided, the cell's value is updated based on the formula.

        :param cell_name: The name of the cell to set.
        :param value: The value to set in the cell.
        :param formula: The formula to set in the cell.
        """
        if not self.is_valid_cell_name(cell_name):
            print(f"Invalid cell name '{cell_name}'."
                  f" Cell names must be in the format 'A1', 'B2', 'AZ10' etc.")
            return

        # Ensure the cell exists in the dictionary; if not, create a new one
        if cell_name not in self.cells:
            self.cells[cell_name] = Cell()

        # Update the cell's value or formula
        cell = self.cells[cell_name]
        if value is not None:
            if cell.formula:
                self.remove_cell(cell_name)
            self.set_cell_value(cell, value)
        if formula is not None:
            self.set_cell_formula(cell, cell_name, formula)

    def set_cell_value(self, cell: Cell, value: Any) -> None:
        """
        Sets the value of a cell in the spreadsheet.
        If the value can be converted to a float, it is stored as a float.
        Otherwise, it is stored as is.

        :param cell: The cell to set the value for.
        :param value: The value to set in the cell.
        """
        try:
            cell.set_value(float(value))
        except:
            cell.value = value

    def set_cell_formula(self, cell: Cell, cell_name: str, formula: str) -> None:
        """
        Sets the formula of a cell in the spreadsheet.
        If the formula is a cell name, the cell's value is updated based on the referenced cell's value.
        If the formula is a range of cells, the cell's value is updated based on the values of the cells in the range.

        :param cell: The cell to set the formula for.
        :param cell_name: The name of the cell to set the formula for.
        :param formula: The formula to set in the cell.
        """
        if isinstance(formula, int) or isinstance(formula, float):
            cell.value = formula
            cell.formula = None
            return
        # Check if the formula is a cell name
        if self.is_valid_cell_name(formula):
            referenced_cell = self.get_cell(formula)
            if formula == cell_name:
                print("The cell cannot be dependent on itself.")
                return
            if referenced_cell:
                cell.value = referenced_cell.value
                cell.formula = formula
                # Add the referenced cell to the dependents of the current cell
                cell.add_dependent(formula)
        else:
            # If the formula is not a cell name, it might be a range of cells,
            # unless it's a SQRT formula
            dependencies = []
            if formula.startswith("SQRT"):
                # Iterate over the formula to find the opening parenthesis
                for index in range(len(formula)):
                    if formula[index] == "(":
                        # Remove the operation and the opening parenthesis from the formula
                        check_formula = formula[index + 1:]
                        # Remove the closing parenthesis from the formula
                        check_formula = check_formula.replace(")", "")
                        if not self.is_valid_cell_name(check_formula):
                            print("SQRT formula must be in the format 'SQRT(cell)'.\n"
                                  "For example: 'SQRT(A1)'.")
                        else:
                            dependencies.append(check_formula)
                        break
            if ":" in formula:
                cells = self.valid_cells_index(formula)
                # Get the list of cell names in the range
                if cells:
                    dependencies = self.get_range_cells(cells[0], cells[1])
            else:
                # If the formula is not a range of cells, it might be a single cell
                for index, sign in enumerate(formula):
                    if sign in ['*', '/', '+', '-']:
                        parts = [formula[:index], sign, formula[index + 1]]
                        if self.is_valid_cell_name(parts[0]):
                            dependencies.append(parts[0])
                        if self.is_valid_cell_name(parts[2]):
                            dependencies.append(parts[2])
                        break
            if cell_name in dependencies:
                print("The cell cannot be dependent on itself.")
                return
            for dep_name in dependencies:
                if dep_name not in self.cells:
                    self.cells[dep_name] = Cell()
                # Add the cell to the dependents of the referenced cells
                self.cells[dep_name].add_dependent(cell_name)
            cell.formula = formula
            cell.value = cell.calculated_value(self)
            if cell.value is None:
                return

    def get_cell(self, cell_name: str) -> Any:
        """
        retrieves a Cell object from the cell's dictionary.

        :param cell_name: The name of the cell to retrieve.
        :return: The Cell object if found, None otherwise.
        """
        if not self.is_valid_cell_name(cell_name):
            return
        if cell_name in self.cells:
            return self.cells[cell_name]
        return

    def get_cell_value(self, cell_name: str) -> Any:
        """
        Retrieves the value of a cell in the spreadsheet.
        If the cell has a formula, the formula is evaluated and the result is returned.

        :param cell_name: The name of the cell to retrieve the value for.
        :return: The value of the cell, or an error message if the cell does not exist.
        """
        if not self.is_valid_cell_name(cell_name):
            print(f"Invalid cell name '{cell_name}'."
                  f" Cell names must be in the format 'A1', 'B2', 'AZ10' etc.")
            return
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
        :return: the answer of the formula. if the formula doesn't meet the string requirements, None.
        """
        # Iterate over the formula to find the opening parenthesis
        for index in range(len(formula)):
            if formula[index] == "(":
                # Remove the operation and the opening parenthesis from the formula
                formula = formula[index + 1:]
                break
        # Remove the closing parenthesis from the formula
        formula = formula.replace(")", "")
        if self.is_valid_cell_name(formula):
            return self.get_cell_value(formula)
        parts = []
        for index, sign in enumerate(formula):
            if sign in ['*', '/', '+', '-']:
                parts = [formula[:index], sign, formula[index + 1:]]
        if len(parts) == 3:
            side1, operation, side2 = parts
            # Try converting the first operand to a number, if it fails, treat it as a cell reference
            try:
                value1 = float(side1)
            except ValueError:
                value1 = self.get_cell_value(side1)
                if not isinstance(value1, int) and not isinstance(value1, float):
                    return
            # Repeat the process for the second operand
            try:
                value2 = float(side2)
            except ValueError:
                value2 = self.get_cell_value(side2)
                if not isinstance(value2, int) and not isinstance(value2, float):
                    return
            if value2 is None or value1 is None:
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
                    print("Error: Division by zero.")
                    return
            else:
                print("Unsupported operation.")
                return
        else:
            print("Invalid formula format.")
            return

    def evaluate_formula(self, formula: str) -> Any:
        """
        Evaluates a formula in the spreadsheet.
        The formula can be a regular formula, or a special formula like "AVERAGE", "SUM", "MIN", "MAX", or "SQRT".

        :param formula: The formula to evaluate.
        :return: The result of the formula, or an error message if the formula is invalid.
        """
        # Check if the formula starts with "AVERAGE"
        if formula.startswith("AVERAGE"):
            try:
                # Extract the range of cells from the formula
                cells = self.valid_cells_index(formula)
                # Calculate and return the average of the cells in the range
                return self.calculate_average(cells[0], cells[1])
            except Exception as err:
                print(f"Error: {str(err)}")
                return
        # Check if the formula starts with "SUM"
        if formula.startswith("SUM"):
            try:
                # Extract the range of cells from the formula
                cells = self.valid_cells_index(formula)
                # Calculate and return the sum of the cells in the range
                return self.calculate_sum(cells[0], cells[1])
            except Exception as err:
                print(f"Error: {str(err)}")
                return

        # Check if the formula starts with "MIN"
        if formula.startswith("MIN"):
            try:
                # Extract the range of cells from the formula
                cells = self.valid_cells_index(formula)
                # Find and return the minimum value in the range
                return self.find_min(cells[0], cells[1])
            except Exception as err:
                print(f"Error: {str(err)}")
                return

        # Check if the formula starts with "MAX"
        if formula.startswith("MAX"):
            try:
                # Extract the range of cells from the formula
                cells = self.valid_cells_index(formula)
                # Find and return the maximum value in the range
                return self.find_max(cells[0], cells[1])
            except Exception as err:
                print(f"Error: {str(err)}")
                return
        # Check if the formula starts with "SQRT"
        if formula.startswith("SQRT"):
            try:
                # Extract the cell name from the formula
                formula = formula[5:-1]
                value = float(self.get_cell_value(formula))
                return math.sqrt(value)
            except Exception as err:
                print(f"Error: {str(err)}")
                return
        # If the formula does not start with any of the special keywords,
        # treat it as a regular formula
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
        return values

    def find_min(self, start: str, end: str) -> Any:
        """
        finds the minimum cell value in a specific range that was given

        :param start: The starting cell name of the range.
        :param end: The ending cell name of the range.
        :return: float: the minimum value in the range.
        """
        cell_names = self.cells_values_list(start, end)
        if cell_names:
            return min(cell_names)
        return

    def find_max(self, start: str, end: str) -> Any:
        """
        finds the maximum cell value in a specific range that was given

        :param start: The starting cell name of the range.
        :param end: The ending cell name of the range.
        :return: float: the maximum value in the range.
        """
        cell_names = self.cells_values_list(start, end)
        if cell_names:
            return max(cell_names)
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

    def valid_cells_index(self, formula: str) -> Any:
        """
        gets a string with specific formula.
        Retrieves the specific 2 cells that in the formula.

        :param formula: a string of an operation and 2 specific cells to check the range between them.
        :return: the 2 cells that in the formula. for example:  ("AVERAGE(A1:B2)" -> ("A1", "B2")
        """

        # Iterate over the formula to find the opening parenthesis
        for index in range(len(formula)):
            if formula[index] == "(":
                # Remove the operation and the opening parenthesis from the formula
                formula = formula[index + 1:]
                break
        # Remove the closing parenthesis from the formula
        formula = formula.replace(")", "")
        # Split the formula into two cell names
        cell_list = formula.split(":")
        # Check if the formula contains exactly two cell names
        if len(cell_list) != 2:
            print("the formula does not fit the requirements")
            return
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
                return sum(values_list) / len(values_list)
            except Exception as err:
                print(f"Error: {str(err)}")
        return

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

    def get_range_cells(self, start: str, end: str) -> Any:
        """
        creates a list of cell names in a range that can span multiple columns and rows.
        for example: (A1, B2) -> ["A1", "A2", "B1", "B2"]

        :param start: The starting cell name of the range.
        :param end: The ending cell name of the range.
        :return List[str]: A list of cell names within the specified range.
        """
        if not self.is_valid_cell_name(start) or not self.is_valid_cell_name(end):
            print(f"Invalid cell name. Cell names must be in the format 'A1', 'B2', 'AZ10' etc.")
            return
        # Use regular expressions to separate the letters and digits in each cell
        start_match = re.match(r"([A-Z]+)([0-9]+)", start)
        end_match = re.match(r"([A-Z]+)([0-9]+)", end)

        if start_match is None or end_match is None:
            print(f"Invalid cell name. Cell names must be in the format 'A1', 'B2', 'AZ10' etc.")
            return

        start_col, start_row = start_match.groups()
        end_col, end_row = end_match.groups()
        # Translates the col index into an integer
        start_col_index = self.col_letter_to_index(start_col)
        end_col_index = self.col_letter_to_index(end_col)

        if start_col_index > end_col_index or start_row > end_row:
            print(f"Invalid cells range. '{end}' comes after '{start}")
            return
        # Creates the list of all the indexes as strings.
        cells = []
        for col in range(start_col_index, end_col_index + 1):
            for row in range(int(start_row), int(end_row) + 1):
                col_letter = self.col_index_to_letter(col)
                cell_name = f"{col_letter}{row}"
                cells.append(cell_name)
        return cells

    def remove_cell(self, cell_name: str) -> None:
        """
        Removes a cell from the spreadsheet.
        If the cell has a formula, it is also removed from the dependents of other cells.

        :param cell_name: The name of the cell to remove.
        """
        if cell_name in self.cells:
            cell = self.get_cell(cell_name)
            if cell.formula:
                # If the cell has a formula, it might be a dependent of other cells
                for other_cell in self.cells.values():
                    if cell_name in other_cell.dependents:
                        # If the cell is a dependent of another cell, remove it
                        other_cell.remove_dependent(cell_name)
            # remove the cell's arguments
            cell.value = None
            cell.formula = None

    def max_row(self) -> int:
        """
        Retrieves the maximum row index that has been used in the spreadsheet.

        :return: The maximum row index.
        """
        # If there are no cells, return 0
        if not self.cells:
            return 0
        rows = [int(cell.lstrip(ALL_LETTERS)) for cell in self.cells.keys()]
        return max(rows)

    def max_col_index(self) -> int:
        """
        Retrieves the maximum column index that has been used in the spreadsheet.

        :return: The maximum column index.
        """
        # If there are no cells, return 0
        if not self.cells:
            return 0
        cols = [self.col_letter_to_index(cell.rstrip(ALL_DIGITS)) for cell in self.cells.keys()]
        return max(cols)

    def to_dict(self) -> Dict[str, Dict[str, Any]]:
        """
        Converts the spreadsheet to a dictionary.

        :return: A dictionary representation of the spreadsheet.
        """
        return {
            cell_name: self.update_and_get_cell_dict(cell_name)
            for cell_name in self.cells.keys()
        }

    def update_and_get_cell_dict(self, cell_name: str) -> Dict[str, Any]:
        """
        Updates the value of a cell (if it has a formula) and returns its dictionary representation.

        :param cell_name: The name of the cell to update and convert to a dictionary.
        :return: A dictionary representation of the cell.
        """
        cell = self.cells[cell_name]
        if cell.formula:
            cell.value = cell.calculated_value(self)
        return cell.to_dict()

    def create_graph(self, graph_type: str, x_range: str, y_range: str) -> None:
        """
        Creates a graph based on the values of the cells in two ranges in the spreadsheet.

        :param graph_type: The type of the graph to create. Can be "bar" or "pie".
        :param x_range: The range of cells to use for the x-axis of the graph.
        :param y_range: The range of cells to use for the y-axis of the graph.
        """
        try:
            # Get the cell names for the x and y data
            x_cells = self.get_range_cells(*x_range.split(':'))
            y_cells = self.get_range_cells(*y_range.split(':'))

            if not x_cells or not y_cells:
                print(GRAPH_ERROR)
                return

            # Retrieve the values of these cells
            x_data = [self.get_cell(cell).value for cell in x_cells]
            y_data = [self.get_cell(cell).value for cell in y_cells]

            # Create the graph
            if graph_type.lower() == 'bar':
                x_label = input("Enter the x-axis label: ")
                y_label = input("Enter the y-axis label: ")
                plt.xlabel(x_label)
                plt.ylabel(y_label)
                title = input("Enter the title of the graph: ")
                plt.title(title)
                plt.bar(x_data, y_data)
            elif graph_type.lower() == 'pie':
                plt.pie(y_data, labels=x_data, autopct='%1.1f%%')
            else:
                print(f"Invalid graph type: {graph_type}")
                return
            # Display the graph and wait for it to be closed before continuing
            plt.show()

        except:
            print(GRAPH_ERROR)
            return
