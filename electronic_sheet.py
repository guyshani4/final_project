import json

class Cell:
    """
    Represents a single cell in a Spreadsheet.
    """
    def __init__(self, value=None):
        """
        Initializes a new Cell instance.
        :param value: The initial value of the cell.
        if nothing was inserted as value, None as default.
        """
        self.value = value
        self.formula = None

    def evaluate(self, spreadsheet):
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
    Represents a spreadsheet, similar to an Excel page.
    """
    def __init__(self):
        """
        Initializes a new Spreadsheet instance with an empty dictionary of cells.
        """
        self.cells = {}

    def set_cell(self, cell_name, value, formula=None):
        """
        Sets or updates a cell's value and/or formula.
        :param cell_name: A string identifier for the cell (for example: "A1").
        :param value: The value to set in the cell.
        :param formula: An optional formula for the cell.
        If provided, the cell's value will be determined by this formula.
        """
        cell = Cell(value)
        cell.formula = formula
        self.cells[cell_name] = cell

    def get_cell(self, cell_name):
        """
        retrieves a Cell object from the cell's dictionary.
        :param cell_name: The name of the cell to retrieve.
        :return: The Cell object if found, None otherwise.
        """
        if cell_name in self.cells:
            return self.cells[cell_name]
        else:
            return None

    def get_cell_value(self, cell_name):
        """
        Retrieves the evaluated value of a cell.
        :param cell_name: The name of the cell to evaluate.
        :return: The evaluated value of the cell or an error message if the cell does not exist.
        """
        cell = self.get_cell(cell_name)
        if cell:
            try:
                return cell.evaluate(self)
            except Exception as err:
                return f"Error: {str(err)}"
        return "Cell does not exist"

    def evaluate_formula(self, formula):
        """
        calculates a given formula represented as a string.
        :param formula: A string formula, for example: "A1 + B1".
        :return: The result of the formula calculation.
        :raises ValueError: If the formula is invalid or contains unknown operations.
        """
        parts = formula.split()
        if len(parts) == 3:
            side1, operation, side2 = parts
            value1 = self.get_raw_value(side1)
            value2 = self.get_raw_value(side2)

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

    def get_raw_value(self, cell_name):
        """
        Retrieves the raw value of a cell without evaluating its formula.
        :param cell_name: The name of the cell to retrieve the value from.
        :return: The value of the cell.
        :raises ValueError: If the cell is empty, has an invalid value, or does not exist.
        """
        cell = self.get_cell(cell_name)
        if cell:
            if cell.value is not None:
                return cell.value
            else:
                raise ValueError(f"Cell {cell_name} is empty or has an invalid value.")
        else:
            raise ValueError(f"Cell {cell_name} does not exist.")


    def save_as(self, filename):
        """
        Saves the current state of the spreadsheet to a file in JSON format.
        :param filename: The name of the file to save the spreadsheet to.
        """
        data_as_dict = {name: {'value': cell.value, 'formula': cell.formula} for name, cell in self.cells.items()}
        with open(filename, 'w') as f:
            json.dump(data_as_dict, f)

    def load(self, filename):
        """
        Loads a spreadsheet from a file saved in JSON format.
        it loads the json to a Spreadsheet file again, initializing every cell again.
        :param filename: The name of the file to load the spreadsheet from.
        """
        with open(filename, 'r') as f:
            loaded_dict = json.load(f)
        for name, data in loaded_dict.items():
            self.set_cell(name, data['value'], data['formula'])