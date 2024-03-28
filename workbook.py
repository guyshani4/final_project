from electronic_sheet import *
import csv, json
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

class Workbook:
    """
    Represents a workbook containing multiple spreadsheet tabs,
    similar to a workbook in Excel containing multiple sheets.
    """

    def __init__(self, name = None):
        """
        Initializes a new workbook with an empty dictionary of sheets.
        """
        self.sheets = {}
        self.name = name

    def add_sheet(self, sheet_name):
        """
        Adds a new spreadsheet to the workbook with the given name.
        :param sheet_name: The name of the new sheet. Must be unique within the workbook.
        """
        if sheet_name in self.sheets:
            print(f"Sheet '{sheet_name}' already exists.")
        else:
            self.sheets[sheet_name] = Spreadsheet(sheet_name)
            print(f"Sheet '{sheet_name}' added to the workbook.")

    def remove_sheet(self, sheet_name):
        """
        Removes the spreadsheet with the given name from the workbook, if it exists.
        :param sheet_name: The name of the sheet to be removed.
        """
        if sheet_name in self.sheets:
            del self.sheets[sheet_name]
            print(f"Sheet '{sheet_name}' has been removed.")
        else:
            print(f"Sheet '{sheet_name}' does not exist.")

    def get_sheet(self, sheet_name):
        """
        Retrieves the spreadsheet with the given name, if it exists.
        :param sheet_name: The name of the sheet to retrieve.
        :return: The Spreadsheet object with the given name, or None if it does not exist.
        """
        return self.sheets.get(sheet_name, None)

    def list_sheets(self):
        """
        Lists the names of all spreadsheets in the workbook.

        :return: A list of sheet names in the workbook.
        """
        return list(self.sheets.keys())

    def print_list(self):
        print("Sheets in the workbook:")
        for name in self.sheets.keys():
            print(name)


    def rename_sheet(self, old_name, new_name):
        """
        Renames an existing sheet from old_name to new_name, if old_name exists and new_name does not.
        :param old_name: The current name of the sheet to be renamed.
        :param new_name: The new name for the sheet.
        """
        if old_name not in self.sheets:
            print(f"Sheet '{old_name}' does not exist.")
        elif new_name in self.sheets:
            print(f"Sheet '{new_name}' already exists.")
        else:
            self.sheets[new_name] = self.sheets.pop(old_name)
            print(f"Sheet '{old_name}' has been renamed to '{new_name}'.")

    def to_dict(self):
        return {sheet_name: sheet.to_dict() for sheet_name, sheet in self.sheets.items()}

    def dict_print(self):
        print(self.to_dict())

    def load_and_open_workbook(self, filename):
        """
        Loads a workbook file and opens it for editing.
        :param filename: The name of the workbook file to be opened.
        :return: The loaded Workbook instance.
        """
        with open(filename, 'r') as f:
            workbook_dict = json.load(f)

        for sheet_name, sheet_data in workbook_dict.items():
            spreadsheet = Spreadsheet()
            spreadsheet.load(sheet_data)
            self.sheets[sheet_name] = spreadsheet
        return self

    def export_to_json(self, filename):
        """
        Saves the workbook to a file in JSON format.
        :param filename: The name of the file to save the workbook to.
        """
        workbook_dict = {name: sheet.to_dict() for name, sheet in self.sheets.items()}
        with open(filename + ".json", 'w') as f:
            json.dump(workbook_dict, f)

    def export_to_pdf(self, filename):
        """
        Exports the workbook to a PDF file.
        :param filename: The name of the PDF file to be created.
        """
        for sheet_name, spreadsheet in self.sheets.items():
            c = canvas.Canvas(f"{filename}_{sheet_name}.pdf", pagesize=letter)
            width, height = letter
            # Increased x_offset for more space from the left border
            x_offset = 50  # Increased offset for better visibility of the first column
            y_offset = 70  # Adjusted for aesthetic spacing from the top

            # Adjusted column spacing for a prettier layout
            column_spacing = 60  # Increase if your columns are too close to each other

            # Adjusted row spacing for a prettier layout
            row_spacing = 20  # Increase if your rows are too close to each other

            for i in range(1, spreadsheet.max_row() + 1):
                for j in range(spreadsheet.max_col_index() + 1):
                    cell_name = f"{spreadsheet.col_index_to_letter(j)}{i}"
                    cell_value = spreadsheet.get_cell_value(cell_name)
                    if cell_value is not None:
                        x_position = x_offset + j * column_spacing
                        y_position = height - y_offset - i * row_spacing
                        c.drawString(x_position, y_position, str(cell_value))
            c.save()

    def export_to_csv(self, filename):
        """
        Exports the workbook to a CSV file.
        :param filename: The name of the CSV file to be created.
        """
        for spreadsheet in self.sheets:
            with open(f"{filename}_{spreadsheet.name}.csv", 'w', newline='') as f:
                writer = csv.writer(f)
                for i in range(1, spreadsheet.max_col_index() + 1):
                    row = [spreadsheet.get_cell_value(f"{spreadsheet.col_index_to_letter(j)}{i}")
                           for j in range(1, spreadsheet.max_col() + 1)]
                    writer.writerow(row)

