from electronic_sheet import *
import csv, json, xlsxwriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from typing import *


class Workbook:
    """
    Represents a workbook containing multiple spreadsheet tabs,
    similar to a workbook in Excel containing multiple sheets.
    """

    def __init__(self, name: Optional[str] = None) -> None:
        """
        Initializes a new workbook with an empty dictionary of sheets.

        :param name: The name of the workbook. If not provided, the workbook will be unnamed.
        """
        self.sheets = {}
        self.name = name

    def add_sheet(self, sheet_name: str) -> None:
        """
        Adds a new spreadsheet to the workbook with the given name.
        If a sheet with the same name already exists, a message is printed and no sheet is added.

        :param sheet_name: The name of the new sheet. Must be unique within the workbook.
        """
        if sheet_name in self.sheets:
            print(f"Sheet '{sheet_name}' already exists.")
        else:
            self.sheets[sheet_name] = Spreadsheet(sheet_name)
            print(f"Sheet '{sheet_name}' added to the workbook.")

    def remove_sheet(self, sheet_name: str) -> None:
        """
        Removes the spreadsheet with the given name from the workbook, if it exists.
        If the sheet does not exist, a message is printed and no sheet is removed.

        :param sheet_name: The name of the sheet to be removed.
        """
        if sheet_name in self.sheets:
            del self.sheets[sheet_name]
            print(f"Sheet '{sheet_name}' has been removed.")
        else:
            print(f"Sheet '{sheet_name}' does not exist.")

    def get_sheet(self, sheet_name: str) -> Optional[Spreadsheet]:
        """
        Retrieves the spreadsheet with the given name, if it exists.
        If the sheet does not exist, None is returned.

        :param sheet_name: The name of the sheet to retrieve.
        :return: The Spreadsheet object with the given name, or None if it does not exist.
        """
        return self.sheets.get(sheet_name, None)

    def list_sheets(self) -> List[str]:
        """
        Lists the names of all spreadsheets in the workbook.

        :return: A list of sheet names in the workbook.
        """
        return list(self.sheets.keys())

    def print_list(self) -> None:
        """
        Prints the names of all spreadsheets in the workbook to the console.
        """
        print("Sheets in the workbook:")
        for name in self.sheets.keys():
            print(name)

    def rename_sheet(self, old_name: str, new_name: str) -> None:
        """
        Renames an existing sheet from old_name to new_name, if old_name exists and new_name does not.
        If old_name does not exist or new_name already exists, a message is printed and no sheet is renamed.

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

    def to_dict(self) -> Dict[str, Dict[str, Any]]:
        """
        Converts the workbook to a dictionary where each key is a sheet name and each value is a dictionary representation of the corresponding sheet.
        :return: A dictionary representation of the workbook.
        """
        return {sheet_name: sheet.to_dict() for sheet_name, sheet in self.sheets.items()}

    def dict_print(self) -> None:
        """
        Prints the dictionary representation of the workbook to the console.
        """
        print(self.to_dict())

    def export_to_json(self, filename: str) -> None:
        """
        Saves the workbook to a file in JSON format.

        :param filename: The name of the file to save the workbook to.
        """
        workbook_dict = {name: sheet.to_dict() for name, sheet in self.sheets.items()}
        with open(filename + ".json", 'w') as f:
            json.dump(workbook_dict, f)

    def export_to_pdf(self, filename: str) -> None:
        """
        Exports the workbook to a PDF file, with a table-like appearance including grid lines and row numbers.
        Each sheet is saved to a separate PDF file.

        :param filename: The base name of the PDF files to be created.
        The sheet name and .pdf extension are added automatically.
        """

        for sheet_name, spreadsheet in self.sheets.items():
            c = canvas.Canvas(f"{filename}_{sheet_name}.pdf", pagesize=letter)
            width, height = letter

            # Configuration for aesthetics
            x_offset = 60  # Adjusted to provide space for row numbers
            y_offset = 100
            column_spacing = 80  # Adjust for cell content width
            row_spacing = 20  # Adjust for cell content height
            cell_height = 18  # Height of each cell row

            # Draw table header for column names
            col_headers = [spreadsheet.col_index_to_letter(i) for i in range(spreadsheet.max_col_index() + 1)]
            c.setFont("Helvetica-Bold", 12)
            for j, col_header in enumerate(col_headers, start=0):
                x_position = x_offset + j * column_spacing
                c.drawString(x_position, height - y_offset + row_spacing, col_header)

            # Reset font for table content
            c.setFont("Helvetica", 10)

            # Draw cells, grid lines, and row numbers
            for i in range(1, spreadsheet.max_row() + 1):
                # Draw row numbers
                c.drawString(x_offset - 50, height - y_offset - i * row_spacing + (cell_height / 4), str(i))

                for j in range(spreadsheet.max_col_index() + 1):
                    col_letter = spreadsheet.col_index_to_letter(j)
                    cell_name = f"{col_letter}{i}"
                    cell_value = spreadsheet.get_cell_value(cell_name)
                    cell_value_str = "-" if cell_value is None else str(cell_value)

                    x_position = x_offset + j * column_spacing
                    y_position = height - y_offset - i * row_spacing
                    c.drawString(x_position, y_position, cell_value_str)

                    # Drawing the grid line around the cell
                    c.rect(x_position - 2, y_position - 2, column_spacing - 4, cell_height, fill=0)

            c.save()

    def export_to_csv(self, filename: str) -> None:
        """
        Exports the workbook to a CSV file.
        Each sheet is saved to a separate CSV file.

        :param filename: The base name of the CSV files to be created.
        The sheet name and .csv extension are added automatically.
        """
        # Iterate over each sheet in the workbook
        for sheet_name, spreadsheet in self.sheets.items():
            # Open a new CSV file for each sheet
            with open(f"{filename}_{sheet_name}.csv", 'w', newline='') as f:
                writer = csv.writer(f)
                for i in range(1, spreadsheet.max_row() + 1):
                    # For each row, create a list of cell values
                    row = [spreadsheet.get_cell_value(f"{spreadsheet.col_index_to_letter(j)}{i}")
                           for j in range(1, spreadsheet.max_col_index() + 1)]
                    writer.writerow(row)

    def export_to_excel(self, filename: str) -> None:
        """
        Exports the workbook to an Excel file.
        Each sheet is saved to a separate tab in the Excel file.
        :param filename: The name of the Excel file to be created. The .xlsx extension is added automatically.
        """

        # Create a new Excel workbook
        workbook = xlsxwriter.Workbook(f"{filename}.xlsx")
        # Iterate over each sheet in the workbook
        for sheet_name, spreadsheet in self.sheets.items():
            worksheet = workbook.add_worksheet(sheet_name)
            # Iterate over each cell in the sheet
            for i in range(1, spreadsheet.max_row() + 1):
                for j in range(spreadsheet.max_col_index() + 1):
                    cell_name = f"{spreadsheet.col_index_to_letter(j)}{i}"
                    cell_value = spreadsheet.get_cell_value(cell_name)
                    # Write the cell value to the Excel worksheet
                    worksheet.write(i - 1, j, cell_value)

        workbook.close()


def load_and_open_workbook(filename: str) -> Workbook:
    """
    Loads a workbook file and opens it for editing.
    The file is expected to be in JSON format,
    with each key being a sheet name and each value being a dictionary representation of the corresponding sheet.
    :param filename: The name of the workbook file to be opened.
    The .json extension is expected to be included in the filename.
    :return: The loaded Workbook instance.
    """
    # Open the JSON file
    with open(filename, 'r') as f:
        workbook_dict = json.load(f)
    # Remove the .json extension from the filename
    workbook_name = filename.rsplit('.', 1)[0]
    workbook = Workbook(workbook_name)
    # Iterate over the dictionary
    for sheet_name, sheet_data in workbook_dict.items():
        spreadsheet = Spreadsheet(sheet_name)
        for cell_name, cell_data in sheet_data.items():
            value = cell_data.get('value')
            formula = cell_data.get('formula')
            dependents = cell_data.get('dependents', [])
            cell = Cell(value=value, formula=formula)
            # Update the cell's dependents
            cell.update_dependents(dependents)
            # Set the cell in the Spreadsheet object
            spreadsheet.cells[cell_name] = cell
        # Add the Spreadsheet object to the workbook
        workbook.sheets[sheet_name] = spreadsheet

    return workbook
