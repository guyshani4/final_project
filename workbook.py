from electronic_sheet import *

class Workbook:
    """
    Represents a workbook containing multiple spreadsheet tabs,
    similar to a workbook in Excel containing multiple sheets.
    """
    def __init__(self):
        """
        Initializes a new workbook with an empty dictionary of sheets.
        """
        self.sheets = {}

    def add_sheet(self, sheet_name):
        """
        Adds a new spreadsheet to the workbook with the given name.
        :param sheet_name: The name of the new sheet. Must be unique within the workbook.
        """
        if sheet_name in self.sheets:
            print(f"Sheet '{sheet_name}' already exists.")
        else:
            self.sheets[sheet_name] = Spreadsheet()
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
        print("Sheets in the workbook:")
        for name in self.sheets.keys():
            print(name)
        return list(self.sheets.keys())

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