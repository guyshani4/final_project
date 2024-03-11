from electronic_sheet import *
from workbook import *

def main():
    workbook = Workbook()
    print("Welcome to the your WorkBook! first, you'll need to type the name of the first sheet in your project.")
    sheet_name = input("type here: > ").strip()
    workbook.add_sheet(sheet_name)
    spreadsheet = workbook.get_sheet(sheet_name)
    print(f"You're in {sheet_name} sheet. Type 'help' for options, or start editing.")

    while True:
        command = input("> ").strip()
        if command.lower() == "quit":
            if input("Are you sure you want to quit? ").lower() == "yes":
                if input("Would you like to save the spreadsheet? ").lower() == "yes":
                    filename = input("what file name? ")
                    spreadsheet.save_as(filename)
                else:
                    print("exiting spreadsheet...")
                    break
            else:
                continue
        elif command.lower() == "help":
            print("The Optional Commands:")
            print("  - set [cell] [value] - Set the value of a cell (value can be a number or a string).")
            print("  - formula [cell] [formula] - Set the formula for a cell and updates its value.\n"
                  "             PAY ATTENTION! the formulas should be combination of numbers and cells only.\n"
                  "             there are 4 special formulas: 'AVERAGE' 'MIN' 'MAX' 'SUM'. \n             "
                  "these formulas should be typed in a specific form. \n"
                  "             for example: MAX(A1:B2) is correct and will set"
                  "the maximum number in the range of A1 and B2 ")
            print("  - get [cell] - Get the value of a cell. if not exists, print '-'.")
            print("  - quit - Exit the program with option to save.")
            print("  - show - shows the spreadsheet in an organized table")
            print("  - remove [cell] - Removes the cell's value")
            print("  - new - opens a new sheet in your workbook")
            print("  - sheets - if you want to see the sheet's list and choose which sheet to open")
            print("  - rename - if you want to rename a sheet")
            print("  - change sheet - if you want to rename a sheet")




        elif command.startswith("set "):
            try:
                _, cell_name, value = command.split(maxsplit=2)
                spreadsheet.set_cell(cell_name, value=value)
            except Exception as err:
                print("oops. not a valid command")
                print(f"Error: {str(err)}")
                continue
            if spreadsheet.cells != {}:
                print(spreadsheet)


        elif command.startswith("formula "):
            try:
                _, cell_name, formula = command.split(maxsplit=2)
                spreadsheet.set_cell(cell_name, formula=formula)
            except Exception as err:
                print("oops. not a valid command")
                print(f"Error: {str(err)}")
                continue
            if spreadsheet.cells != {}:
                print(spreadsheet)

        elif command.startswith("get "):
            _, cell_name = command.split()
            value = spreadsheet.get_cell_value(cell_name)
            print(f"{cell_name}: {value if value is not None else '-'}")

        elif command.startswith("show"):
            print(spreadsheet)

        elif command.startswith("remove"):
            _, cell_name = command.split(maxsplit=1)
            spreadsheet.remove_cell(cell_name)
            print(spreadsheet)
        elif command.startswith("new"):
            sheet_name = input("name the new sheet: ")
            workbook.add_sheet(sheet_name)
            spreadsheet = workbook.get_sheet(sheet_name)
            print(f"You're in {sheet_name} sheet. Type 'help' for options, or start editing.")
        elif command.startswith("sheets"):
            print(workbook.list_sheets())
            sheet_name = input("which sheet would you like to open? ")
            spreadsheet = workbook.get_sheet(sheet_name)
            print(f"You're in {sheet_name} sheet. Type 'help' for options, or start editing.")
        elif command.startswith("rename"):
            workbook.print_list()
            sheet_name = input("which sheet would you like to rename? ")
            while sheet_name not in workbook.list_sheets():
                sheet_name = input("name did not found..."
                                   "which sheet would you like to rename? ")
            new_name = input("which name would you like to call it? ")
            workbook.rename_sheet(sheet_name, new_name)
            spreadsheet = workbook.get_sheet(new_name)
            print(f"You're in {new_name} sheet. Type 'help' for options, or start editing.")
        elif command.startswith("change"):
            new_name = input("which sheet would you like to get into? ")
            workbook.print_list()
            while new_name not in workbook.list_sheets():
                new_name = input("name did not found..."
                                 "which sheet would you like to get into? ")

            spreadsheet = workbook.get_sheet(new_name)
            print(f"You're in {new_name} sheet. Type 'help' for options, or start editing.")








if __name__ == "__main__":
    main()