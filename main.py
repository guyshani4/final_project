from electronic_sheet import *
from workbook import *
def get_spreadsheet():
    workbook = Workbook()
    sheet_name = ""
    decision1 = input("Welcome to WorkBook! would you like to load an existing workbook, or start a new one?\n"
                      "type 'new' or 'open' to start: ")
    while decision1 not in ["new", "open"]:
        decision1 = input("not a valid command. type 'new' or 'open' to start: ")
    if decision1.lower() == "open":
        filename = ""
        while filename == "":
            filename = input("Enter the name of the file you want to open: ")
            try:
                workbook = Workbook.load_and_open_workbook(filename)
                print(f"Opened {filename} successfully.")
            except Exception as err:
                print(f"Error: {str(err)}")
                filename = ""
    else:
        print("Great! let's start a new workbook. ")
        workbook_name = input("what would you like to call your workbook? ")
        workbook = Workbook(workbook_name)
        print("type the name of the first sheet in your project.")
        sheet_name = input("type here: > ").strip()
        workbook.add_sheet(sheet_name)
        print(f"You're in {sheet_name} sheet. Type 'help' for options, or start editing.")
    return workbook, workbook.get_sheet(sheet_name)


def main():
    workbook, spreadsheet = get_spreadsheet()
    print("Type 'help' for options, or start editing.")
    while True:
        command = input("> ").strip()
        if command.lower() == "quit":
            if input("Are you sure you want to quit? ").lower() == "yes":
                if input("Would you like to save the workbook? ").lower() == "yes":
                    filename = input("what file name? ")
                    workbook.save_workbook(filename)
                    print("exiting workbook... Bye!")
                    break
                else:
                    print("exiting workbook... Bye!")
                    break
            else:
                continue
        elif command.lower() == "help":
            print("The Optional Commands:")
            print("  - set [cell] [value] - Set the value of a cell (value can be a number or a string).")
            print("  - formula [cell] [formula] - Set the formula for a cell and updates its value.\n"
                  "             PAY ATTENTION! the formulas should be combination of numbers and cells only.\n"
                  "             there are 4 special formulas: 'AVERAGE' 'MIN' 'MAX' 'SUM' 'SQRT'. \n             "
                  "these formulas should be typed in a specific form. \n"
                  "             for example: MAX(A1:B2) is correct and will set"
                  "the maximum number in the range of A1 and B2. \n "
                  "             for SQRT operator a valid form: SQRT(A1). \n")
            print("  - details - Get the detailed version of the spreadsheet.")
            print("  - quit - Exit the program with option to save.")
            print("  - show - shows the spreadsheet in an organized table")
            print("  - remove [cell] - Removes the cell's value")
            print("  - new - opens a new sheet in your workbook")
            print("  - sheets - if you want to see the sheet's list and choose which sheet to open")
            print("  - rename - if you want to rename a sheet")
            print("  - change sheet - if you want to rename a sheet")
            print("  - remove sheet - if you want to removes a sheet")
            print("  - save - if you want to save the workbook")
            print("  - export - if you want to export the workbook to a different file type")

        elif command.lower() == "save":
            filename = input("what file name? ")
            workbook.save_workbook(filename)
            print(f"Saved {filename} successfully.")

        elif command.lower() == "export":
            print("You can save the workbook in the following formats:")
            print("  - csv")
            print("  - pdf")
            save_format = input("Please enter the format you want to save the spreadsheet in: ")
            while save_format not in ["csv", "pdf"]:
                save_format = input("Invalid format. Please enter either 'csv', 'pdf', or 'json'.")
            filename = input("Please enter the filename: ")
            if save_format.lower() == "csv":
                workbook.export_to_csv(filename)
            elif save_format.lower() == "pdf":
                workbook.export_to_pdf(filename)
            else:
                print("Invalid format. Please enter either 'csv' or 'pdf'.")


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

        elif command.startswith("details"):
            print(spreadsheet.dict_print())

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
            workbook.print_list()
            new_name = input("which sheet would you like to get into? ")
            while new_name not in workbook.list_sheets():
                workbook.print_list()
                new_name = input("name did not found..."
                                 "which sheet would you like to open? ")
            spreadsheet = workbook.get_sheet(new_name)
            print(f"You're in {new_name} sheet. Type 'help' for options, or start editing.")

        elif command.startswith("remove sheet"):
            workbook.print_list()
            sheet_name = input("which sheet would you like to remove? ")
            while sheet_name not in workbook.list_sheets():
                workbook.print_list()
                sheet_name = input("name did not found..."
                                   "which sheet would you like to remove? ")




if __name__ == "__main__":
    main()