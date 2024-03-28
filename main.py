from workbook import *


def get_spreadsheet():
    workbook = Workbook()
    decision1 = input("Welcome to WorkBook! would you like to load an existing workbook, or start a new one?\n"
                      "type 'new' or 'open' to start: ")
    while decision1.lower() not in ["new", "open"]:
        decision1 = input("not a valid command. type 'new' or 'open' to start: ")
    if decision1.lower() == "open":
        filename = ""
        while filename == "":
            filename = input("Enter the file you want to open: ")
            try:
                workbook = load_and_open_workbook(filename)
                print(f"Opened {filename} successfully.")
            except:
                print("it seems like the file does not fit the requirements.")
                filename = ""
        workbook.print_list()
        sheet_name = input("which sheet would you like to open? ")
        while sheet_name not in workbook.list_sheets():
            sheet_name = input("name did not found..."
                               "which sheet would you like to open? ")
        print(f"You're in {sheet_name} sheet. Type 'help' for options, or start editing.")

    else:
        print("Great! let's start a new workbook. ")
        workbook_name = input("what would you like to call your workbook? ")
        workbook = Workbook(workbook_name)
        print("type the name of the first sheet in your project.")
        sheet_name = input("type here: > ").strip()
        workbook.add_sheet(sheet_name)
        print(f"You're in {sheet_name} sheet. Type 'help' for options, or start editing.")
    return workbook, workbook.get_sheet(sheet_name)


help_text = """
                The Optional Commands:
                  - set [cell] [value] - Set the value of a cell (value can be a number or a string).
                  - set [cell] [formula] - Set the formula for a cell and updates its value.
                            PAY ATTENTION! the formula must start with "=" sign.
                            the formulas should be combination of numbers and cells only.
                            there are 4 special formulas: 'AVERAGE' 'MIN' 'MAX' 'SUM' 'SQRT'. 
                            these formulas should be typed in a specific form: 
                            for example: =MAX(A1:B2) is correct and set the maximum number in the range of A1 and B2. 
                            for SQRT operator a valid form: =SQRT(A1).
                  - details - Get the detailed version of the spreadsheet.
                  - quit - Exit the program with option to save.
                  - show - shows the spreadsheet in an organized table
                  - remove [cell] - Removes the cell's value
                  - new - opens a new sheet in your workbook
                  - sheets - if you want to see the sheet's list and choose which sheet to open
                  - rename - if you want to rename a sheet
                  - change sheet - if you want to rename a sheet
                  - remove sheet - if you want to removes a sheet
                  - save - if you want to save the workbook
                  - export - if you want to export the workbook to a different file type
            """


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
        if command.lower() == "help":
            print(help_text)
            continue

        if command.lower() == "save":
            workbook.export_to_json(workbook.name)
            print(f"Saved {workbook.name} successfully.")
            continue

        if command.lower() == "export":
            print("You can save the workbook in the following formats:")
            print("  - csv")
            print("  - pdf")
            print("  - excel")
            save_format = input("Please enter the format you want to save the spreadsheet in: ").lower()
            while save_format not in ["csv", "pdf", "excel"]:
                save_format = input("Invalid format. Please enter either 'csv', 'pdf', or 'json'.").lower()
            if save_format.lower() == "csv":
                workbook.export_to_csv(workbook.name)
            elif save_format.lower() == "pdf":
                workbook.export_to_pdf(workbook.name)
            elif save_format.lower() == "excel":
                workbook.export_to_excel(workbook.name)
            continue

        if command.lower().startswith("set"):
            try:
                _, cell_name, value = command.split(maxsplit=2)
                if not value.startswith("="):
                    spreadsheet.set_cell(cell_name, value=value)
                else:
                    formula = value[1:]
                    spreadsheet.set_cell(cell_name, formula=formula)
            except Exception as err:
                continue
            if spreadsheet.cells != {}:
                print(spreadsheet)
            continue

        if command.lower().startswith("details"):
            print(workbook.dict_print())
            continue

        if command.lower() == "show":
            print(spreadsheet)
            continue

        if command.lower() == "remove":
            _, cell_name = command.split(maxsplit=1)
            spreadsheet.remove_cell(cell_name)
            print(spreadsheet)
            continue

        if command.lower() == "new":
            sheet_name = input("name the new sheet: ")
            workbook.add_sheet(sheet_name)
            spreadsheet = workbook.get_sheet(sheet_name)
            print(f"You're in {sheet_name} sheet. Type 'help' for options, or start editing.")
            continue

        if command.lower() == "sheets":
            print(workbook.list_sheets())
            sheet_name = input("which sheet would you like to open? ")
            spreadsheet = workbook.get_sheet(sheet_name)
            print(f"You're in {sheet_name} sheet. Type 'help' for options, or start editing.")
            continue

        if command.lower() == "rename":
            workbook.print_list()
            sheet_name = input("which sheet would you like to rename? ")
            while sheet_name not in workbook.list_sheets():
                sheet_name = input("name did not found..."
                                   "which sheet would you like to rename? ")
            new_name = input("which name would you like to call it? ")
            workbook.rename_sheet(sheet_name, new_name)
            spreadsheet = workbook.get_sheet(new_name)
            print(f"You're in {new_name} sheet. Type 'help' for options, or start editing.")
            continue

        if command.lower() == "change":
            workbook.print_list()
            new_name = input("which sheet would you like to get into? ")
            while new_name not in workbook.list_sheets():
                workbook.print_list()
                new_name = input("name did not found..."
                                 "which sheet would you like to open? ")
            spreadsheet = workbook.get_sheet(new_name)
            print(f"You're in {new_name} sheet.")
            continue

        if command.lower() == "remove sheet":
            workbook.print_list()
            sheet_name = input("which sheet would you like to remove? ")
            while sheet_name not in workbook.list_sheets():
                workbook.print_list()
                sheet_name = input("name did not found..."
                                   "which sheet would you like to remove? ")
            workbook.remove_sheet(sheet_name)
            spreadsheet = workbook.get_sheet
            if workbook.list_sheets():
                sheet_name = workbook.list_sheets()[0]
                spreadsheet = workbook.get_sheet(sheet_name)
                print(f"You're in {sheet_name} sheet.")
            else:
                print("There are no sheets in the workbook.")
            continue


if __name__ == "__main__":
    main()
