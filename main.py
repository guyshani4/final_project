from electronic_sheet import *

def main():
    spreadsheet = Spreadsheet()
    print("Welcome to the Spreadsheet CLI. Type 'help' for options, or start editing.")

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
            print("  - redo/undo ")


        elif command.startswith("set "):
            try:
                _, cell_name, value = command.split(maxsplit=2)
                spreadsheet.set_cell(cell_name, value=value)
                # dfiuwdgivcbjkdbcijwhfubejfjudfbjdbcjwbciebfeb
                command = SetValueCommand(spreadsheet, cell_name, spreadsheet.get_cell_value(cell_name))
                spreadsheet.perform_command(command)
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

        elif command.startswith("redo"):
            spreadsheet.redo()
            print(spreadsheet)

        elif command.startswith("undo"):
            spreadsheet.undo()
            print(spreadsheet)




if __name__ == "__main__":
    main()