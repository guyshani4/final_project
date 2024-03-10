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
            print("  set [cell] [value] - Set the value of a cell (value can be a number or a string).")
            print("  formula [cell] [formula] - Set the formula for a cell and updates its value.")
            print("  get [cell] - Get the value of a cell, if exists.")
            print("  quit - Exit the program with option to save.")
        elif command.strip().startswith("set "):
            _, cell_name, value = command.split(maxsplit=2)
            spreadsheet.set_cell(cell_name, value=value)
            print(f"Set {cell_name} to {value}.")
        elif command.startswith("formula "):
            _, cell_name, formula = command.split(maxsplit=2)
            ss.set_cell(cell_name, formula=formula)
            print(f"Set formula for {cell_name} to {formula}.")
        elif command.startswith("get "):
            _, cell_name = command.split()
            value = ss.get_cell_value(cell_name)
            print(f"{cell_name}: {value if value is not None else '-'}")
        elif command.startswith("eval "):
            _, formula = command.split(maxsplit=1)
            try:
                result = ss.evaluate_formula(formula)
                print(f"Result: {result}")
            except Exception as e:
                print(f"Error evaluating formula: {str(e)}")
        else:
            print("Unknown command. Type 'help' for options.")

if __name__ == "__main__":
    main()