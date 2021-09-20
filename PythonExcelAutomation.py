import operator
import openpyxl as pyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.xml.constants import MIN_COLUMN

def update_workbook(file_name):

    operations = {
        "multiplying" : operator.mul,
        "dividing" : operator.truediv,
        "adding" : operator.add,
        "subtracting" : operator.sub
    }


    wrkbk = pyxl.load_workbook(file_name)                                                                                       # Load the user specified excel workbook
    sheet_selected = input("Which sheet are you accessing? Type it exactly, with no spaces, between 'Sheet' and the number: ")
    sheet = wrkbk[sheet_selected]                                                                                               # Retrieve the user specified desired excel sheet
    print("Your selected sheet has", sheet.max_row, "rows.")
    iterated_rows = int(input("How many rows would you like to iterate for: "))
    user_operation = input("What operation are we performing? Please type multiplying, dividing, adding, or subtracting: ").lower()

    while (user_operation not in ["multiplying", "dividing", "adding", "subtracting"]):
        user_operation = input("Invalid operation! Please try again: ")


    operation_value = float(input("Enter numerical value for the operation: "))
    col_selected = int(input(f"Which column number values should the program be {user_operation} {operation_value} by (this column should contain your initial/original values): "))
    new_value_col = int(input("Which column number should the program store the newly calculated values: "))

    for row in range(2, iterated_rows + 1):                                                 # Start from 2 in order to ignore named/titled first row
        cell = sheet.cell(row, col_selected)
        corrected_price = operations[f"{user_operation}"](cell.value, operation_value)      # Allows the earlier, user specified, operation to be used here to calc new price
        corrected_price_cell = sheet.cell(row, new_value_col)                               # Creates the cell object
        corrected_price_cell.value = corrected_price                                        # Value of cell object now holds corrected price

    make_graph = input("Would you like to a basic bar graph of the newly calculated values to be made? Please type 'Yes' or 'No': ").lower()
    if make_graph == "yes":
        graph_cell = input("In what cell would you like to place the graph in: ")
        values = Reference(sheet, min_row = 2, max_row = sheet.max_row, min_col = new_value_col, max_col = new_value_col)   # Reference, to select a range of values - setup our future graph
        chart = BarChart()
        chart.add_data(values)
        sheet.add_chart(chart, f"{graph_cell}")     # Place chart starting in user-defined cell
    elif make_graph == "no":
        pass

    new_file_name = input("What would you like to save the file as: ")
    wrkbk.save(new_file_name)                                                                                                       
