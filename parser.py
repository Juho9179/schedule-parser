################
#
# Parser
#
################

#
#
#
# Parsimisprosessi:
# 1. lue excel josta haetaan dataa
# 2. etsi excel kirjasto
#
#
#
#
#
#
#
#
#
#
#
#

import openpyxl
import openpyxl.utils as utils
import sys

filename = ""
employee_name = ""
workbook = ""

# Initialises file name and employee name.
def init():
    global filename
    global employee_name
    global workbook
    argument_list = sys.argv[1:]
    for i in argument_list:
        if i.endswith(".xlsx"):
            filename = argument_list.pop(argument_list.index(i))
    employee_name = ' '.join(argument_list)
    
    workbook = openpyxl.load_workbook(filename)


# Returns row number for employee and checks whether dual row or single-row
# Returns object:
# {'coordinate': 'B15', 'column': 2, 'row': 15, 'dual': False}
# Returns False if employee is not found.
def get_row_info(employee_name):
    row_info = {}
    row_info["coordinate"] = ""
    row_info["column"] = ""
    row_info["row"] = 0
    row_info["dual"] = False

    ws = workbook.active
    for row in ws.iter_rows():
        for cell in row:
            # check if cell contents match employee name
            if (cell.internal_value == employee_name):
                row_info["coordinate"] = cell.coordinate
                row_info["column"] = cell.column
                row_info["row"] = cell.row
                break
                

    # Check for errors
    try:
        # check whether row below is another employee or empty: singular or dual row
        next_entry = ws.cell(row_info["row"] + 1, row_info["column"]).internal_value
        if (next_entry == None):
            row_info["dual"] = True

    except TypeError:
        print("ERROR: Employee " + employee_name + " not found.")
        return False

    return row_info

def main():
    # parse  n stuf
    init()
    print(get_row_info(employee_name))
    return

if __name__ == "__main__":
    main()