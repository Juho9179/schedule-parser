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
workbook = openpyxl.load_workbook("./test-material/ENSIH9_21(14.6-4.7).xlsx")


# Returns row number for employee and checks whether dual row or single-row
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
                
    # check whether row below is another employee or empty: singular or dual row
    next_entry = ws.cell(row_info["row"] + 1, row_info["column"]).internal_value
    if (next_entry == None):
        row_info["dual"] = True

    return row_info

def main():
    # parse  n stuf
    print(get_row_info("EH9"))
    return

if __name__ == "__main__":
    main()