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

workbook = openpyxl.load_workbook("./test-material/ENSIH9_21(14.6-4.7).xlsx")


# Returns True / False whether row is a dual-row or a single-row entry
def is_dual(row):
    # not implemented
    return False

# Returns row number for employee
def get_row(employee):
    # not implemented
    rows = workbook.active.iter_rows()
    print(rows)
    return 0

def main():
    # parse  n stuf
    print(get_row("EH9"))
    return

if __name__ == "__main__":
    main()