import xlwings as xw
import os

# load workbook and define target worksheet by name
curr_dir = os.getcwd() + "\\ExcelAutomations\\Excel Files\\"
excel3 = "all-shifts.xlsx"

wbxl=xw.Book(curr_dir + excel3)
sh = wbxl.sheets['Shifts']


# evaluate a single cell
eval1 = sh.range('H3').value
print(eval1)


# evaluate an entire column - need to limit the range manually to avoid Nones in the results
lastrow = sh.range('H:H').current_region.last_cell.row
eval1 = sh.range('H3:H' + str(lastrow)).value
print(eval1)

