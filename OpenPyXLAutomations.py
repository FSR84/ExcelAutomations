from openpyxl.utils.cell import get_column_letter
from openpyxl.styles import Font, PatternFill, Border, Side
import openpyxl
import os

# load workbook and define target worksheet by name
curr_dir = os.getcwd() + "\\ExcelAutomations\\Excel Files\\"
excel3 = "all-shifts.xlsx"

wb = openpyxl.load_workbook(curr_dir + excel3)
ws = wb["Shifts"]  # wb.active  # (ActiveSheet in VBA)

# delete first row if cell B1 is empty
if ws["B1"].value == None:
    ws.delete_rows(1, 1)

# insert 1st row if cell B1 is not empty
if ws["B1"].value != None:
    ws.insert_rows(1, 1)

# add missing column name
if ws["A2"].value == None:
    ws["A2"] = "Index"

# change font size and apply bold and interior color to an entire row (each cell separately, needs a separate loop for each row)
for cell in ws["2:2"]:
    cell.font = Font(size=12, color="403151", bold=True)
    cell.fill = PatternFill(fgColor="E4DFEC", fill_type="solid")

# define current region
lastrow = ws.max_row
lastcol = ws.max_column
lastcolletter = get_column_letter(lastcol)

print(str(lastrow) + " is the last row number.")
print(str(lastcol) + " is the last column number.")
print(lastcolletter + str(lastrow) + " is the last cell in the current region.")

# remove all borders, from row 3 onwards remove interior color (fill)
no_fill = PatternFill(fill_type=None)
side = Side(border_style=None)
no_border = Border(left=side, right=side, top=side, bottom=side)

for row in ws:
    for cell in row:
        cell.border = no_border
        if cell.row > 2:
            cell.fill = no_fill

# add borders to table (limited by current region)
side = Side(border_style="thin")
border = Border(left=side, right=side, top=side, bottom=side)

for row in ws.iter_rows(min_row=2, max_row=lastrow, min_col=1, max_col=lastcol):
    for cell in row:
        cell.border = border

# save workbook
wb.save(curr_dir + excel3)

# open Excel instance with the workbook
os.chdir(curr_dir)
os.system(excel3)

### https://automatetheboringstuff.com/chapter12/
