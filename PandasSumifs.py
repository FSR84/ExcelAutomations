import pandas as pd
import os


# load workbook
curr_dir = os.getcwd() + "\\ExcelAutomations\\Excel Files\\"
excel1 = curr_dir + "shift-data.xlsx"

df = pd.read_excel(excel1, sheet_name='second')


# sumif: one result with one condition
sum1 = df.groupby("Name")["Products Produced (Units)"].sum()
print(sum1)


# double sumif: two results with one condition
sum2 = df.groupby("Name")["Production Run Time (Min)", "Products Produced (Units)"].sum()
print(sum2)


# sumifs: one result with two conditions
sum3 = df.groupby(['Name', 'Product'])['Products Produced (Units)'].sum()
print(sum3)


# save as new excel file
sum3.to_excel(curr_dir + 'second-shift-sumifs.xlsx', merge_cells=False)
