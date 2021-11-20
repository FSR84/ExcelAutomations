import pandas as pd
import os

curr_dir = os.getcwd() + "/ExcelAutomations/Excel Files/"

# 1st wb has data from two shifts in 2 different tabs, third shift data is in the second wb
excel1 = curr_dir + "shift-data.xlsx"
excel2 = curr_dir + "third-shift-data.xlsx"

df_first = pd.read_excel(excel1, sheet_name="first")
df_second = pd.read_excel(excel1, sheet_name="second")
df_third = pd.read_excel(excel2)

print(df_first)
print(df_first["Product"])

# create a union between three data sources; all 3 have the same columns; adds an index in column A, but without a column name in A1
df_all = pd.concat([df_first, df_second, df_third])

print(df_all)

# save as new excel file [overwrites without warning!]
df_all.to_excel(curr_dir + "all-shifts.xlsx", sheet_name="Shifts")
