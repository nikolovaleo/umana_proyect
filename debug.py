import pandas as pd 
file = "test.xlsx" 
data = pd.ExcelFile(file)

names_sheets = data.sheet_names

#print(names_sheets)


fieldbus_names = []

for sheet in names_sheets:

    if 'Fieldbus' in sheet and 'OvationFieldbusPort' not in sheet:

        fieldbus_names.append(sheet)


print(fieldbus_names)