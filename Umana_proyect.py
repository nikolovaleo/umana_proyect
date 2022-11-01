import pandas as pd 
file = "test.xlsx" 
data = pd.ExcelFile(file)
fieldbus_device_sheet = data.parse("FieldbusDevice")
names_sheets = data.sheet_names
fieldbus_sheets = []

for sheet in names_sheets:
    if ('Fieldbus' in sheet) and ('OvationFieldbusPort' not in sheet) and ('FieldbusDevice' not in sheet):
        fieldbus_sheets.append(sheet)

first_device =  "17HPS_LIT_030A" #fieldbus_device_sheet["ObjectName"][0]
#for data_sheet in fieldbus_sheets:

new = data.parse("FieldbusDevice")["ObjectName"].isin([first_device])
data1 = data.parse("FieldbusDevice")
print(data1[new])

archivo = data1[new]
archivo.to_excel("output.xlsx", sheet_name="FieldbusDevice")

#column_names = []
#for sheet in fieldbus_names:
#    for col in sheet.columns:
#        column_names.append(col)


#device_df = pd.DataFrame(columns = column_names)
#print(data.sheet_names)






# Para ver el contenido de una de las hojas se utiliza la siguiente sintaxis: X_sheet = data.parse("Nombre de la hoja en excel")
# Para accesar a una columna de una hoja en especifico: X_sheet["Nombre de la columna en la hoja"]
# Para accesar a un elemento de una celda en una hoja y columna especifica: X_sheet["Nombre de la columna"][posicion numerada desde 0 hasta n]
# Para conocer el tamano de una columna en especifico X_sheet["Columna"].size
