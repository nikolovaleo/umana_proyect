import pandas as pd 
#from xlsxwriter import Workbook
#import openpyxl
file = "test.xlsx" 
data = pd.ExcelFile(file)
fieldbus_device_sheet = data.parse("FieldbusDevice")
names_sheets = data.sheet_names
fieldbus_sheets = []
# Ciclo que toma las hojas que contienen la palabra "Fieldbus" en su nombre y las guarda en la lista vacia llamada Fieldbus_sheets
for sheet in names_sheets:
    if ('Fieldbus' in sheet) and ('OvationFieldbusPort' not in sheet) and ('FieldbusDevice' not in sheet) and ("FieldbusVCR" not in sheet):
        fieldbus_sheets.append(sheet)


print(fieldbus_sheets)

first_device =  "17HPS_LIT_030A" #fieldbus_device_sheet["ObjectName"][0]



new1 = data.parse("FieldbusDevice")["ObjectName"].isin([first_device])
data1 = data.parse("FieldbusDevice")
print(data1[new1])


write = pd.ExcelWriter(f'{first_device}.xlsx', engine='xlsxwriter')


archivo1 = data1[new1]
archivo1.to_excel(write, sheet_name="FieldbusDevice", index=False)


for f_sheet in fieldbus_sheets:
    new2 = data.parse(f_sheet)["ParName_11"].isin([first_device])
    data2 = data.parse(f_sheet)
    print(data2[new2])
    archivo2 = data2[new2]
   
    archivo2.to_excel(write, sheet_name = f_sheet, index = False)

write.save()

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
