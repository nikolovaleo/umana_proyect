import pandas as pd 
file = "test.xlsx" 
data = pd.ExcelFile(file)
fieldbus_device_sheet = data.parse("FieldbusDevice")
names_sheets = data.sheet_names
fieldbus_sheets = []
# Ciclo que toma las hojas que contienen la palabra "Fieldbus" en su nombre y las guarda en la lista vacia llamada Fieldbus_sheets
for sheet in names_sheets:
    if ('Fieldbus' in sheet) and ('OvationFieldbusPort' not in sheet) and ('FieldbusDevice' not in sheet) and ("FieldbusVCR" not in sheet):
        fieldbus_sheets.append(sheet)

FieldbusDevice_sheet = data.parse("FieldbusDevice")
#Ciclo que itera sobre los nombres de los devices y crea las hojas para cada device
for i, device_name in enumerate(FieldbusDevice_sheet["ObjectName"]):
    #Verifica si device_name esta en la columna de ObjectName en la hoja de FieldbusDevice
    new1 = data.parse("FieldbusDevice")["ObjectName"].isin([device_name])
    data1 = data.parse("FieldbusDevice")
    print(data1[new1])
    write = pd.ExcelWriter(f'data/{i+1}-{device_name}.xlsx', engine='xlsxwriter')
    archivo1 = data1[new1]
    archivo1.to_excel(write, sheet_name="FieldbusDevice", index=False)

    #Ciclo que itera sobre las hojas de Fieldbus y crea los tabs que contienen la informacion del device en las nuevas hojas de excel
    for f_sheet in fieldbus_sheets:
        new2 = data.parse(f_sheet)["ParName_11"].isin([device_name])
        data2 = data.parse(f_sheet)
        print(data2[new2])
        archivo2 = data2[new2]
        archivo2.to_excel(write, sheet_name = f_sheet, index = False)

    write.save()

