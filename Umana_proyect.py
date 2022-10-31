import pandas as pd 
file = "test.xlsx" 
data = pd.ExcelFile(file)
drop_sheet = data.parse("Drop")
unit_sheet = data.parse("Unit")

print(data.sheet_names)



# Para ver el contenido de una de las hojas se utiliza la siguiente sintaxis: X_sheet = data.parse("Nombre de la hoja en excel")
# Para accesar a una columna de una hoja en especifico: X_sheet["Nombre de la columna en la hoja"]
# Para accesar a un elemento de una celda en una hoja y columna especifica: X_sheet["Nombre de la columna"][posicion numerada desde 0 hasta n]
# Para conocer el tamano de una columna en especifico X_sheet["Columna"].size