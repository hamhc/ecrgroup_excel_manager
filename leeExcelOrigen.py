from openpyxl import load_workbook

#Asignamos la ruta donde estará el archivo
path = '/Users/johamhernandez/Documents/Proyectos Python/ECR_Group/Proyecto/Planillas/p1_sample.xlsx'

#Leemos el archivo
wb = load_workbook(filename = path)

#Leemos todas las columnas de la hoja indicada
#TODO: Asignar el nombre de la hoja como parámetro
#sheet_columns = wb['Hoja1'].columns
sheet_rows = wb['Hoja1'].max_row

#Recorremos la hoja para obtener el encabezado de cada columna

# for i in sheet_columns:
# 	for j in range(sheet_rows):
# 		print(i[j].value)

#Recorremos la hoja de acuerdo a los valores de cada fila y cada columna
for j in range(sheet_rows):
	#i = 0
	sheet_columns = wb['Hoja1'].columns
	for i in sheet_columns:
		print(i[j].value)


#TODO: Guardar los resultados para pasarlos a la creación de un excel
#Es importante evaluar el campo Estatus ya que ese es el que indicará
#Si de debe cargar o no 

# for i in range(1, sheet_ranges + 1):
#     cell_obj = sheet_ranges.cell(row = 1, column = i)
#     print(cell_obj.value)


# for i in sheet_ranges:
# 	ch = 'A'
# 	x = chr(ord(ch) + i)
# 	print(x)


