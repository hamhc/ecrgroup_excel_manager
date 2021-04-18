from openpyxl import load_workbook
from generaExcel import genera_archivo

#TODO: Ruta que venga desde un archivo
path = '/Users/johamhernandez/Documents/Proyectos Python/ECR_Group/Proyecto/Planillas/p1_sample.xlsx'

#Leemos el archivo
wb = load_workbook(filename = path)

#asignamos contenido de hoja por nombre
hoja = wb['DEPOR']

#Buscamos valor del campo estatus
#estatus = hoja['Z3']
#print(estatus.value)

#Obtenemos las celdas que tienen data en el excel
#rango_celdas = hoja['A1':'Z4']

#Obtenemos total de columnas
sheet_columns_qty = hoja.max_column
#Obtenemos total de filas
sheet_rows_qty = hoja.max_row

#Asignamos al rango desde A1 hasta Z + el total de filas que hay en la hoja
rango_celdas = hoja['A1':'Z'+str(sheet_rows_qty)]

encabezado = []
#contenido_tmp = []
contenido = []

#print(sheet_rows_qty)
#print(sheet_columns_qty)
#print(rango_celdas)
#print(rango_celdas[0][0].value)

for j in range(0,sheet_rows_qty):

	contenido_tmp = []
	for i in range(0,sheet_columns_qty):
		#primer valor = Filas
		#segundo valor = columnas
		#print(rango_celdas[0][i].value)	

		#print(rango_celdas[j][sheet_columns_qty-1].value)
		valor = str(rango_celdas[j][sheet_columns_qty-1].value)
		#print(valor)

		if(j == 0):
			#print(rango_celdas[0][i].value)
			encabezado.append(rango_celdas[j][i].value)

		elif(j > 0):
			if not ((valor.upper()) == 'SUSPENDIDO'):
				#print(rango_celdas[j][i].value)
				contenido_tmp.append(rango_celdas[j][i].value)

	if(j > 0):
		contenido.append(contenido_tmp)

#print(encabezado)
#print(contenido)


#Iniciamos contador para saber cuando no estamos leyendo el encabezado
# counter = -1
# for i in rango_celdas:
# 	counter += 1
# 	print('counter is ' + str(counter))
# 	#print(i[sheet_columns_qty-1].value)

# 	#Obtenemos el valor del campo Estatus de la hoja
# 	valor = str(i[sheet_columns_qty-1].value)
	
# 	if(counter == 1):
# 		print("Entró en counter")
# 		for k in range(0,sheet_columns_qty):
# 				print("K es: "+ str(k))
# 				print(i[k].value)

# 	if(counter > 0):
# 		if not((valor.upper()) == 'SUSPENDIDO'):
# 			#Consideramos solo los distintos a suspendido
# 			for j in range(0,sheet_columns_qty):
# 				print(i[j].value)

genera_archivo("DEPOR", "Prueba1", encabezado, contenido)