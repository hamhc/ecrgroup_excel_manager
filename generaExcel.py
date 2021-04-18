##Generación de archivo excel

import openpyxl, os

#Crea nuevo archivo excel

def genera_archivo(nombre_cliente, nombre_archivo, encabezado, cuerpo):
	"""Creamos archivo"""
	wb = openpyxl.Workbook()

	#Obtenemos cual es la hoja Activa del archivo
	active = wb.active

	#Asignamos el nombre a la hoja generada
	active.title = nombre_cliente

	#Asignamos la hoja Deport (del libro) a una variable
	sheet = wb[nombre_cliente]
	#Asignamos valores a las celdas A1 y A2
	#sheet['A1'] = 42
	#sheet['A2'] = 'Caraotas'
	#for i in encabezado:
	sheet.append(encabezado)
	for i in cuerpo:
		sheet.append(i)

	guarda_archivo(nombre_archivo, '/Users/johamhernandez/Documents/Proyectos Python/ECR_Group/Proyecto/Results', wb)


def guarda_archivo(nombre_archivo, path, wb_object):
	"""Indicamos ruta, nombre y guardamos el archivo"""

	#Nos cambiamos de directorio para guardar el archivo
	##os.chdir('/Users/johamhernandez/Documents/Proyectos Python/Practicas Excel')
	#os.chdir('/Users/johamhernandez/Documents/Proyectos Python/ECR_Group/Proyecto/Results')
	os.chdir(path)
	#Guardamos el achivo excel después de cambiarnos de ruta
	#NOTA: si no lo guardamos, solo existe en memoria

#	wb.save('archivo.xlsx')
	wb_object.save(nombre_archivo +'.xlsx')



#encabezadoC = ['Planta', 'Unidad Administrativa', 'N°', 'Rut', 'Dv', 'Rut', 'A.Paterno', 'A.Materno', 'Nombres', 'F.Ingreso', 'Cargo_Trabajador', 'Dirección', 'Comuna', 'Región', 'Trabajador', 'Sucursal', 'NO PUEDE TRABAJAR', 'Banco', 'Numero de Cuenta', 'Tipo de cuenta', 'Fecha de comunicación', 'Inicio Suspensión', 'Correo Electronico', 'Teléfono', 'ID SUSPENSIÓN MAR2021', 'Estatus']
#cuerpo = [['COMERCIAL DEPOR ', 'UMBRO', 167, 18532277, '7', '18532277-7', 'URRUTIA', 'SAEZ', 'ABIGAIL FRANCISCA', 43238, 'PROMOTOR (A)                            ', 'JUANA WEBER N 4864', 'ESTACION CENTRAL                  ', 'SANTIAGO', 'ABIGAIL FRANCISCA URRUTIA SAEZ', 'SANTIAGO                                ', None, 'Banco del Estado de Chile     ', '18532277                            ', 'CUENTA RUT ', None, None, 'francisca.urrutia23@gmail.com>', 996830288, 'LEY CRIANZA PROTEGIDA 14-01-21 al 30-04-21', None], [], [], [None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None], [None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None], [None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None], [None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None], [None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None], [None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None]]

#genera_archivo("DEPOR", "Prueba1", encabezadoC, cuerpo)