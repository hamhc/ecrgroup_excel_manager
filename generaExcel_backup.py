##Generación de archivo excel

import openpyxl, os

#Crea nuevo archivo excel

wb = openpyxl.Workbook()

#Obtenemos cual es la hoja Activa del archivo
active = wb.active

#Para hacerlo con el nombre de una hoja, podemos seleccionar
#wb2 = wb['DEPOR']
#wb2.title = 'nuevonombre' 

#Asignamos el nombre a la hoja generada
#TODO = Parametro del nombre
active.title = 'DEPOR'

#print(active)


#Asignamos la hoja Deport (del libro) a una variable
sheet = wb['DEPOR']
#Asignamos valores a las celdas A1 y A2
sheet['A1	'] = 42
sheet['A2'] = 'Caraotas'

#Nos cambiamos de directorio para guardar el archivo
##os.chdir('/Users/johamhernandez/Documents/Proyectos Python/Practicas Excel')
os.chdir('/Users/johamhernandez/Documents/Proyectos Python/ECR_Group/Proyecto/Results')


#Guardamos el achivo excel después de cambiarnos de ruta
#NOTA: si no lo guardamos, solo existe en memoria

wb.save('archivo.xlsx')

#Para agregar otra hoja al libro, se agrega el titulo opcionalmente

#sheet2 = wb.create_sheet('Hoja2')

#Para cambiar titulo de hoja

#sheet2.title = 'NuevoTitulo'


wb.save('archivo.xlsx')