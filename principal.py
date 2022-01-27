from creacionArchivoConciliacion import *
from calculoCincominutalMDA import *
from extraccionPuntoBase import *
from calculoEnergiaMTR import *
from extraccionWeb import *
from subtotalesPML import *
from openpyxl import load_workbook
from pathlib import Path


# Hoja de trabajo de operación, si no existe, se crea
file = r"conciliacion.xlsx"
fileObj = Path(file)
if fileObj.is_file() == False:
    creacionArchivoFinal()
filesheet = "conciliacion.xlsx"
 
fecha = input("Ingresa la fecha de la conciliacion: ")
    
wb = load_workbook(filesheet)
sheet = wb.worksheets[0]

#Llenado de fecha y hora, columnas A y B
a=1
b=1
celda = sheet.cell(row=a, column=b)
while celda.value:
    a=a+1
    celda = sheet.cell(row=a, column=b)
        
for i in range (1,25):
    sheet.cell(row=a, column=b).value = fecha
    sheet.cell(row=a, column=b+1).value=i
    a+=1
    b=1
wb.save('conciliacion.xlsx')
#Lleenado de columna C, valores cincominutales ya calculados a Megawatts/Hora
arch = input("Seleccione archivo cincominutal MDA no generado:")
calculoCM(arch, 3)

#Llenado de columna D, punto base
arch = input("Seleccione archivo de energía asignada:")
puntoBase(arch)

#Llenado de columna E
restaColumnas(2,3,5) #Por el uso de diferentes librerias, para la lectura las columnas se cuentan desde el numero 0 y para la escritura desde el numero 1

#Llenado de columna F
arch = input("Seleccione archivo cincominutal MTR generado:")
calculoCM(arch, 6)

#Llenado de columna G
restaColumnas(5,2,7)


extraccion()

multColumnas(3, 4, 7, 11, 16, 20)


archivo = "conciliacion.xlsx"
sheet= pd.read_excel(io = archivo, sheet_name="Precio", header = None, usecols = [0])
a = len(sheet)
filesheet = "conciliacion.xlsx"
wb = load_workbook(filesheet)
hojaEscritura = wb.worksheets[2]
hojaEscritura.cell(row=3, column=2).value = str(a)
wb.save('conciliacion.xlsx')















