from urllib import request
import xml.etree.ElementTree as ET
from openpyxl import load_workbook

def extraccion():
    
    url1 = "https://ws01.cenace.gob.mx:8082/SWPML/SIM/SIN/MDA/01VAJ-230/2021/11/26/2021/11/26/XML"
    url2 = "https://ws01.cenace.gob.mx:8082/SWPML/SIM/SIN/MTR/01VAJ-230/2021/11/26/2021/11/26/XML"

    """
    sistema = input("Ingrese el tipo de sistema: ")
    url = url + sistema + "/"
    proceso = input("Ingrese el tipo de proceso: ")
    url = url + proceso + "/"
    lista_nodos = input("Ingrese el nodo: ")
    url = url + lista_nodos + "/"
    anio_ini = input("Indique el año inicial a consultar: ")
    url = url + anio_ini + "/"
    mes_ini = input("Indique el mes inicial a consultar: ")
    url = url + mes_ini + "/"
    dia_ini = input ("Indique el dia inicial a consultar: ")
    url = url + dia_ini + "/"
    anio_fin = input("Indique el año final a consultar: ")
    url = url + anio_fin + "/"
    mes_fin = input("Indique el mes final a consultar: ")
    url = url + mes_fin + "/"
    dia_fin = input("Indique el dia final a consultar: ")
    url = url + dia_fin + "/"
    formato = input("Formato de salida: ")
    url = url + formato 
    print(url)"""
    
    html = request.urlopen(url1)
    data = html.read()
    tree = ET.fromstring(data)
    results = tree.findall('Resultados/Nodo/Valores/Valor')
    
    filesheet = "conciliacion.xlsx"
    wb = load_workbook(filesheet)
    sheet = wb.worksheets[0]
    a=1
    b=8
    palabra = ""
    for b in range (8,12):
        celda = sheet.cell(row=a, column=b)
        while celda.value:
            a=a+1
            celda = sheet.cell(row=a, column=b)
        if b == 8:
            palabra = 'pml'
        if b == 9:
            palabra = 'pml_ene'
        if b == 10:
            palabra = 'pml_per'
        if b == 11:
            palabra = 'pml_cng'
        for result in results:
            sheet.cell(row=a, column=b).value = result.find(palabra).text
            a += 1
        a = 1
    wb.save('conciliacion.xlsx')
    
    html = request.urlopen(url2)
    data = html.read()
    tree = ET.fromstring(data)
    results = tree.findall('Resultados/Nodo/Valores/Valor')
    filesheet = "conciliacion.xlsx"
    wb = load_workbook(filesheet)
    sheet = wb.worksheets[0]
    a=1
    b=12
    palabra = ""
    for b in range (12,16):
        celda = sheet.cell(row=a, column=b)
        while celda.value:
            a=a+1
            celda = sheet.cell(row=a, column=b)
        if b == 12:
            palabra = 'pml'
        if b == 13:
            palabra = 'pml_ene'
        if b == 14:
            palabra = 'pml_per'
        if b == 15:
            palabra = 'pml_cng'
        for result in results:
            sheet.cell(row=a, column=b).value = result.find(palabra).text
            a += 1
        a = 1
    
    wb.save('conciliacion.xlsx')
        
#extraccion()