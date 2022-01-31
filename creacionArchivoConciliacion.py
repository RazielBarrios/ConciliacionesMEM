import xlsxwriter


def creacionArchivoFinal():
    a=0
    data1 = ["Fecha", "Hora", "Medición MDA \n (MWh)", "Energía MDA \n (MWh)", "Energía MTR \n (MWh)", "Generación MTR" + '\n' + "(MWh)", "Cargas\n(MWh)", 
            "PML MDA", "Energía MDA", "Pérdida MDA", "Congestión MDA", "PML MTR", "Energía MTR", "Pédida MTR", "Congestión MTR", "Importe PML MDA",
            "Importe energía MDA", "Importe pédidas MDA", "Importe congestión MDA", "Importe PML MTR", "Importe energía MTR", "Importe pérdida MTR",
            "Importe congestión MTR"]
    
    data2 = ["Fecha", "Medicion MWh", "Venta Energía MDA", "Importe PML MDA", "Importe de Energía MDA", "Importe de Pérdidas", "Importe de Congestión", 
            "Compra Energía MTR", "Importe PML MTR", "Importe de Energía MTR", "Importe de Pérdidas", "Importe de Congestión", "Operación CENACE",
            "Transmisión", "Servicios Conexos", "Medición MWh", "Energía MTR", "Medición MWh", "Operación CENACE", "Transmisión", "Nota de Débito",
            "Nota de Crédito", " ", "Subtotal Egreso", "Oferta Energía MDA", "Precio Energía MDA", "Energía MTR", "Precio Energía MTR", "Ingreso Subtotal",
            "Nota de Débito", "Nota de Crédito"]
    
    data3 = ["1a. Reliquidación", "2a. Reliquidación", "3a. Reliquidación", "Reliquidación por Controversia"]
    
    libro = xlsxwriter.Workbook('conciliacion.xlsx') 
    hoja1 = libro.add_worksheet('Precio')
    hoja2 = libro.add_worksheet('Resumen')
    hoja3 = libro.add_worksheet('Tarifas')
    hoja4 = libro.add_worksheet('NotasDC')
    
    #Formato de hoja 1
    for a in range(0, len(data1)):
        hoja1.write(0, a, data1[a])
        print(data1[a])
    formato = libro.add_format({'bold': True, 'font_color': 'red', 'center_across': True})    
    hoja1.set_column('A:W', 25, formato)
    
    #Formato de hoja 2
    for a in range(0, len(data2)):
        hoja2.write(0, a, data2[a])
        print(data2[a])
    formato = libro.add_format({'bold': True, 'font_color': 'green', 'center_across': True})    
    hoja2.set_column('A:AE', 25, formato)
    
    #Formato de hoja 3
    hoja3.write(0,0, "Costo por MWh")
    hoja3.write(1,0, "Cobro por MWh")
    hoja3.write(2,0, "Ultima Celda")
    hoja3.write(0,1, "$ 3.28")
    hoja3.write(1,1, "$ 104.70")
    hoja3.write(2,1, "0")
    formato = libro.add_format({'bold': True, 'font_color': 'blue', 'center_across': True})    
    hoja3.set_column('A:B', 25, formato)
    
    #Formato de hoja 4
    formato = libro.add_format({
                   'bold':     True,
                   'border':   6,
                   'align':    'center',#Center horizontalmente
                   'valign':   'vcenter',# Centro vertical
                   'font_color': 'orange',
                   })
    
    hoja4.merge_range('B1:I1', 'CENACE', formato)
    hoja4.merge_range('J1:Q1', 'PARTICIPANTE', formato)
    
    hoja4.merge_range('B2:E2', 'NOTAS DE DÉBITO', formato)
    hoja4.merge_range('F2:I2', 'NOTAS DE CRÉDITO', formato)
    hoja4.merge_range('J2:M2', 'NOTAS DE DÉBITO', formato)
    hoja4.merge_range('N2:Q2', 'NOTAS DE CRÉDITO', formato)
    
    hoja4.write(0, 0, '-', formato)
    hoja4.write(1, 0, '-', formato)
    hoja4.write(2, 0, 'Fecha', formato)
    b = 0
    for i in range (0,4):
        for a in range(0, 4):
            hoja4.write(2, b+1, data3[a], formato)
            b += 1
    
    formato = libro.add_format({'bold': True, 'font_color': 'orange', 'center_across': True}) 
    hoja4.set_column('A:Q', 35, formato)
    
    libro.close()


#creacionArchivoFinal()

