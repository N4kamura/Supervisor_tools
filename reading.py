from openpyxl import load_workbook
import numpy as np
import os

def read_excel_peatonal(excel_path): #define la lectura del excel, colocación a 0 de typo flotante
    #global hojas #invoca variables fuera de la función
    wb = load_workbook(excel_path,read_only=True,data_only=True)
    data =[]
    ws = wb["Data Peatonal"]
    filas = ws["L19:UY19"]
    valores = [celda.value for fil in filas for celda in fil]
    data_hoja = [[celda.value for celda in fila] for fila in ws['L20:UY79']] #fila-celda SE CAMBIA EL RANGO
    try:
        matriz_numpy = np.array(data_hoja,dtype="float") #conversión a np
    except ValueError:
        print("Hay un valor que no es número en la base de datos")
        _, excel_name = os.path.split(excel_path)
        print(f"Excel = {excel_name}")
        print(f"Hoja = Data Peatonal")
    matriz_numpy[np.isnan(matriz_numpy)] = 0 #0 0 a vacios
    data.append(matriz_numpy)

    ws = wb["Inicio"] #guarda fecha
    fecha = str(ws['G5'].value) #Por cada excel

    return data,fecha,valores

def read_excel_vehicular(excel_path):
    hojas = ["N","S","E","O"]
    wb = load_workbook(excel_path,read_only=True,data_only=True)
    data=[]
    val =[]
    for hoja in hojas:
        ws = wb[hoja]
        filas = ws["K15:HB15"]
        valores = [celda.value for fil in filas for celda in fil] #Encabezado
        val.append(valores)
        data_hoja = [[celda.value for celda in fila] for fila in ws['K16:HB111']] #fila-celda
        try:
            matriz_numpy = np.array(data_hoja,dtype="float") #conversión a np
        except ValueError:
            print("Hay un valor que no es número en la base de datos")
            _,excel_name = os.path.split(excel_path)
            print(f"Excel = {excel_name}")
            print(f"Hoja = {hoja}")
            return print("^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^\n################### PROCESO DETENIDO - ERROR #####################")
        matriz_numpy[np.isnan(matriz_numpy)] = 0 #0 0 a vacios
        data.append(matriz_numpy)

    ws = wb["Inicio"] #guarda fecha
    fecha = str(ws['G8'].value) #Por cada excel

    return data,fecha,val