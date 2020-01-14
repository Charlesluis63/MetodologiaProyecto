import pandas as pd
from openpyxl import load_workbook
from pandas import ExcelWriter
import xlrd

def agrega_Columna(nombre,nombre_hoja,valores_agregar,numero_columna):
    book = load_workbook(nombre)
    df = pd.DataFrame(valores_agregar)
    columnas = []
    for i in valores_agregar:
        columnas.append(i)
    df = df[columnas]
    writer = pd.ExcelWriter(nombre)
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    df.to_excel(writer, sheet_name=nombre_hoja, startcol=numero_columna, index=False)
    writer.save()

nombre = "..\\Datos\\MateriasTodas.xls"
openfile = xlrd.open_workbook(nombre)
hoja = openfile.sheet_by_name("Sheet1")
print("Hay "+ str(hoja.nrows) + " registros actualmente")

nombreM = "..\\Excel_generados\\listado_materias.xlsx"
openfile = xlrd.open_workbook(nombreM)
hojaM = openfile.sheet_by_name("materias")
print("Hay "+ str(hojaM.nrows) + " registros actualmente")
lista_creditos = []

contador = 0
for i in range(hojaM.nrows):
    codigoMat = str((hojaM.cell_value(i, 2)))
    nombreMat = str((hojaM.cell_value(i, 1)))

    for j in range(hoja.nrows):
        codigo = str((hoja.cell_value(j, 0)))
        creditos = str( (hoja.cell_value(j, 10)))  # En caso que sean teorica y practica, se suma con hoja.cell_value(i,11)
        creditos = creditos.split(".")
        if codigoMat == codigo:
            print(codigoMat+" -- nombre:"+nombreMat+" -- creditos:"+creditos[0])
            lista_creditos.append(creditos[0])
            contador = contador +1
print(contador)

diccionario = {"Creditos": lista_creditos}

agrega_Columna(nombreM,"materias",diccionario,3)