import pandas as pd
from openpyxl import load_workbook
import pandas.io.formats.excel
import xlrd

#Funcion para agregar columna
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

# Configuramos Pandas y cargamos el archivo correspondiente (en este caso se llama archivo.xlsx)   

nombre = "..\\Excel_generados\\graduados.xlsx"

openfile = xlrd.open_workbook(nombre)


hoja_materias_graduados = openfile.sheet_by_name("listamaterias_graduados2")
hoja_materias_por_semestre_graduados = openfile.sheet_by_name("materias_graduados")

def obtenerListas(hoja_excel,lista_de_valores,lista_de_indices,lista_tipos):
    numero_valores = len(lista_de_valores)
    for i in range(hoja_excel.nrows):
        if i!= 0:
            for j in range(numero_valores):
                campo = lista_tipos[j](hoja_excel.cell_value(i,lista_de_indices[j]))
                lista_de_valores[j].append(campo)

#Variables a cargar en memoria de Creditos
codigos_materias = []
creditos_materias = []
codigos_materias_por_semestre =[]

lista_valores = [codigos_materias,creditos_materias]
lista_indices = [2,3]
lista_tipos = [str,int]

lista_valores2 = [codigos_materias_por_semestre]
lista_indices2 = [6]
lista_tipos2 = [str]

obtenerListas(hoja_materias_graduados,lista_valores,lista_indices,lista_tipos)
obtenerListas(hoja_materias_por_semestre_graduados,lista_valores2,lista_indices2,lista_tipos2)

def asignar_creditos_estudiantes(codigo_materias,creditos_materias,codigos_materias_por_semestre):
    creditos_asignados = []
    for i in range(len(codigos_materias_por_semestre)):
        for j in range(len(codigos_materias)):
            if codigos_materias_por_semestre[i]==codigos_materias[j]:
                creditos_asignados.append(creditos_materias[j])
    return creditos_asignados

creditos_estudiante = asignar_creditos_estudiantes(codigos_materias,creditos_materias,codigos_materias_por_semestre)


diccionario = {"creditos_materias":creditos_estudiante}
agrega_Columna(nombre,"materias_graduados",diccionario,8)










