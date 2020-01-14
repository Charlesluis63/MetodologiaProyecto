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
nombre2 = "..\\Excel_generados\\listado_materias.xlsx"
openfile = xlrd.open_workbook(nombre)
openfile2 = xlrd.open_workbook(nombre2)
hoja_materias_graduados = openfile.sheet_by_name("listamaterias_graduados2")
hojas_creditos_todas_materias = openfile2.sheet_by_name("materias")

def obtenerListas(hoja_excel,lista_de_valores,lista_de_indices,lista_tipos):
    numero_valores = len(lista_de_valores)
    for i in range(hoja_excel.nrows):
        if i!= 0:
            for j in range(numero_valores):
                campo = lista_tipos[j](hoja_excel.cell_value(i,lista_de_indices[j]))
                lista_de_valores[j].append(campo)

#Variables a cargar en memoria de Creditos
codigos_materias_con_creditos =[]
creditos_todas_materias = []
lista_valores = [creditos_todas_materias,codigos_materias_con_creditos]
lista_indices = [3,2]
lista_tipos = [int,str]

obtenerListas(hojas_creditos_todas_materias,lista_valores,lista_indices,lista_tipos)

codigos_materias = []
lista_valores = [codigos_materias]
lista_indices = [2]
lista_tipos = [str]

obtenerListas(hoja_materias_graduados,lista_valores,lista_indices,lista_tipos)



def asignar_creditos(codigos_sin_creditos,codigos_con_creditos,creditos):
    creditos_asignados = []
    for c in codigos_sin_creditos :
        for cod in range(len(codigos_con_creditos)):
            if(c == codigos_con_creditos[cod]):
                creditos_asignados.append(creditos[cod])
    return creditos_asignados

creditos_asignados = asignar_creditos(codigos_materias,codigos_materias_con_creditos,creditos_todas_materias)
diccionario = {"creditos_materias":creditos_asignados}
agrega_Columna(nombre,"listamaterias_graduados2",diccionario,3)
#creditos_estudiante = asignar_creditos_estudiantes(codigos_materias,creditos_materias,codigos_materias_por_semestre)


#diccionario = {"creditos_materias":creditos_estudiante}
#agrega_Columna(nombre,"materias_graduados",diccionario,6)










