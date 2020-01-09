import pandas as pd
from openpyxl import load_workbook
import pandas.io.formats.excel
import xlrd
#Variables a cargar en memoria
materias = []
materias_estudiantes = []
estado_materias = []
nombre = "../Excel_generados/graduados.xlsx"
openfile = xlrd.open_workbook(nombre)
hoja_graduados = openfile.sheet_by_name("materias_graduados")
hoja_materias = openfile.sheet_by_name("listamaterias_graduados")

for i in range(hoja_materias.nrows):
    if (i!=0) :
        mat = (hoja_materias.cell_value(i,1))
        materias.append(mat)

for j in range(hoja_graduados.nrows):
    if j!= 0:
        nombre_materia = hoja_graduados.cell_value(j,2)
        estados = hoja_graduados.cell_value(j,3)
        materias_estudiantes.append(nombre_materia)
        estado_materias.append(estados)
