import pandas as pd
from openpyxl import load_workbook
import pandas.io.formats.excel
from Creditos_Asignados_Por_Semestre import obtenerListas,agrega_Columna
import numpy as np
import xlrd
#Variables a cargar en memoria
materias = []
materias_estudiantes = []
estado_materias = []
nombre = "../Excel_generados/graduados.xlsx"
openfile = xlrd.open_workbook(nombre)
hoja_graduados = openfile.sheet_by_name("materias_graduados")
hoja_materias = openfile.sheet_by_name("listamaterias_graduados2")

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
reprobados = np.zeros(len(materias))
aprobados = np.zeros(len(materias))
por_faltas = np.zeros(len(materias))

for m in range(len(materias)):
    for x in range(len(materias_estudiantes)):
        if materias[m] == materias_estudiantes[x]:
            if estado_materias[x]== 'RP':
                reprobados[m]+= 1
            elif estado_materias[x]== 'AP':
                aprobados[m]+=1
            else:
                por_faltas[m]+=1

diccionario = {"aprobados":aprobados,"reprobados":reprobados,"reprobados_por_faltas":por_faltas}
agrega_Columna(nombre,"listamaterias_graduados",diccionario,3)

vale = len(materias_estudiantes)
print(aprobados.sum() +por_faltas.sum()+ reprobados.sum() == vale)