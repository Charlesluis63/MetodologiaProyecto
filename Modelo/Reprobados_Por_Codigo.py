import pandas as pd
from openpyxl import load_workbook
import pandas.io.formats.excel
from Creditos_Asignados_Por_Semestre import obtenerListas,agrega_Columna
import numpy as np
import xlrd
#Variables a cargar en memoria

codigos_todas_materias_graduados = []

codigos_materias_por_semestre = []

estado_materias = []

nombre = "../Excel_generados/graduados.xlsx"
openfile = xlrd.open_workbook(nombre)
hoja_materias_por_semestre= openfile.sheet_by_name("materias_graduados")
hoja_todas_materias_graduados = openfile.sheet_by_name("listamaterias_graduados2")


lista_valores =[codigos_todas_materias_graduados]
lista_indices =[2]
lista_tipos=[str]
obtenerListas(hoja_todas_materias_graduados,lista_valores,lista_indices,lista_tipos)

lista_valores =[codigos_materias_por_semestre,estado_materias]
lista_indices =[6,3]
lista_tipos=[str,str]
obtenerListas(hoja_materias_por_semestre,lista_valores,lista_indices,lista_tipos)

reprobados = np.zeros(len(codigos_todas_materias_graduados))
aprobados = np.zeros(len(codigos_todas_materias_graduados))
por_faltas = np.zeros(len(codigos_todas_materias_graduados))

for m in range(len(codigos_todas_materias_graduados)):
    for x in range(len(codigos_materias_por_semestre)):
        if codigos_todas_materias_graduados[m] == codigos_materias_por_semestre[x]:
            if estado_materias[x]== 'RP':
                reprobados[m]+= 1
            elif estado_materias[x]== 'AP':
                aprobados[m]+=1
            else:
                por_faltas[m]+=1

diccionario = {"aprobados":aprobados,"reprobados":reprobados,"reprobados_por_faltas":por_faltas}
agrega_Columna(nombre,"listamaterias_graduados2",diccionario,4)

vale = len(codigos_materias_por_semestre)
print(aprobados.sum() +por_faltas.sum()+ reprobados.sum() == vale)