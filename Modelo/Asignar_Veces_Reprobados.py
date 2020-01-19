import pandas as pd
from openpyxl import load_workbook
import pandas.io.formats.excel
import numpy as np
from Creditos_Asignados_Por_Semestre import agrega_Columna,obtenerListas
import xlrd

nombre = "../Excel_generados/graduados.xlsx"
openfile = xlrd.open_workbook(nombre)
hoja_estudiantes_graduados= openfile.sheet_by_name("estudiantes_graduados")
hoja_veces_reprobado = openfile.sheet_by_name("veces_reprobados")

#Cargar en memoria
matriculas_reprobados = []
veces_reprobados = []

lista_valores = [matriculas_reprobados,veces_reprobados]
lista_indices = [1,2]
lista_tipos = [str,int]

matriculas_estudiantes_graduados = []


lista_valores2 = [matriculas_estudiantes_graduados]
lista_indices2 = [1]
lista_tipos2 = [str]
obtenerListas(hoja_veces_reprobado,lista_valores,lista_indices,lista_tipos)
obtenerListas(hoja_estudiantes_graduados,lista_valores2,lista_indices2,lista_tipos2)

def asignarVecesReprobados(matriculas_reprobados,matriculas_estudiantes_graduados,veces_reprobados):
    valores_asignados = np.zeros(len(matriculas_estudiantes_graduados))
    for i in range(len(matriculas_estudiantes_graduados)):
        for j in range(len(matriculas_reprobados)):
            if (matriculas_estudiantes_graduados[i]==matriculas_reprobados[j]):
                valores_asignados[i]=veces_reprobados[j]

    return valores_asignados

valores_asignados = asignarVecesReprobados(matriculas_reprobados,matriculas_estudiantes_graduados,veces_reprobados)

diccionarios = {"veces_reprobado":valores_asignados}
agrega_Columna(nombre,"estudiantes_graduados",diccionarios,10)