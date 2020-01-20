from Creditos_Asignados_Por_Semestre import agrega_Columna,obtenerListas
from pandas import ExcelWriter
import xlrd
from openpyxl import load_workbook
import numpy as np
#Memoria
ruta = "..\\Excel_generados\\graduados.xlsx"
ruta2 = "..\\Excel_generados\\materias_malla.xlsx"

archivo_graduados = xlrd.open_workbook(ruta)
archivo_malla = xlrd.open_workbook(ruta2)

hoja_estudiantes_graduados = archivo_graduados.sheet_by_name("estudiantes_graduados")
hoja_primer_semestre= archivo_graduados.sheet_by_name("primer_semestre")
hoja_malla = archivo_malla.sheet_by_name("sistemas_multimedia")

#Hoja de materias del primer semestre
matriculas_primer_semestre = []
codigos_primer_semestre = []


#Hoja de la malla
codigos_malla =[]
semestre_malla = []

#Hoja de las Estudiantes Graduados
matriculas_graduados =[]



#Cargando a memoria Malla
valores = [codigos_malla,semestre_malla]
indices = [1,3]
tipos =[str,int]
obtenerListas(hoja_malla,valores,indices,tipos)

#Cargando a memoria primer semestre
valores = [matriculas_primer_semestre,codigos_primer_semestre]
indices = [1,6]
tipos = [str,str]
obtenerListas(hoja_primer_semestre,valores,indices,tipos)
#Cargando a memoria estudiantes_graduados
valores = [matriculas_graduados]
indices = [1]
tipos =[str]

obtenerListas(hoja_estudiantes_graduados,valores,indices,tipos)

def materias_malla_semestre(numero_semestre,codigos,semestres):
    codigos_semestre = []
    for i in range(len(codigos)):
        if(semestres[i]== numero_semestre):
            codigos_semestre.append(codigos[i])
    return codigos_semestre

malla_primer_semestre = materias_malla_semestre(1,codigos_malla,semestre_malla)

def asignar_carrera(mallaprimer,matriculasestudiante,matriculaprimer,codprimer):
    otra_carrera = np.ones(len(matriculasestudiante))
    for i in range(len(matriculasestudiante)):
        for j in range(len(matriculaprimer)):
            if(matriculasestudiante[i].split(".")[0] == matriculaprimer[j]):
                if(codprimer[j] in mallaprimer):
                    otra_carrera[i]= 0

    return  otra_carrera

o = asignar_carrera(malla_primer_semestre,matriculas_graduados,matriculas_primer_semestre,codigos_primer_semestre)
diccionario = {"Otra_Carrera":o}
agrega_Columna(ruta,"estudiantes_graduados",diccionario,11)

