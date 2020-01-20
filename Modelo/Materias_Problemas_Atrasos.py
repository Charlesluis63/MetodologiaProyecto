from Creditos_Asignados_Por_Semestre import agrega_Columna,obtenerListas
from pandas import ExcelWriter
import xlrd
from openpyxl import load_workbook
#Memoria
ruta = "..\\Excel_generados\\graduados.xlsx"
ruta2 = "..\\Excel_generados\\materias_malla.xlsx"

archivo_graduados = xlrd.open_workbook(ruta)
archivo_malla = xlrd.open_workbook(ruta2)

hoja_materias_problema = archivo_graduados.sheet_by_name("estudiantes_materiasProblema")
hoja_malla = archivo_malla.sheet_by_name("sistemas_multimedia")
hoja_estudiantes = archivo_graduados.sheet_by_name("estudiantes_graduados")


#Hoja de la malla
codigos_malla =[]
semestre_malla = []

#Hoja de las Materias Problemas
matriculas_graduados =[]
semestre_graduados = []
termino_graduados =[]
codigos_materias_graduados = []

#Hoja de estudiantes Graduados
matriculas_unicas = []
otra_carrera = []

#Cargando a memoria Malla
valores = [codigos_malla,semestre_malla]
indices = [1,3]
tipos =[str,int]
obtenerListas(hoja_malla,valores,indices,tipos)

#Cargando a memoria Materias_Problema
valores = [matriculas_graduados,semestre_graduados,codigos_materias_graduados,termino_graduados]
indices = [1,7,6,5]
tipos= [str,int,str,str]
obtenerListas(hoja_materias_problema,valores,indices,tipos)

#Cargando a memoria Estudiantes
valores=[matriculas_unicas,otra_carrera]
indices = [1,11]
tipos = [str,int]
obtenerListas(hoja_estudiantes,valores,indices,tipos)
mat_otra_carrera = {}

for i in range(len(matriculas_unicas)):
    mu = matriculas_unicas[i].split(".")[0]
    if mu not in mat_otra_carrera:
        mat_otra_carrera[mu]=otra_carrera[i]


#AUN FALTA VERIFICAR QUE NO VENGA DE OTRA CARRERA

def obtener_Atrasos(matricula,semgrad,semmalla,codigograd,codigomalla,termgrad,diccionario):
    atrasos = []
    for i in range(len(matricula)):
        for j in range(len(codigomalla)):
            if(codigograd[i]==codigos_malla[j]):
                if(semgrad[i] == semmalla[j]):
                    atrasos.append("no")
                elif(semgrad[i]<semmalla[j] and diccionario[matricula[i]]!= 1):
                    atrasos.append("adelantado")
                else:
                    atrasos.append("si")
    return  atrasos
atrasos = obtener_Atrasos(matriculas_graduados,semestre_graduados,semestre_malla,codigos_materias_graduados,codigos_malla,termino_graduados,mat_otra_carrera)

diccionario = {"atrasado":atrasos}
agrega_Columna(ruta,"estudiantes_materiasProblema",diccionario, 8)






