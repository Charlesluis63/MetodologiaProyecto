import pandas as pd
from openpyxl import load_workbook
from pandas import ExcelWriter
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

#Leemos el archivo de los alumnos graduados, la hoja que contiene las materias de los alumnos
archivoGraduados = "..\\Excel_generados\\graduados.xlsx"
openfileA = xlrd.open_workbook(archivoGraduados)
hojaGraduados = openfileA.sheet_by_name("materias_graduados")

semestreAlumno = []
semestre_estudiante=1
for i in range(hojaGraduados.nrows):
    if i>0:
        matricula = str((hojaGraduados.cell_value(i, 1)))
        matricula_anterior = str((hojaGraduados.cell_value(i-1, 1)))
        #print("Posicion",i," Matricula actual " + matricula + "-- Matricula anterior " + matricula_anterior)

        if matricula == matricula_anterior or matricula_anterior=="matricula":
            semestre = str((hojaGraduados.cell_value(i, 4)))+"-"+str((hojaGraduados.cell_value(i, 5)))
            semestre_anterior = str((hojaGraduados.cell_value(i-1, 4)))+"-"+str((hojaGraduados.cell_value(i-1, 5)))
            #print("Posicion", i, " Semestre actual " + semestre + "-- Semestre anterior " + semestre_anterior)

            if semestre == semestre_anterior or semestre_anterior=="año-termino":
                info_semestre = str(semestre_estudiante)+" semestre"
                semestreAlumno.append(info_semestre)
                print("Posicion",i," Matricula actual " + matricula + "-- Semestre actual " + semestre+" Se encuentra en su ",semestre_estudiante," Semestre")
            else:
                semestre_estudiante = semestre_estudiante+1
                info_semestre = str(semestre_estudiante)+" semestre"
                semestreAlumno.append(info_semestre)
                print("Posicion", i," Matricula actual " + matricula + "-- Semestre actual " + semestre + " Se encuentra en su ",semestre_estudiante, " Semestre")
        else:
            semestre_estudiante = 1
            info_semestre = str(semestre_estudiante) + " semestre"
            semestreAlumno.append(info_semestre)
            print("Posicion", i," Matricula actual " + matricula + "-- Semestre actual " + semestre + " Se encuentra en su ",semestre_estudiante, " Semestre")

diccionario = {"Semestre_materia": semestreAlumno}
print(semestreAlumno)

agrega_Columna(archivoGraduados,"materias_graduados",diccionario,7)


