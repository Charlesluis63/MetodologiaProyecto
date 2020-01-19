import pandas as pd
from pandas import ExcelWriter
from openpyxl import load_workbook
import xlrd

def mallasDiccionario(hoja):
    malla = {}
    materias_semestre = []
    for i in range(hoja.nrows):
        if str(hoja.cell_value(i, 0)) != "Materia":
            semestre = str(hoja.cell_value(i, 3))
            semestre_anterior = str((hoja.cell_value(i - 1, 3)))
            if semestre == semestre_anterior or semestre_anterior == "Semestre":
                materias_semestre.append(str(hoja.cell_value(i, 1)) + "," + str(hoja.cell_value(i, 2)))
            else:
                #print(materias_semestre)
                semestreT = semestre_anterior.split(".")
                semestreT = semestreT[0]
                malla[semestreT] = materias_semestre
                materias_semestre = []
                materias_semestre.append(str(hoja.cell_value(i, 1)) + "," + str(hoja.cell_value(i, 2)))
    # Agregar ultimo semestre
    semestreT = semestre_anterior.split(".")
    semestreT = semestreT[0]
    malla[semestreT] = materias_semestre
    return  malla


archivoHistorialGraduados = "..\\Excel_generados\\graduados.xlsx"
openfileHG = xlrd.open_workbook(archivoHistorialGraduados)
hojaHistorialGraduados = openfileHG.sheet_by_name("materias_graduados")

archivoEstudiantesGraduados = "..\\Excel_generados\\graduados.xlsx"
openfileEG = xlrd.open_workbook(archivoEstudiantesGraduados)
hojaEstudiantesGraduados = openfileEG.sheet_by_name("estudiantes_graduados")

MallaSM = "..\\Excel_generados\\materias_malla.xlsx"
FileSM = xlrd.open_workbook(MallaSM)
hojaSM = FileSM.sheet_by_name("sistemas_multimedia")

MallaSI = "..\\Excel_generados\\materias_malla.xlsx"
FileSI = xlrd.open_workbook(MallaSI)
hojaSI = FileSI.sheet_by_name("sistemas_de_informacion")

MallaST = "..\\Excel_generados\\materias_malla.xlsx"
FileST = xlrd.open_workbook(MallaST)
hojaST = FileST.sheet_by_name("sistemas_tecnologicos")

MallaCom = "..\\Excel_generados\\materias_malla.xlsx"
FileCom = xlrd.open_workbook(MallaCom)
hojaCom = FileCom.sheet_by_name("computacion")

semestreAlumno = []
semestre_estudiante=1
creditos_total = 0

matricula_malla = {}
malla_SM = {}
malla_SI = {}
malla_ST = {}
malla_Com = {}

for i in range(hojaEstudiantesGraduados.nrows):
    if str(hojaEstudiantesGraduados.cell_value(i, 1)) != "matricula":
        matricula_malla[str(hojaEstudiantesGraduados.cell_value(i, 1))]= str(hojaEstudiantesGraduados.cell_value(i, 9))

malla_SM = mallasDiccionario(hojaSM)
malla_SI = mallasDiccionario(hojaSI)
malla_ST = mallasDiccionario(hojaST)
malla_Com = mallasDiccionario(hojaCom)

print(matricula_malla)
print(malla_SM)
print(malla_SI)
print(malla_ST)
print(malla_Com)


for i in range(hojaHistorialGraduados.nrows):
    if i>0:
        matricula = str((hojaHistorialGraduados.cell_value(i, 1)))
        matricula_anterior = str((hojaHistorialGraduados.cell_value(i-1, 1)))
        #print("Posicion",i," Matricula actual " + matricula + "-- Matricula anterior " + matricula_anterior)

        if matricula == matricula_anterior or matricula_anterior=="matricula":
            semestre = str((hojaHistorialGraduados.cell_value(i, 4)))+"-"+str((hojaHistorialGraduados.cell_value(i, 5)))
            semestre_anterior = str((hojaHistorialGraduados.cell_value(i-1, 4)))+"-"+str((hojaHistorialGraduados.cell_value(i-1, 5)))
            creditos = str((hojaHistorialGraduados.cell_value(i, 8)))
            creditos = creditos.split(".")
            creditos = creditos[0]
            #print("Posicion", i, " Semestre actual " + semestre + "-- Semestre anterior " + semestre_anterior)

            if semestre == semestre_anterior or semestre_anterior=="año-termino" or str(hojaHistorialGraduados.cell_value(i, 5)) == "3S":
                info_semestre = str(semestre_estudiante)+" semestre"
                semestreAlumno.append(info_semestre)
                creditos_total += int(creditos)
                print("Posicion",i," Matricula actual " + matricula + "-- Semestre actual " + semestre+" Se encuentra en su ",semestre_estudiante," Semestre")
            else:
                semestre_estudiante = semestre_estudiante+1
                info_semestre = str(semestre_estudiante)+" semestre"
                semestreAlumno.append(info_semestre)
                print("En este semestre tomó en total ", creditos_total, " créditos\n")
                creditos_total=int(creditos)
                print("Posicion", i," Matricula actual " + matricula + "-- Semestre actual " + semestre + " Se encuentra en su ",semestre_estudiante, " Semestre")
        else:
            semestre_estudiante = 1
            info_semestre = str(semestre_estudiante) + " semestre"
            semestreAlumno.append(info_semestre)
            print("En este semestre tomó en total ", creditos_total, " créditos\n")
            print("Posicion", i," Matricula actual " + matricula + "-- Semestre actual " + semestre + " Se encuentra en su ",semestre_estudiante, " Semestre")

