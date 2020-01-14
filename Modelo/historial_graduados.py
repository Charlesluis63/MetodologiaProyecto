import pandas as pd
from pandas import ExcelWriter
import xlrd

def agregar_Excel(diccionario,nombre_archivo,nombre_hoja):
    columnas = []
    for i in diccionario:
        columnas.append(i)
    print(columnas)
    df = pd.DataFrame(dic)
    df = df[columnas]
    writer = ExcelWriter(nombre_archivo, mode='a')
    df.to_excel(writer, sheet_name=nombre_hoja)
    writer.save()

#Leemos el archivo de los alumnos graduados
archivoGraduados = "..\\Excel_generados\\graduados.xlsx"
openfileA = xlrd.open_workbook(archivoGraduados)
hojaGraduados = openfileA.sheet_by_name("estudiantes_graduados")

#Leemos el dataset
archivoTodo = "..\\Datos\\datos_fiec_v1.xlsx"
openfileB = xlrd.open_workbook(archivoTodo)
hojaHistoria = openfileB.sheet_by_name("historia_academica") #obtenemos la hoja "historia_academica"

matriculas_graduados_verificar = []
matriculas_graduados = []
estado_materias_graduados = []
nombre_materias_graduados = []
codigo_materias_graduados = []
semestre_año = []
semestre_termino = []
semestre = []

#Sacamos las matriculas de los alumnos graduados
for i in range(hojaGraduados.nrows):
    matricula = str((hojaGraduados.cell_value(i, 1)))
    matricula = matricula.split(".")
    matriculas_graduados_verificar.append(matricula[0])

print(matriculas_graduados_verificar)
print(len(matriculas_graduados_verificar))



#Recorremos el dataset
for i in range(hojaHistoria.nrows):
    matriculasHistorico = str((hojaHistoria.cell_value(i, 0)))
    matriculasHistorico = matriculasHistorico.split(".")

    # Verificamos si una matricula del dataset se encuentra en el array de alumnos graduados
    if matriculasHistorico[0] in matriculas_graduados_verificar:

        #Obtenemos los datos necesarios y los agregamos a los respectivos arrays
        estado = str((hojaHistoria.cell_value(i, 5)))

        nombre_materia = str((hojaHistoria.cell_value(i, 6)))

        año = str((hojaHistoria.cell_value(i, 1)))
        año = año.split(".")

        termino = str((hojaHistoria.cell_value(i, 2)))

        codigo = str((hojaHistoria.cell_value(i, 3)))

        matriculas_graduados.append(matriculasHistorico[0])
        nombre_materias_graduados.append(nombre_materia)
        estado_materias_graduados.append(estado)
        codigo_materias_graduados.append(codigo)
        semestre_año.append(año[0])
        semestre_termino.append(termino)

print(matriculas_graduados)
print(len(matriculas_graduados))

print(nombre_materias_graduados)
print(estado_materias_graduados)
print(semestre_año)
print(semestre_termino)

#Agregamos a un diccionario
dic = {"matricula":matriculas_graduados,"nombre_materia":nombre_materias_graduados,"estado_materia":estado_materias_graduados,"año":semestre_año,"termino":semestre_termino,"codigo materia":codigo_materias_graduados}
print(dic)

#Agregamos al excel de graduados.xlsx
agregar_Excel(dic,"..\\Excel_generados\\graduados.xlsx","materias_graduados")


