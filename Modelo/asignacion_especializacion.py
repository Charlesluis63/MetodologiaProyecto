import pandas as pd
from pandas import ExcelWriter
import xlrd

def exporta_Excel(diccionario,nombre_archivo,nombre_hoja):
    columnas = []
    for i in diccionario:
        columnas.append(i)
    print(columnas)
    df = pd.DataFrame(dic)
    df = df[columnas]
    writer = ExcelWriter(nombre_archivo)
    df.to_excel(writer, sheet_name=nombre_hoja)
    writer.save()

archivoGraduados = "..\\Excel_generados\\graduados.xlsx"
openfileA = xlrd.open_workbook(archivoGraduados)
hojaGraduados = openfileA.sheet_by_name("materias_graduados")
contador=0

MallaA = "..\\Excel_generados\\materias_malla.xlsx"
FileA = xlrd.open_workbook(MallaA)
hojaA = FileA.sheet_by_name("sistemas_multimedia")

MallaB = "..\\Excel_generados\\materias_malla.xlsx"
FileB = xlrd.open_workbook(MallaB)
hojaB = FileB.sheet_by_name("sistemas_de_informacion")

MallaC = "..\\Excel_generados\\materias_malla.xlsx"
FileC = xlrd.open_workbook(MallaC)
hojaC = FileC.sheet_by_name("sistemas_tecnologicos")

MallaGen = "..\\Excel_generados\\materias_malla.xlsx"
FileGen = xlrd.open_workbook(MallaGen)
hojaGen = FileGen.sheet_by_name("malla_generica")

lista_SM = []
lista_SI = []
lista_ST = []
lista_GEN = []

for i in range(hojaGen.nrows):
    lista_GEN.append(str(hojaGen.cell_value(i, 1)))

for i in range(hojaA.nrows):
    if str(hojaA.cell_value(i, 1)) not in lista_GEN:
        lista_SM.append(str(hojaA.cell_value(i, 1)))

for i in range(hojaB.nrows):
    if str(hojaB.cell_value(i, 1)) not in lista_GEN:
        lista_SI.append(str(hojaB.cell_value(i, 1)))

for i in range(hojaC.nrows):
    if str(hojaC.cell_value(i, 1)) not in lista_GEN:
        lista_ST.append(str(hojaC.cell_value(i, 1)))

print(lista_SM)
print(lista_SI)
print(lista_ST)
contador = 1
print("Estudiante ",contador)
conST = 0
conSM = 0
conSI = 0
conCOM = 0
for i in range(hojaGraduados.nrows):
    if i>0:
        matricula = str((hojaGraduados.cell_value(i, 1)))
        matricula_anterior = str((hojaGraduados.cell_value(i-1, 1)))
        semestre = str((hojaGraduados.cell_value(i, 4))) + "-" + str((hojaGraduados.cell_value(i, 5)))
        #print("Posicion", i," Matricula actual " + matricula + "-- Semestre actual " + semestre)
        if matricula == matricula_anterior or matricula_anterior=="matricula":
            if str(hojaGraduados.cell_value(i, 6)) in lista_ST: #or str(hojaGraduados.cell_value(i, 6))=="FIEC05561" or str(hojaGraduados.cell_value(i, 6))=="FIEC06429":
                #print("Materia de ST "+str(hojaGraduados.cell_value(i, 2)))
                conST +=1
            elif str(hojaGraduados.cell_value(i, 6)) in lista_SM: #or str(hojaGraduados.cell_value(i, 6))=="FIEC05439" or str(hojaGraduados.cell_value(i, 6))=="FIEC05462":
                #print("Materia de SM "+str(hojaGraduados.cell_value(i, 2)))
                conSM += 1
            elif str(hojaGraduados.cell_value(i, 6)) in lista_SI: #or str(hojaGraduados.cell_value(i, 6))=="FIEC05322" or str(hojaGraduados.cell_value(i, 6))=="FIEC06445":
                #print("Materia de SI "+str(hojaGraduados.cell_value(i, 2)))
                conSI += 1
            elif str(hojaGraduados.cell_value(i, 6))=="CCPG1003":
                #print("Materia de Computación "+str(hojaGraduados.cell_value(i, 2))+" "+semestre)
                conCOM += 1
        else:
            contador += 1
            print("Estudiante ",contador,"   Materias ST:",conST," -- Materias SM:",conSM,"-- Materias SI:",conSI," -- Materias COM",conCOM," --- Semestre inicial: "+semestre+" matricula"+matricula)
            conST = conSM = conSI = conCOM = 0