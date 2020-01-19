import pandas as pd
from pandas import ExcelWriter
from openpyxl import load_workbook
import xlrd

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

archivoGraduados = "..\\Excel_generados\\graduados.xlsx"
openfileA = xlrd.open_workbook(archivoGraduados)
hojaGraduados = openfileA.sheet_by_name("materias_graduados")

MallaA = "..\\Excel_generados\\materias_malla.xlsx"
FileA = xlrd.open_workbook(MallaA)
hojaA = FileA.sheet_by_name("sistemas_multimedia")

MallaB = "..\\Excel_generados\\materias_malla.xlsx"
FileB = xlrd.open_workbook(MallaB)
hojaB = FileB.sheet_by_name("sistemas_de_informacion")

MallaC = "..\\Excel_generados\\materias_malla.xlsx"
FileC = xlrd.open_workbook(MallaC)
hojaC = FileC.sheet_by_name("sistemas_tecnologicos")

MallaCom = "..\\Excel_generados\\materias_malla.xlsx"
FileCom = xlrd.open_workbook(MallaCom)
hojaCom = FileCom.sheet_by_name("computacion")

MallaGen = "..\\Excel_generados\\materias_malla.xlsx"
FileGen = xlrd.open_workbook(MallaGen)
hojaGen = FileGen.sheet_by_name("malla_generica")

lista_SM = []
lista_SI = []
lista_ST = []
lista_COM = []
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

for i in range(hojaCom.nrows):
    if str(hojaCom.cell_value(i, 1)) not in lista_GEN:
        lista_COM.append(str(hojaCom.cell_value(i, 1)))


"""print(lista_SM)
print(lista_SI)
print(lista_ST)
print(lista_COM)
print(len(lista_COM))"""

contador = 0


conST = 0
conSM = 0
conSI = 0
conCOM = 0

lista_especializacion= []

for i in range(hojaGraduados.nrows):
    if i>0:
        matricula = str((hojaGraduados.cell_value(i, 1)))
        matricula_anterior = str((hojaGraduados.cell_value(i-1, 1)))
        #print("Posicion", i," Matricula actual " + matricula + "-- Semestre actual " + semestre)

        if matricula_anterior == matricula or matricula_anterior=="matricula":

            if str(hojaGraduados.cell_value(i, 6)) in lista_ST:
                #print("Materia de ST "+str(hojaGraduados.cell_value(i, 2)))
                conST +=1
            elif str(hojaGraduados.cell_value(i, 6)) in lista_SM:
                #print("Materia de SM "+str(hojaGraduados.cell_value(i, 2)))
                conSM += 1
            elif str(hojaGraduados.cell_value(i, 6)) in lista_SI:
                #print("Materia de SI "+str(hojaGraduados.cell_value(i, 2)))
                conSI += 1
            elif str(hojaGraduados.cell_value(i, 6)) in lista_COM:
                #print("Materia de Computación "+str(hojaGraduados.cell_value(i, 2)))
                conCOM += 1

        else:
            semestre = str((hojaGraduados.cell_value(i - 1, 4))) + "-" + str((hojaGraduados.cell_value(i - 1, 5)))
            contador += 1
            print("Estudiante ",contador," Materias ST:",conST," -- Materias SM:",conSM,"-- Materias SI:",conSI," -- Materias COM",conCOM," ---  matricula: "+matricula_anterior)
            if conST>=4 or (conST > conSM and conST > conSI and conST > conCOM):
                lista_especializacion.append("ST")
                print("Graduado como estudiante de Sistemas tecnológicos")
            elif conSM>=4 or (conSM > conSI and conSM > conCOM):
                lista_especializacion.append("SM")
                print("Graduado como estudiante en Sistemas multimedia")
            elif conSI>=4 or (conSI > conCOM):
                lista_especializacion.append("SI")
                print("Graduado como estudiante en Sistemas informáticos")
            elif conST<=3 and conSM<=3 and conSI<=3:
                lista_especializacion.append("C")
                print("Graduado como estudiante en Computación")
            else:
                print("No se puede definir\n")
            conST = conSM = conSI = conCOM = 0
#Ultimo estudiante
print("Estudiante ",contador," Materias ST:",conST," -- Materias SM:",conSM,"-- Materias SI:",conSI," -- Materias COM",conCOM," --- matricula: "+matricula)
if conST > conSM and conST > conSI and conST > conCOM:
    print("Graduado como estudiante de Sistemas tecnológicos")
elif conSM > conSI and conSM > conCOM:
    print("Graduado como estudiante en Sistemas multimedia")
elif conSI > conCOM:
    print("Graduado como estudiante en Sistemas informáticos")
else:
    print("Graduado como estudiante en Computación")

diccionario = {"malla":lista_especializacion}

agrega_Columna("..\\Excel_generados\\graduados.xlsx","estudiantes_graduados",diccionario,9)
