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

nombre = "..\\Datos\\datos_fiec_v1.xlsx"
openfile = xlrd.open_workbook(nombre)
hoja = openfile.sheet_by_name("datos estudiantes")
print("Hay "+ str(hoja.nrows) + " registros actualmente")
estudiantes_graduados = []
fechas_ingreso_graduados = []
fechas_egreso_graduados = []
termino_ingreso_graduado = []
termino_egreso_graduado = []
año_ingreso_graduado = []
titulos_graduados = []
for i in range(hoja.nrows):
        matricula = str((hoja.cell_value(i,0)))
        matricula = matricula.split(".")

        titulo = str((hoja.cell_value(i, 9)))

        fecha_ingreso_hora = str((hoja.cell_value(i, 4)))
        fecha_ingreso = fecha_ingreso_hora.split(" ")[0]


        fecha_egreso = str((hoja.cell_value(i,6)))
        fecha_egreso = fecha_egreso.split(".")

        termino_egreso = str((hoja.cell_value(i, 7)))

        if(titulo!= "NULL" and i!= 0):
            fecha_ingreso_completa = fecha_ingreso.split("-")
            mes_ingreso = fecha_ingreso_completa[1]
            if mes_ingreso == "04" or mes_ingreso == "05":
                termino_ingreso_graduado.append("1S")
            else:
                termino_ingreso_graduado.append("2S")


            estudiantes_graduados.append(matricula[0])
            titulos_graduados.append(titulo)
            fechas_ingreso_graduados.append(fecha_ingreso)
            año_ingreso_graduado.append(fecha_ingreso_completa[0])
            fechas_egreso_graduados.append(fecha_egreso[0])
            termino_egreso_graduado.append(termino_egreso)
print(len(estudiantes_graduados))
print (len(titulos_graduados))
print(len(fechas_ingreso_graduados))
dic = {"matricula":estudiantes_graduados,"titulo_graduado":titulos_graduados,"fecha_ingreso":fechas_ingreso_graduados,"año_ingreso":año_ingreso_graduado,"fecha_egreso":fechas_egreso_graduados,"termino_ingreso":termino_ingreso_graduado,"termino_egreso":termino_egreso_graduado}
print(dic)

exporta_Excel(dic,"graduados.xlsx","estudiantes_graduados")



