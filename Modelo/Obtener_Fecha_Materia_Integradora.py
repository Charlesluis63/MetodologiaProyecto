from Creditos_Asignados_Por_Semestre import agrega_Columna,obtenerListas
from pandas import ExcelWriter
import xlrd
from openpyxl import load_workbook

fechas_de_egreso =[]
terminos_graduados = []
matriculas_graduados =[]


matriculas_integradoras = []
fechas_materia_integradora =[]
terminos_materia_integradora = []
estados_materia_integradora = []

ruta = "..\\Excel_generados\\graduados.xlsx"

openfile = xlrd.open_workbook(ruta)
hoja_estudiantes = openfile.sheet_by_name("estudiantes_graduados")
hoja_materias_integradoras = openfile.sheet_by_name("estudiantes_materiaIntegradora")

valores =[matriculas_graduados,fechas_de_egreso,terminos_graduados]
indices =[1,5,7]
tipos = [str,int,str]
obtenerListas(hoja_estudiantes,valores,indices,tipos)



valores =[matriculas_integradoras, fechas_materia_integradora,
          terminos_materia_integradora, estados_materia_integradora]
indices = [1,4,5,3]
tipos = [str,int,str,str]
obtenerListas(hoja_materias_integradoras,valores,indices,tipos)

def obtenerEstudiantesSinMateriaIntegradora(mat_grad,mat_integr):
    estudiante_sin_integ = []
    for i in mat_grad:
        if i.split(".")[0]not in mat_integr:
            print("El estudiante" + i + " no tiene integradora")
            estudiante_sin_integ.append(i)
    return estudiante_sin_integ
#estudiantes_sin =obtenerEstudiantesSinMateriaIntegradora(matriculas_graduados,matriculas_integradoras)

def obtenerFechasdeMaterias(mat_grad,fecha_egr,mat_integr,fecha_integr,termino_integr,estad):
    matriculas_a_cambiar =[]
    fecha = []
    terminos = []
    todo =[matriculas_a_cambiar,fecha,terminos]
    for i in range(len(mat_grad)):
        for j in range(len(mat_integr)):
                matg = mat_grad[i].split(".")[0]
                if(matg==mat_integr[j]):
                    if matg not in matriculas_a_cambiar:
                        matriculas_a_cambiar.append(matg)
                        fecha.append(fecha_integr[j])
                        terminos.append(termino_integr[j])

    return todo



todo =obtenerFechasdeMaterias(matriculas_graduados,fechas_de_egreso,matriculas_integradoras,fechas_materia_integradora,terminos_materia_integradora,estados_materia_integradora)


def ordenarFechasTerminos(mat_ord,mat_sinord,fec,term):
    fechaorde = []
    termorde = []
    for i in range(len(mat_ord)):
        for j in range(len(mat_sinord)):
            if mat_ord[i].split(".")[0] == mat_sinord[j]:
                fechaorde.append(fec[j])
                termorde.append(term[j])
    return fechaorde,termorde

fechas,terminos = ordenarFechasTerminos(matriculas_graduados,todo[0],todo[1],todo[2])


diccionario = {"fecha_materia_integradora":fechas,"termino_materia_integradora":terminos}
agrega_Columna(ruta,"estudiantes_graduados",diccionario,11)
