#Librerias
library(dplyr)
library(xlsx)
library(readxl)
#Ruta de trabajo
setwd("C:/Users/Charles/Desktop/ESPOL/Semestre 7/Metodología de la Investigacion/Proyecto/Git/MetodologiaProyecto/Excel_generados")
datos <- read_excel("graduados.xlsx",sheet = 3)
View(datos)


#Materias problemas en conjunto
materias_problema <- datos[(datos$`codigo materia` == 'ICM00604') |(datos$`codigo materia` == 'FIEC01735')| (datos$`codigo materia` == 'ICF01131')|(datos$`codigo materia` == 'ICM01941')|(datos$`codigo materia` == 'ICM01958')|(datos$`codigo materia` == 'ICM01966'),]
#materias_problema <-materia_integradora[materia_integradora$estado_materia=='AP',]
materias_problema<-select(materias_problema,c(2,3,4,5,6,7,8))
View(materias_problema)


#Materias problemas individuales
#Calculos
materia_problema_calculo_diferencial <- datos[(datos$`codigo materia` == 'ICM01941'),]
materia_problema_calculo_integral <- datos[(datos$`codigo materia` == 'ICM01958'),]
materia_problema_calculo_varias<- datos[(datos$`codigo materia` == 'ICM01966'),]
#Algebra
materia_problema_algebra <- datos[(datos$`codigo materia` == 'ICM00604'),]
#Relacionado A Fisica
materia_problema_FisicaC <- datos[(datos$`codigo materia` == 'ICF01131'),]
materia_problema_redes_electricas<- datos[(datos$`codigo materia` == 'FIEC01735'),]
#POO
materia_problema_poo<- datos[(datos$`codigo materia` == 'FIEC04622'),]




View(materia_integradora)
write.xlsx(materias_problema,"./graduados.xlsx","estudiantes_materiasProblema",append = TRUE)
