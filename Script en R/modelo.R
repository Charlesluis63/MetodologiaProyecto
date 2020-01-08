library(xlsx)
library(readxl)
setwd("C:/Users/Charles/Desktop/ESPOL/Semestre 7/Metodología de la Investigacion/Proyecto/Datos")
datos <- read_excel("datos_fiec_v1.xlsx",sheet = 3)
as.character(datos$`fecha emision titulo`)
View(datos)
matricula <- datos$matricula
View(matricula)
NROW(matricula)
datos_importantes <- datos[c(1,2,3,5,7,8,10)]
View(datos_importantes)
estudiantes_graduados <- datos_importantes[datos_importantes$título!="NULL",]
View(estudiantes_graduados)
