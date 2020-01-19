library(dplyr)
library(xlsx)
library(readxl)
setwd("C:/Users/Charles/Desktop/ESPOL/Semestre 7/Metodología de la Investigacion/Proyecto/Git/MetodologiaProyecto/Excel_generados")
datos <- read_excel("graduados.xlsx",sheet = 3)
View(datos)
materia_integradora <- datos[(datos$`codigo materia` == 'FIEC07120') | (datos$`codigo materia` == 'CCPG1026')|(datos$`codigo materia` == 'ESPOL00133'),]
materia_integradora <-materia_integradora[materia_integradora$estado_materia=='AP',]
materia_integradora <-select(materia_integradora,c(2,3,4,5,6,7,8,9))


materia_integradora_cienciascomputacionales <- datos[(datos$`codigo materia` == 'FIEC07120'),]
materia_integradora_computacion <- datos[(datos$`codigo materia` == 'CCPG1026'),]


View(materia_integradora)
write.xlsx(materia_integradora,"./graduados.xlsx","estudiantes_materiaIntegradora",append = TRUE)
