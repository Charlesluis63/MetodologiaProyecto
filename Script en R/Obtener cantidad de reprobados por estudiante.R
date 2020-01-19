#OBTENER MATERIAS Y CODIGOS
library(dplyr)
library(xlsx)
library(readxl)
setwd("C:/Users/Charles/Desktop/ESPOL/Semestre 7/Metodología de la Investigacion/Proyecto/Git/MetodologiaProyecto/Excel_generados")
datos <- read_excel("graduados.xlsx",sheet = 3)
View(datos)
datos <- datos[datos$estado_materia == 'RP',]
nrow(datos)
table(datos)
dataframe <- data.frame(matriculas = c(datos))
resultados <- dataframe %>% group_by(datos$matricula) %>% tally()
View(resultados)

write.xlsx(resultados,"./graduados.xlsx","veces_reprobados",append = TRUE)
