#OBTENER MATERIAS Y CODIGOS
library(dplyr)
library(xlsx)
library(readxl)
setwd("C:/Users/Charles/Desktop/ESPOL/Semestre 7/Metodología de la Investigacion/Proyecto/Git/MetodologiaProyecto/Excel_generados")
datos <- read_excel("graduados.xlsx",sheet = 3)
View(datos)

datos <- datos[datos$termino == '3S',]
datos <- select(datos,c(2,3,4,5,7,8))
View(datos)
nrow(datos)
table(datos)

write.xlsx(datos,"./graduados.xlsx","estudiantes_tercer_termino",append = TRUE)

dataframe <- data.frame(matricula = c(datos))
resultados <- dataframe %>% group_by(datos$matricula) %>% tally()
View(resultados)

write.xlsx(resultados,"./graduados.xlsx","veces_tercer_termino",append = TRUE)
