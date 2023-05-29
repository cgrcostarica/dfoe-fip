# Cargando las librerias --------------------------
library(readxl)
library(data.table)
library(writexl)

# Cargando el directorio de trabajo ------------------------
setwd("C:/Users/israel.ortiz/Documents/2023/Otros requerimientos/Seleccion_Para_Llamadas")

# Generar las muestras -------------------------------------------
muestras <- replicate(52, sample(1:39, 6, replace = FALSE))

# Crear el dataframe
df <- data.frame(Muestra = rep(1:52, each = 6), Funcionario = c(muestras))

# Imprimir el dataframe
print(df)

# Genera semanas del anio --------------

# Configurar la zona horaria
Sys.setenv(TZ = "UTC")


# Obtener la fecha actual
fecha_actual <- as.Date(Sys.Date())

# Obtener el número de la semana actual
numero_semana_actual <- as.numeric(format(fecha_actual, "%V"))

# Obtener el año actual
anio_actual <- as.numeric(format(fecha_actual, "%Y"))

# Generar las fechas de inicio y fin para las 52 semanas del año
fechas_inicio <- seq(as.Date(paste(anio_actual, "-01-02", sep = "")), as.Date(paste(anio_actual, "-12-31", sep = "")), by = "week")
fechas_fin <- fechas_inicio + 6

# Ajustar las fechas si la primera semana del año pertenece al año anterior
if (as.numeric(format(fechas_inicio[1], "%Y")) < anio_actual) {
  fechas_inicio <- fechas_inicio[-1]
  fechas_fin <- fechas_fin[-1]
}

# Obtener solo las primeras 52 fechas
fechas_inicio <- fechas_inicio[1:52]
fechas_fin <- fechas_fin[1:52]

# Crear el dataframe de las fechas
df_fechas <- data.frame(Semana = 1:52, Inicio = format(fechas_inicio, "%d-%m-%Y"), Fin = format(fechas_fin, "%d-%m-%Y"))

# Imprimir el dataframe
print(df_fechas)

# Une muestras con fechas ----------------------------

df_unido <- merge(df,df_fechas,by.x = "Muestra",by.y = "Semana",all.x = TRUE)

# Cargando los codigos de funcionarios y extensiones ---------------------

codigos <- data.table(read_excel("Desvío de llamadas_ Pruebas aleatorias.xlsx", 
                                 sheet = "Extensiones compañeros", skip = 1))
View(codigos)

# Uniendo a los datos de muestras con fechas los codigos y extensiones de funcionarios ------------------
head(df_unido)
head(codigos)
Muestra_func <- merge(df_unido,codigos,by.x = "Funcionario",by.y = "N",all.x = TRUE)
Muestra_funcionarios <- Muestra_func[order(Muestra_func$Muestra),][,c(2,1,3:6)][]
Muestra_funcionarios[1:18,][]

# Frecuencias de funcionarios --------------------

setDT(Muestra_funcionarios)[,.N, keyby = .(Nombre)][]

# Exportando a Excel ------------------------

write_xlsx(Muestra_funcionarios,"Muestra Funcionarios.xlsx")
