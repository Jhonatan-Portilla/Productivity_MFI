library("readstata13")
library("deaR")
library("writexl")
library("readxl")
library("zoo")
library("dplyr")
library("tidyr")
library("FEAR")
library("openxlsx")
library(foreign)

#Cargamos la base anual realizada en stata
base <- read.dta13("C:/Users/jport/Dropbox/JHONATAN PORTILLA/Stata/Productividad/base_anual_2003_2021.dta")

new_base <- subset(base, select = c(nmero,nombre,año,empleados,c_o_deflact,i_f_deflact,creditos_deflact,inverted_credprom,totaldeudores))

# Loop para crear dataframes para cada par de años
for (i in 2003:2020) {
  # Define los dos años
  años_a_incluir <- c(i, i + 1)
  
  # Filtra las empresas que tienen datos para ambos años en el par
  df_filtrado <- new_base %>%
    filter(año %in% años_a_incluir) %>%
    group_by(nmero) %>%
    filter(n_distinct(año) == 2) %>%
    ungroup()
  
  # Asigna el dataframe completo al entorno global
  assign(paste0("df_", i, "_", i + 1), df_filtrado)
  
  # Divide el dataframe por cada año y asigna
  for (j in 1:2) {
    df_year <- df_filtrado %>% filter(año == años_a_incluir[j])
    assign(paste0("df_", i, "_", i + 1, "_", j), df_year)
  }
}

# Define years and versions
years <- c("2003_2004", "2004_2005", "2005_2006", "2006_2007", "2007_2008",
           "2008_2009","2009_2010","2010_2011","2011_2012","2012_2013",
           "2013_2014","2014_2015","2015_2016","2016_2017","2017_2018",
           "2018_2019","2019_2020","2020_2021")

versions <- c("1", "2")

# list to save matrix
i_list <- list()
o_list <- list()

# Loop to create matrix
for (year in years) {
  for (version in versions) {
    # Build dataframe´s name
    df_name <- paste0("df_", year, "_", version)
    
    # Evalue expression to obtain the dataframe
    df <- get(df_name)
    
    # Create matrix i and o
    i_matrix <- t(as.matrix(subset(df, select = c(empleados, c_o_deflact))))
    o_matrix <- t(as.matrix(subset(df, select = c(i_f_deflact, creditos_deflact, inverted_credprom, totaldeudores))))
    
    # Save matrix in list
    i_list[[paste0("i_", year, "_", version)]] <- i_matrix
    o_list[[paste0("o_", year, "_", version)]] <- o_matrix
  }
}

# list to save results
malquimst_list <- list()
pvalues_list <- list()
gm_list <- list()
malquimst_ci_list <- list()


# Loop to calculate malquimst
for (year in years){
  malquimst_result <- malmquist(i_list[[paste0("i_", year, "_", "1")]],
                       o_list[[paste0("o_", year, "_", "1")]],
                       i_list[[paste0("i_", year, "_", "2")]],
                       o_list[[paste0("o_", year, "_", "2")]], 
                       ORIENTATION = 3, OUTPUT = 4, MVEC = NULL, NREP = 5000, 
                       NBCR =100, errchk = TRUE)
  
  malquimst_list[[paste0("r_",year)]] <- malquimst_result[["indices"]]
  gm_list[[paste0("r_",year)]] <- malquimst_result[["gm"]]
  pvalues_list[[paste0("r_",year)]] <- malquimst_result[["pval"]]
  malquimst_ci_list[[paste0("r_",year)]] <- malquimst_result[["ci"]]
  }

# Define years
years <- 2004:2021

# Inicializa una lista para guardar los dataframes resultantes
malquimst_dataframes <- list()

# Loop para combinar dataframes y matrices
for (year in years) {
  # Construir el nombre del dataframe y de la matriz en la lista
  df_name <- paste0("df_", year - 1, "_", year+0, "_2")
  matrix_name <- paste0("r_",  year - 1, "_", year+0)
  
  # Evaluar la expresión para obtener el dataframe y la matriz correspondiente
  df <- get(df_name)
  matrix <- malquimst_list[[matrix_name]]
  
  # Crear el dataframe combinado
  malquimst_combined <- cbind(subset(df, select = c(nmero, nombre, año)), matrix)
  
  #Assign new game
  new_names <- c("malm","e1","e2","t1","t2","s1","s2","s3")
  colnames(malquimst_combined)[4:ncol(malquimst_combined)] <- new_names
  
  # Guardar el dataframe en la lista
  malquimst_dataframes[[paste0("malquimst_", year)]] <- malquimst_combined
}

# Apppend dataframes
malquimst_2004_2021 <- do.call(rbind, malquimst_dataframes)

# Ordenar el dataframe en formato long
malquimst_2004_2021 <- malquimst_2004_2021 %>%
  arrange(nmero, year)



# Crear un nuevo workbook
wb <- createWorkbook()

# Añadir una hoja al workbook
addWorksheet(wb, "Datos")

# Escribir el dataframe en la hoja
writeData(wb, sheet = "Datos", malquimst_2004_2021)

# Guardar el workbook en un archivo Excel
saveWorkbook(wb, file = "malquimst_ordered.xlsx", overwrite = TRUE)




##########################Geometric means

# Define years
years <- 2004:2021

# Inicializa una lista para guardar los dataframes resultantes
malquimst_gm_df <- list()

# Loop para combinar dataframes y matrices
for (year in years) {
  matrix_name <- paste0("r_",  year-1,"_",year+0)
  
  matrix <- t(gm_list[[matrix_name]])
  
  # Crear el dataframe combinado
  malquimst_gm <- data.frame(matrix)
  
  #Assign new name
  new_names <- c("malm","e1","e2","t1","t2","s1","s2","s3")
  colnames(malquimst_gm) <- new_names
  
  #
  malquimst_gm <- cbind(data.frame(año = year), malquimst_gm)
  
  # Guardar el dataframe en la lista
  malquimst_gm_df[[paste0("malquimst_gm_", year)]] <- malquimst_gm
}

# Apppend dataframes
malquimst_gm_2004_2021 <- do.call(rbind, malquimst_gm_df)

# Crear un nuevo workbook
wb <- createWorkbook()

# Añadir una hoja al workbook
addWorksheet(wb, "Datos_gm")

# Escribir el dataframe en la hoja
writeData(wb, sheet = "Datos_gm", malquimst_gm_2004_2021)

# Guardar el workbook en un archivo Excel
saveWorkbook(wb, file = "malquimst_gm.xlsx", overwrite = TRUE)





######################PVALUES
# Define years
years <- 2004:2021

# Inicializa una lista para guardar los dataframes resultantes
malquimst_pval_df <- list()

# Loop para combinar dataframes y matrices
for (year in years) {
  matrix_name <- paste0("r_",  year-1,"_",year+0)
  
  matrix <- t(pvalues_list[[matrix_name]][1,])
  
  # Crear el dataframe combinado
  malquimst_pval <- data.frame(matrix)
  
  #Assign new name
  new_names <- c("malm","e1","e2","t1","t2","s1","s2","s3")
  colnames(malquimst_pval) <- new_names
  
  #
  malquimst_pval <- cbind(data.frame(año = year), malquimst_pval)
  
  # Guardar el dataframe en la lista
  malquimst_pval_df[[paste0("malquimst_pval_", year)]] <- malquimst_pval
}

# Apppend dataframes
malquimst_pval_2004_2021 <- do.call(rbind, malquimst_pval_df)

# Crear un nuevo workbook
wb <- createWorkbook()

# Añadir una hoja al workbook
addWorksheet(wb, "Datos_pval")

# Escribir el dataframe en la hoja
writeData(wb, sheet = "Datos_pval", malquimst_pval_2004_2021)

# Guardar el workbook en un archivo Excel
saveWorkbook(wb, file = "malquimst_pval.xlsx", overwrite = TRUE)



##### SAVE DTA STATA
# Suponiendo que tu dataframe se llama df
write.dta(malquimst_2004_2021, "malquimst_2004_2021.dta")
