library("readstata13")
library("deaR")
library("writexl")
library("readxl")
library("zoo")
library("dplyr")
library("tidyr")
library("FEAR")
library("openxlsx")
library("foreign")

library("here")
library("readstata13")

# Define database path
path_data <- here("GitHub/Productivity_MFI/data", "base_2003_2021.dta")

#Load annual microfinance database 2003-2020
base <- read.dta13(path_data)

#Select variables set
new_base <- subset(base, select = c(nmero,nombre,año,empleados,c_o_deflact,i_f_deflact,creditos_deflact,inverted_credprom,totaldeudores))


########################### Data preparation

### Loop to create dataframes for every pair of years (2003-2004, 2004-2005, ... ,2020-2021)
for (i in 2003:2020) {
  # Define every pair 
  años_a_incluir <- c(i, i + 1)
  
  # Filter MFIs that have data for every couple of years
  df_filtrado <- new_base %>%
    filter(año %in% años_a_incluir) %>%
    group_by(nmero) %>%
    filter(n_distinct(año) == 2) %>%
    ungroup()
  
  # Assigns the entire dataframe to the global environment
  assign(paste0("df_", i, "_", i + 1), df_filtrado)
  
  # Split the dataframe by each pair of year and assign a name
  for (j in 1:2) {
    df_year <- df_filtrado %>% filter(año == años_a_incluir[j])
    assign(paste0("df_", i, "_", i + 1, "_", j), df_year)
  }
}

### Create input anda ouput matrix for each pair of year

# Define years and versions
years <- c("2003_2004", "2004_2005", "2005_2006", "2006_2007", "2007_2008",
           "2008_2009","2009_2010","2010_2011","2011_2012","2012_2013",
           "2013_2014","2014_2015","2015_2016","2016_2017","2017_2018",
           "2018_2019","2019_2020","2020_2021")

versions <- c("1", "2")

# Create list to save matrix
i_list <- list() #input list
o_list <- list() #output list

# Loop to create input and output matrix
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

########################### Estimate Malquimst Index

# List to save results
malquimst_list <- list() #Malquimst indices list
pvalues_list <- list() #Estimated  p-values for tests of significant difference from 1 for the geometric mean Malmquist index
gm_list <- list() #Geometric means of the indices returned in indices
malquimst_ci_list <- list() #Estimated confidence intervals corresponding to the geometric means


# Loop to calculate malquimst index for each pair of year
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

########################### Export MFI malquimst indices dataframes

# Create a list to save the MFI malquimst indices dataframes
malquimst_dataframes <- list()

# Loop to append each MFI malquimst indices dataframes
years <- 2004:2021
for (year in years) {
  # Create names of dataframes and matrices in the list
  df_name <- paste0("df_", year - 1, "_", year+0, "_2")
  matrix_name <- paste0("r_",  year - 1, "_", year+0)
  
  # Evaluate the expression to obtain the dataframe and the corresponding matrix
  df <- get(df_name)
  matrix <- malquimst_list[[matrix_name]]
  
  # Combine dataframe and matrix
  malquimst_combined <- cbind(subset(df, select = c(nmero, nombre, año)), matrix)
  
  #Assign column names
  new_names <- c("malm","e1","e2","t1","t2","s1","s2","s3")
  colnames(malquimst_combined)[4:ncol(malquimst_combined)] <- new_names
  
  # Save each MFI malquimst indices dataframes
  malquimst_dataframes[[paste0("malquimst_", year)]] <- malquimst_combined
}

# Apppend dataframes
malquimst_2004_2021 <- do.call(rbind, malquimst_dataframes)

# Save the dataframe in long format
malquimst_2004_2021 <- malquimst_2004_2021 %>%
  arrange(nmero, year)


# Export MFI malquimst indices in Excel
wb <- createWorkbook() #Create workbook
addWorksheet(wb, "Data") #Create a sheet
writeData(wb, sheet = "Data", malquimst_2004_2021) #Write data in a sheet   
saveWorkbook(wb, file = "malquimst_indeces.xlsx", overwrite = TRUE) #Save in a Excel

# Export MFI malquimst indices in dta file (Stata)
write.dta(malquimst_2004_2021, "malquimst_2004_2021.dta")



########################### Export geometric means of malquimst indices and decompositions

#Create a list to save annual geometric means of malquimst indices and decompositions
malquimst_gm_df <- list()

# Loop to append each annual geometric mean of malquimst indices and decompositions
years <- 2004:2021
for (year in years) {
  matrix_name <- paste0("r_",  year-1,"_",year+0)
  
  matrix <- t(gm_list[[matrix_name]])
  
  # Create dataframe
  malquimst_gm <- data.frame(matrix)
  
  #Assign new name columns
  new_names <- c("malm","e1","e2","t1","t2","s1","s2","s3")
  colnames(malquimst_gm) <- new_names
  
  #Combine MFIs data and geometric mean of malquimst indices and decompositions
  malquimst_gm <- cbind(data.frame(año = year), malquimst_gm)
  
  # Save datafame in a list
  malquimst_gm_df[[paste0("malquimst_gm_", year)]] <- malquimst_gm
}

# Apppend dataframes
malquimst_gm_2004_2021 <- do.call(rbind, malquimst_gm_df)

# Export geometric means of malquimst indices and decompositions in Excel
wb <- createWorkbook() #Create workbook
addWorksheet(wb, "malquimst_index_gm") #Create a sheet
writeData(wb, sheet = "malquimst_index_gm", malquimst_gm_2004_2021) #Write data in a sheet
saveWorkbook(wb, file = "malquimst_gm.xlsx", overwrite = TRUE) #Save in a Excel


########################### Export pvalues of malquimst indices and decompositions


#Create a list to save pvalues of malquimst indices and decompositions
malquimst_pval_df <- list()

# Loop to append pvalues of malquimst indices and decompositionss
years <- 2004:2021
for (year in years) {
  matrix_name <- paste0("r_",  year-1,"_",year+0)
  
  matrix <- t(pvalues_list[[matrix_name]][1,])
  
  #Create dataframe
  malquimst_pval <- data.frame(matrix)
  
  #Assign new name columns
  new_names <- c("malm","e1","e2","t1","t2","s1","s2","s3")
  colnames(malquimst_pval) <- new_names
  
  #Combine MFIs data and pvalues of malquimst indices and decompositions
  malquimst_pval <- cbind(data.frame(año = year), malquimst_pval)
  
  #Save datafame in a list
  malquimst_pval_df[[paste0("malquimst_pval_", year)]] <- malquimst_pval
}

# Apppend dataframes
malquimst_pval_2004_2021 <- do.call(rbind, malquimst_pval_df)

# Export pvalues of malquimst indices and decompositions in Excel
wb <- createWorkbook() #Create workbook
addWorksheet(wb, "malquimst_index_pval") #Create a sheet
writeData(wb, sheet = "malquimst_index_pval", malquimst_pval_2004_2021)  #Write data in a sheet
saveWorkbook(wb, file = "malquimst_pval.xlsx", overwrite = TRUE) #Save in a Excel
