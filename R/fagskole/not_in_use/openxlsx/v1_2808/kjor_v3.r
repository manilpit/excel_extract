behandle_regnskap_filer <- function(
  mappe_sti, 
  output_fil,
  resultat_rad = 6,
  balanse_rad = 6
) {
  suppressPackageStartupMessages(library(openxlsx))
  
  # Finn alle Excel-filer i mappen
  excel_filer <- list.files(path = mappe_sti, pattern = "\\.xlsx$", full.names = TRUE)
  
  # For første fil, vis tilgjengelige ark
  første_fil <- excel_filer[1]
  wb_test <- loadWorkbook(første_fil)
  cat("Tilgjengelige ark i første fil:", første_fil, "\n")
  print(names(wb_test))
  
  # Stopp her midlertidig for å se arknavnene
  return(names(wb_test))
}

source("v3.r")
arknavnene <- behandle_regnskap_filer(
  mappe_sti = "//wsl.localhost/Ubuntu-24.04/home/manilpit/github/manilpit_github/excel_extract/data/fagskole",
  output_fil = "//wsl.localhost/Ubuntu-24.04/home/manilpit/github/manilpit_github/excel_extract/R/fagskole/mal/Kontroll_private_fagskoler_2025.xlsx",
  resultat_rad = 6,
  balanse_rad = 6
)