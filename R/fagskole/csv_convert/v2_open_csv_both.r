############### PAKKESJEKK OG INSTALLASJON ###############
if (!require("data.table")) install.packages("data.table")
if (!require("tools")) install.packages("tools")
if (!require("openxlsx")) install.packages("openxlsx")

library(data.table)
library(tools)
library(openxlsx)

############### KONVERTERING XLSX TIL CSV ###############
konverter_xlsx_til_csv <- function(mappe_sti) {
  xlsx_filer <- list.files(path = mappe_sti, pattern = "\\.xlsx$", full.names = TRUE)
  csv_mappe <- file.path(mappe_sti, "csv_temp")
  dir.create(csv_mappe, showWarnings = FALSE)
  
  for (xlsx_fil in xlsx_filer) {
    system(sprintf('soffice --headless --convert-to csv --outdir "%s" "%s"', 
                   csv_mappe, xlsx_fil))
    cat("Konverterte", basename(xlsx_fil), "til CSV\n")
  }
  
  return(csv_mappe)
}

############### HOVEDFUNKSJON ###############
behandle_regnskap_filer <- function(
  mappe_sti, 
  output_fil,
  resultat_rad = 6,
  balanse_rad = 6
) {
  # Konverter XLSX til CSV
  csv_mappe <- konverter_xlsx_til_csv(mappe_sti)
  
  # Les målfil med openxlsx
  wb <- loadWorkbook(output_fil)
  mal_data <- read.xlsx(wb, sheet = "Tab_resultat", colNames = FALSE)
  fagskole_navn <- toupper(trimws(mal_data[3:29, 1]))
  
  cat("Tilgjengelige fagskoler i målfilen:\n")
  print(fagskole_navn)
  
  # Finn alle CSV-filer
  csv_filer <- list.files(path = csv_mappe, pattern = "\\.csv$", full.names = TRUE)
  
  for (fil in csv_filer) {
    cat("\nBehandler fil:", basename(fil), "\n")
    
    # Les CSV-data
    resultat_data <- fread(fil)
    
    # Hent fagskolens navn fra første rad
    første_rad <- resultat_data[1, 1]
    if (!grepl("^Fagskolens navn:", første_rad)) {
      cat("ADVARSEL: Uventet format på første rad\n")
      next
    }
    
    fagskole_navn_fra_fil <- sub("^Fagskolens navn:\\s*", "", første_rad)
    fagskole_navn_fra_fil <- toupper(trimws(fagskole_navn_fra_fil))
    
    if (fagskole_navn_fra_fil == "") {
      cat("ADVARSEL: Fant tomt fagskolenavn\n")
      next
    }
    
    match_index <- which(fagskole_navn == fagskole_navn_fra_fil)
    if (length(match_index) == 0) {
      cat("ADVARSEL: Fant ikke match for", fagskole_navn_fra_fil, "\n")
      next
    }
    
    gjeldende_rad <- match_index + resultat_rad - 1
    
    # Hent ut verdier og beregn nøkkeltall for 2024
    verdier_2024 <- list(
      driftsinntekter = as.numeric(resultat_data[12, 3]),
      driftskostnader = as.numeric(resultat_data[20, 3]),
      driftsresultat = as.numeric(resultat_data[24, 3]),
      arsresultat = as.numeric(resultat_data[33, 3]),
      omlopsmidler = sum(as.numeric(resultat_data[c(35, 40, 46, 51), 3]), na.rm = TRUE),
      egenkapital = as.numeric(resultat_data[20, 3]),
      kortsiktig_gjeld = as.numeric(resultat_data[47, 3]),
      totalkapital = as.numeric(resultat_data[51, 3])
    )
    
    verdier_2024$driftsmargin <- (verdier_2024$driftsresultat / verdier_2024$driftsinntekter) * 100
    verdier_2024$egenkapitalgrad <- (verdier_2024$egenkapital / verdier_2024$totalkapital) * 100
    verdier_2024$finansieringsgrad2 <- (verdier_2024$omlopsmidler / verdier_2024$kortsiktig_gjeld) * 100
    
    # Hent ut verdier og beregn nøkkeltall for 2023
    verdier_2023 <- list(
      driftsinntekter = as.numeric(resultat_data[12, 4]),
      driftskostnader = as.numeric(resultat_data[20, 4]),
      driftsresultat = as.numeric(resultat_data[24, 4]),
      arsresultat = as.numeric(resultat_data[33, 4]),
      omlopsmidler = sum(as.numeric(resultat_data[c(35, 40, 46, 51), 4]), na.rm = TRUE),
      egenkapital = as.numeric(resultat_data[20, 4]),
      kortsiktig_gjeld = as.numeric(resultat_data[47, 4]),
      totalkapital = as.numeric(resultat_data[51, 4])
    )
    
    verdier_2023$driftsmargin <- (verdier_2023$driftsresultat / verdier_2023$driftsinntekter) * 100
    verdier_2023$egenkapitalgrad <- (verdier_2023$egenkapital / verdier_2023$totalkapital) * 100
    verdier_2023$finansieringsgrad2 <- (verdier_2023$omlopsmidler / verdier_2023$kortsiktig_gjeld) * 100
    
    # Skriv nøkkeltall til målfil
    writeData(wb, sheet = "Tab_resultat", x = round(verdier_2024$driftsmargin, 1), startCol = 5, startRow = gjeldende_rad)
    writeData(wb, sheet = "Tab_resultat", x = round(verdier_2024$egenkapitalgrad, 1), startCol = 6, startRow = gjeldende_rad)
    writeData(wb, sheet = "Tab_resultat", x = round(verdier_2024$finansieringsgrad2, 1), startCol = 7, startRow = gjeldende_rad)
    
    writeData(wb, sheet = "Tab_resultat", x = round(verdier_2023$driftsmargin, 1), startCol = 8, startRow = gjeldende_rad)
    writeData(wb, sheet = "Tab_resultat", x = round(verdier_2023$egenkapitalgrad, 1), startCol = 9, startRow = gjeldende_rad)
    writeData(wb, sheet = "Tab_resultat", x = round(verdier_2023$finansieringsgrad2, 1), startCol = 10, startRow = gjeldende_rad)
    
    cat("Behandlet", basename(fil), "for", fagskole_navn_fra_fil, "\n")
  }
  
  saveWorkbook(wb, output_fil, overwrite = TRUE)
  unlink(csv_mappe, recursive = TRUE)
}