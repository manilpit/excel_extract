############### PAKKESJEKK OG INSTALLASJON ###############
if (!require("data.table")) {
  install.packages("data.table")
  library(data.table)
}

if (!require("tools")) {
  install.packages("tools")
  library(tools)
}

############### KONVERTERING XLSX TIL CSV ###############
konverter_xlsx_til_csv <- function(mappe_sti) {
  # Finn alle xlsx filer
  xlsx_filer <- list.files(path = mappe_sti, pattern = "\\.xlsx$", full.names = TRUE)
  
  # Opprett en temp-mappe for CSV-filer
  csv_mappe <- file.path(mappe_sti, "csv_temp")
  dir.create(csv_mappe, showWarnings = FALSE)
  
  for (xlsx_fil in xlsx_filer) {
    # Konverter med LibreOffice
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
  
  # Les målfil for fagskolenavn
  mal_data <- fread(output_fil, sheet = "Tab_resultat")
  fagskole_navn <- mal_data[3:29, 1]  # Bare faktiske fagskoler
  cat("Tilgjengelige fagskoler i målfilen:\n")
  print(fagskole_navn)
  
  # Finn alle CSV-filer
  csv_filer <- list.files(path = csv_mappe, pattern = "\\.csv$", full.names = TRUE)
  
  for (fil in csv_filer) {
    cat("\nBehandler fil:", basename(fil), "\n")
    
    # Les data fra CSV-filene
    resultat_data <- fread(fil, sheet = "Resultatregnskap")
    eiendeler_data <- fread(fil, sheet = "Balanse - eiendeler")
    gjeld_ek_data <- fread(fil, sheet = "Balanse - egenkapital og gjeld")
    
    # Finn fagskolenavn
    første_rad <- resultat_data[1]
    if (!grepl("^Fagskolens navn:", første_rad)) {
      cat("ADVARSEL: Uventet format på første rad\n")
      next
    }
    
    fagskole_navn_fra_fil <- sub("^Fagskolens navn:\\s*", "", første_rad)
    fagskole_navn_fra_fil <- trimws(toupper(fagskole_navn_fra_fil))
    
    if (fagskole_navn_fra_fil == "") {
      cat("ADVARSEL: Fant tomt fagskolenavn\n")
      next
    }
    
    # Match mot målfil
    match_index <- which(toupper(fagskole_navn) == fagskole_navn_fra_fil)
    if (length(match_index) == 0) {
      cat("ADVARSEL: Fant ikke match for", fagskole_navn_fra_fil, "\n")
      next
    }
    
    gjeldende_rad <- match_index + resultat_rad - 1
    
    # Hent ut verdier
    verdier_2024 <- list(
      resultat = list(
        driftsinntekter = as.numeric(resultat_data[12, 3]),
        driftskostnader = as.numeric(resultat_data[20, 3]),
        driftsresultat = as.numeric(resultat_data[24, 3]),
        arsresultat = as.numeric(resultat_data[33, 3])
      ),
      balanse = list(
        omlopsmidler = sum(
          as.numeric(eiendeler_data[35, 3]),
          as.numeric(eiendeler_data[40, 3]),
          as.numeric(eiendeler_data[46, 3]),
          as.numeric(eiendeler_data[51, 3]),
          na.rm = TRUE
        ),
        egenkapital = as.numeric(gjeld_ek_data[20, 3]),
        kortsiktig_gjeld = as.numeric(gjeld_ek_data[47, 3]),
        totalkapital = as.numeric(gjeld_ek_data[51, 3])
      )
    )
    
    # Beregn nøkkeltall 2024
    verdier_2024$resultat$driftsmargin <- 
      (verdier_2024$resultat$driftsresultat / verdier_2024$resultat$driftsinntekter) * 100
    verdier_2024$balanse$egenkapitalgrad <- 
      (verdier_2024$balanse$egenkapital / verdier_2024$balanse$totalkapital) * 100
    verdier_2024$balanse$finansieringsgrad2 <- 
      (verdier_2024$balanse$omlopsmidler / verdier_2024$balanse$kortsiktig_gjeld) * 100
    
    # Gjør det samme for 2023
    verdier_2023 <- list(
      resultat = list(
        driftsinntekter = as.numeric(resultat_data[12, 4]),
        driftskostnader = as.numeric(resultat_data[20, 4]),
        driftsresultat = as.numeric(resultat_data[24, 4]),
        arsresultat = as.numeric(resultat_data[33, 4])
      ),
      balanse = list(
        omlopsmidler = sum(
          as.numeric(eiendeler_data[35, 4]),
          as.numeric(eiendeler_data[40, 4]),
          as.numeric(eiendeler_data[46, 4]),
          as.numeric(eiendeler_data[51, 4]),
          na.rm = TRUE
        ),
        egenkapital = as.numeric(gjeld_ek_data[20, 4]),
        kortsiktig_gjeld = as.numeric(gjeld_ek_data[47, 4]),
        totalkapital = as.numeric(gjeld_ek_data[51, 4])
      )
    )
    
    # Beregn nøkkeltall 2023
    verdier_2023$resultat$driftsmargin <- 
      (verdier_2023$resultat$driftsresultat / verdier_2023$resultat$driftsinntekter) * 100
    verdier_2023$balanse$egenkapitalgrad <- 
      (verdier_2023$balanse$egenkapital / verdier_2023$balanse$totalkapital) * 100
    verdier_2023$balanse$finansieringsgrad2 <- 
      (verdier_2023$balanse$omlopsmidler / verdier_2023$balanse$kortsiktig_gjeld) * 100
    
    # Skriv til målfil
    # Her må vi implementere skriving til Excel-fil
    # Dette kan gjøres med openxlsx eller ved å lage en ny CSV som senere konverteres til Excel
    
    cat("Behandlet", basename(fil), "for", fagskole_navn_fra_fil, "\n")
  }
  
  # Rydd opp temp-mappe
  unlink(csv_mappe, recursive = TRUE)
}
