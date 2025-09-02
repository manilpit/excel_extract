behandle_regnskap_filer <- function(
  mappe_sti, 
  output_fil,
  resultat_rad = 12,  # Standard startrad for resultatverdier
  balanse_rad = 12,   # Standard startrad for balanseverdier
  resultat_målark = "Tab resultat",
  balanse_målark = "balanse"
) {
  # Last nødvendige pakker
  suppressPackageStartupMessages(library(openxlsx))
  
  # Finn alle Excel-filer i mappen
  excel_filer <- list.files(path = mappe_sti, pattern = "\\.xlsx$", full.names = TRUE)
  
  # Opprett eller last inn målfil
  if (file.exists(output_fil)) {
    wb <- loadWorkbook(output_fil)
  } else {
    wb <- createWorkbook()
    addWorksheet(wb, resultat_målark)
    addWorksheet(wb, balanse_målark)
  }
  
  # Kolonnedefinisjon for målarket
  resultat_kolonner <- list(
    driftsinntekter = 4,    # Kolonne D
    driftskostnader = 8,    # Kolonne H
    arsresultat = 16        # Kolonne P
  )
  
  balanse_kolonner <- list(
    sum_eiendeler = 4,      # Kolonne D
    egenkapital = 16,       # Kolonne P
    kortsiktig_gjeld = 28,  # Kolonne AB
    sum_ek_gjeld = 36       # Kolonne AJ
  )
  
  # For hver fil i mappen
  for (fil_nr in seq_along(excel_filer)) {
    fil <- excel_filer[fil_nr]
    cat(paste("\nBehandler fil:", basename(fil), "\n"))
    
    # Last inn arbeidsboken
    kilde_wb <- loadWorkbook(fil)
    
    # Definer arknavnene
    resultat_ark <- "Resultatregnskap"
    eiendeler_ark <- "Balanse - eiendeler"
    gjeld_ek_ark <- "Balanse - gjeld og egenkapital"
    
    # Funksjon for å finne verdi basert på kode
    finn_verdi_med_kode <- function(data, kode) {
      for(rad in 1:nrow(data)) {
        if(!is.na(data[rad, 5]) && data[rad, 5] == kode) {
          return(list(
            verdi_2024 = as.numeric(gsub("[^0-9.-]", "", as.character(data[rad, 3]))),
            verdi_2023 = as.numeric(gsub("[^0-9.-]", "", as.character(data[rad, 4])))
          ))
        }
      }
      return(list(verdi_2024 = NA, verdi_2023 = NA))
    }
    
    # Les data fra arkene
    resultat_data <- read.xlsx(fil, sheet = resultat_ark)
    eiendeler_data <- read.xlsx(fil, sheet = eiendeler_ark)
    gjeld_ek_data <- read.xlsx(fil, sheet = gjeld_ek_ark)
    
    # Hent verdier fra resultatregnskapet
    driftsinntekter <- finn_verdi_med_kode(resultat_data, "RE.6")
    driftskostnader <- finn_verdi_med_kode(resultat_data, "RE.12")
    arsresultat <- finn_verdi_med_kode(resultat_data, "RE.19")
    
    # Hent verdier fra balanse - eiendeler
    varer <- finn_verdi_med_kode(eiendeler_data, "EI.21")
    fordringer <- finn_verdi_med_kode(eiendeler_data, "EI.24")
    investeringer <- finn_verdi_med_kode(eiendeler_data, "EI.28")
    bank <- finn_verdi_med_kode(eiendeler_data, "EI.31")
    
    # Hent verdier fra balanse - gjeld og egenkapital
    egenkapital <- finn_verdi_med_kode(gjeld_ek_data, "GK.7")
    kortsiktig_gjeld <- finn_verdi_med_kode(gjeld_ek_data, "GK.26")
    sum_ek_gjeld <- finn_verdi_med_kode(gjeld_ek_data, "GK.28")
    
    # Beregn sum for balanse - eiendeler
    sum_eiendeler_2024 <- sum(
      varer$verdi_2024, 
      fordringer$verdi_2024, 
      investeringer$verdi_2024, 
      bank$verdi_2024, 
      na.rm = TRUE
    )
    
    sum_eiendeler_2023 <- sum(
      varer$verdi_2023, 
      fordringer$verdi_2023, 
      investeringer$verdi_2023, 
      bank$verdi_2023, 
      na.rm = TRUE
    )
    
    # Skriv verdier til resultatarket
    gjeldende_rad <- resultat_rad + fil_nr - 1
    
    # Skriv resultatverdier
    writeData(wb, resultat_målark, driftsinntekter$verdi_2024, 
             startRow = gjeldende_rad, startCol = resultat_kolonner$driftsinntekter)
    writeData(wb, resultat_målark, driftskostnader$verdi_2024, 
             startRow = gjeldende_rad, startCol = resultat_kolonner$driftskostnader)
    writeData(wb, resultat_målark, arsresultat$verdi_2024, 
             startRow = gjeldende_rad, startCol = resultat_kolonner$arsresultat)
    
    # Skriv balanseverdier
    writeData(wb, balanse_målark, sum_eiendeler_2024, 
             startRow = gjeldende_rad, startCol = balanse_kolonner$sum_eiendeler)
    writeData(wb, balanse_målark, egenkapital$verdi_2024, 
             startRow = gjeldende_rad, startCol = balanse_kolonner$egenkapital)
    writeData(wb, balanse_målark, kortsiktig_gjeld$verdi_2024, 
             startRow = gjeldende_rad, startCol = balanse_kolonner$kortsiktig_gjeld)
    writeData(wb, balanse_målark, sum_ek_gjeld$verdi_2024, 
             startRow = gjeldende_rad, startCol = balanse_kolonner$sum_ek_gjeld)
    
    # Skriv ut resultatene for verifisering
    cat("\nSkrevet følgende verdier for", basename(fil), "til rad", gjeldende_rad, ":\n")
    cat("Driftsinntekter 2024:", driftsinntekter$verdi_2024, "\n")
    cat("Driftskostnader 2024:", driftskostnader$verdi_2024, "\n")
    cat("Årsresultat 2024:", arsresultat$verdi_2024, "\n")
    cat("Sum eiendeler 2024:", sum_eiendeler_2024, "\n")
    cat("Sum EK og gjeld 2024:", sum_ek_gjeld$verdi_2024, "\n")
  }
  
  # Lagre målfilen
  saveWorkbook(wb, output_fil, overwrite = TRUE)
  cat(paste("\nLagret resultater til:", output_fil, "\n"))
}

# Eksempel på bruk:
# behandle_regnskap_filer(
#   mappe_sti = "sti/til/mappe/med/excel/filer",
#   output_fil = "output.xlsx",
#   resultat_rad = 12,
#   balanse_rad = 12,
#   resultat_målark = "Tab resultat",
#   balanse_målark = "balanse"
# )