behandle_regnskap_filer <- function(
  mappe_sti, 
  output_fil,
  resultat_rad = 12,
  balanse_rad = 12
) {
  suppressPackageStartupMessages(library(openxlsx))
  
  # Finn alle Excel-filer i mappen
  excel_filer <- list.files(path = mappe_sti, pattern = "\\.xlsx$", full.names = TRUE)
  
  # Sjekk arknavnene i første fil
  første_fil <- excel_filer[1]
  wb_test <- loadWorkbook(første_fil)
  cat("Tilgjengelige ark i første fil:", basename(første_fil), "\n")
  print(names(wb_test))
  
  # Spør om du vil fortsette
  cat("\nVil du fortsette med behandlingen av filene? (y/n): ")
  svar <- readline()
  if (tolower(svar) != "y") {
    return(invisible(NULL))
  }
  
  # Definer kolonneplasseringer for 2024 og 2023
  kolonner_2024 <- list(
    resultat = list(
      driftsinntekter = which(LETTERS == "D"),    # D
      driftskostnader = which(LETTERS == "H"),    # H
      arsresultat = which(LETTERS == "P")         # P
    ),
    balanse = list(
      omlopsmidler = which(LETTERS == "D"),       # D
      egenkapital = which(LETTERS == "H"),        # H
      kortsiktig_gjeld = which(LETTERS == "L"),   # L
      totalkapital = which(LETTERS == "P")        # P
    )
  )
  
  kolonner_2023 <- list(
    resultat = list(
      driftsinntekter = which(LETTERS == "C"),    # C
      driftskostnader = which(LETTERS == "G"),    # G
      arsresultat = which(LETTERS == "O")         # O
    ),
    balanse = list(
      omlopsmidler = which(LETTERS == "C"),       # C
      egenkapital = which(LETTERS == "G"),        # G
      kortsiktig_gjeld = which(LETTERS == "K"),   # K
      totalkapital = which(LETTERS == "O")        # O
    )
  )
  
  for (fil_nr in seq_along(excel_filer)) {
    fil <- excel_filer[fil_nr]
    cat(paste("\nBehandler fil:", basename(fil), "\n"))
    
    # Last inn kildefil
    kilde_wb <- loadWorkbook(fil)
    # les data fra arkene
    resultat_data <- read.xlsx(fil, sheet = "Resultatregnskap")
    eiendeler_data <- read.xlsx(fil, sheet = "Balanse - eiendeler")
    gjeld_ek_data <- read.xlsx(fil, sheet = "Balanse - egenkapital og gjeld")  # Endret denne linjen
    
    # Hent 2024 verdier (kolonne C i kildefilen)
    verdier_2024 <- list(
      resultat = list(
        driftsinntekter = as.numeric(resultat_data[12, 3]),    # C12
        driftskostnader = as.numeric(resultat_data[20, 3]),    # C20
        arsresultat = as.numeric(resultat_data[33, 3])         # C33
      ),
      balanse = list(
        omlopsmidler = sum(
          as.numeric(eiendeler_data[35, 3]),    # C35
          as.numeric(eiendeler_data[40, 3]),    # C40
          as.numeric(eiendeler_data[46, 3]),    # C46
          as.numeric(eiendeler_data[51, 3]),    # C51
          na.rm = TRUE
        ),
        egenkapital = as.numeric(gjeld_ek_data[20, 3]),        # C20
        kortsiktig_gjeld = as.numeric(gjeld_ek_data[47, 3]),   # C47
        totalkapital = as.numeric(gjeld_ek_data[51, 3])        # C51
      )
    )
    
    # Hent 2023 verdier (kolonne D i kildefilen)
    verdier_2023 <- list(
      resultat = list(
        driftsinntekter = as.numeric(resultat_data[12, 4]),    # D12
        driftskostnader = as.numeric(resultat_data[20, 4]),    # D20
        arsresultat = as.numeric(resultat_data[33, 4])         # D33
      ),
      balanse = list(
        omlopsmidler = sum(
          as.numeric(eiendeler_data[35, 4]),    # D35
          as.numeric(eiendeler_data[40, 4]),    # D40
          as.numeric(eiendeler_data[46, 4]),    # D46
          as.numeric(eiendeler_data[51, 4]),    # D51
          na.rm = TRUE
        ),
        egenkapital = as.numeric(gjeld_ek_data[20, 4]),        # D20
        kortsiktig_gjeld = as.numeric(gjeld_ek_data[47, 4]),   # D47
        totalkapital = as.numeric(gjeld_ek_data[51, 4])        # D51
      )
    )
    
    gjeldende_rad <- resultat_rad + fil_nr - 1
    
    # Skriv 2024 verdier
    # Resultat
    writeData(wb, "Tab_resultat", verdier_2024$resultat$driftsinntekter, 
             startRow = gjeldende_rad, startCol = kolonner_2024$resultat$driftsinntekter)
    writeData(wb, "Tab_resultat", verdier_2024$resultat$driftskostnader, 
             startRow = gjeldende_rad, startCol = kolonner_2024$resultat$driftskostnader)
    writeData(wb, "Tab_resultat", verdier_2024$resultat$arsresultat, 
             startRow = gjeldende_rad, startCol = kolonner_2024$resultat$arsresultat)
    
    # Balanse
    writeData(wb, "Tab_balanse", verdier_2024$balanse$omlopsmidler, 
             startRow = gjeldende_rad, startCol = kolonner_2024$balanse$omlopsmidler)
    writeData(wb, "Tab_balanse", verdier_2024$balanse$egenkapital, 
             startRow = gjeldende_rad, startCol = kolonner_2024$balanse$egenkapital)
    writeData(wb, "Tab_balanse", verdier_2024$balanse$kortsiktig_gjeld, 
             startRow = gjeldende_rad, startCol = kolonner_2024$balanse$kortsiktig_gjeld)
    writeData(wb, "Tab_balanse", verdier_2024$balanse$totalkapital, 
             startRow = gjeldende_rad, startCol = kolonner_2024$balanse$totalkapital)
    
    # Skriv 2023 verdier
    # Resultat
    writeData(wb, "Tab_resultat", verdier_2023$resultat$driftsinntekter, 
             startRow = gjeldende_rad, startCol = kolonner_2023$resultat$driftsinntekter)
    writeData(wb, "Tab_resultat", verdier_2023$resultat$driftskostnader, 
             startRow = gjeldende_rad, startCol = kolonner_2023$resultat$driftskostnader)
    writeData(wb, "Tab_resultat", verdier_2023$resultat$arsresultat, 
             startRow = gjeldende_rad, startCol = kolonner_2023$resultat$arsresultat)
    
    # Balanse
    writeData(wb, "Tab_balanse", verdier_2023$balanse$omlopsmidler, 
             startRow = gjeldende_rad, startCol = kolonner_2023$balanse$omlopsmidler)
    writeData(wb, "Tab_balanse", verdier_2023$balanse$egenkapital, 
             startRow = gjeldende_rad, startCol = kolonner_2023$balanse$egenkapital)
    writeData(wb, "Tab_balanse", verdier_2023$balanse$kortsiktig_gjeld, 
             startRow = gjeldende_rad, startCol = kolonner_2023$balanse$kortsiktig_gjeld)
    writeData(wb, "Tab_balanse", verdier_2023$balanse$totalkapital, 
             startRow = gjeldende_rad, startCol = kolonner_2023$balanse$totalkapital)
    
    # Skriv ut for verifisering
    cat(paste("\nSkrevet verdier for", basename(fil), "til rad", gjeldende_rad, "\n"))
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
#   balanse_rad = 12
# )