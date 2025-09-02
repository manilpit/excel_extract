behandle_regnskap_filer <- function(
  mappe_sti, 
  output_fil,
  resultat_rad = 6,
  balanse_rad = 6
) {
  suppressPackageStartupMessages(library(openxlsx))
  
  # Finn alle Excel-filer i mappen
  excel_filer <- list.files(path = mappe_sti, pattern = "\\.xlsx$", full.names = TRUE)
  
  # Last først inn output-filen for å lese fagskolenavnene
  if (file.exists(output_fil)) {
    mal_wb <- loadWorkbook(output_fil)
    fagskole_navn <- read.xlsx(output_fil, sheet = "Tab_resultat", cols = 1)
    print("Tilgjengelige fagskoler i målfilen:")
    print(fagskole_navn)
  } else {
    stop("Målfilen må eksistere med fagskolenavn i kolonne A")
  }
  
  # Vis første fils arknavn
  første_fil <- excel_filer[1]
  wb_test <- loadWorkbook(første_fil)
  cat("\nTilgjengelige ark i første fil:", basename(første_fil), "\n")
  print(names(wb_test))
  
  # Definer kolonneplasseringer i målfilen
  kolonner_2023 <- list(
    resultat = list(
      driftsinntekter = which(LETTERS == "C"),    # C
      driftskostnader = which(LETTERS == "G"),    # G
      driftsresultat = which(LETTERS == "K"),     # K
      arsresultat = which(LETTERS == "O"),        # O
      driftsmargin = which(LETTERS == "S")        # S
    ),
    balanse = list(
      omlopsmidler = which(LETTERS == "C"),       # C
      egenkapital = which(LETTERS == "G"),        # G
      kortsiktig_gjeld = which(LETTERS == "K"),   # K
      totalkapital = which(LETTERS == "O"),       # O
      egenkapitalgrad = which(LETTERS == "S"),    # S
      finansieringsgrad2 = which(LETTERS == "W")  # W
    )
  )
  
  kolonner_2024 <- list(
    resultat = list(
      driftsinntekter = which(LETTERS == "D"),    # D
      driftskostnader = which(LETTERS == "H"),    # H
      driftsresultat = which(LETTERS == "L"),     # L
      arsresultat = which(LETTERS == "P"),        # P
      driftsmargin = which(LETTERS == "T")        # T
    ),
    balanse = list(
      omlopsmidler = which(LETTERS == "D"),       # D
      egenkapital = which(LETTERS == "H"),        # H
      kortsiktig_gjeld = which(LETTERS == "L"),   # L
      totalkapital = which(LETTERS == "P"),       # P
      egenkapitalgrad = which(LETTERS == "T"),    # T
      finansieringsgrad2 = which(LETTERS == "X")  # X
    )
  )

  for (fil_nr in seq_along(excel_filer)) {
    fil <- excel_filer[fil_nr]
    cat(paste("\nBehandler fil:", basename(fil), "\n"))
    
    # Les data fra arkene
    resultat_data <- read.xlsx(fil, sheet = "Resultatregnskap")
    eiendeler_data <- read.xlsx(fil, sheet = "Balanse - eiendeler")
    gjeld_ek_data <- read.xlsx(fil, sheet = "Balanse - egenkapital og gjeld")
    
    # Hent fagskolens navn fra A1 og rengjør det
    fagskole_navn_fra_fil <- gsub("Fagskolens navn: ", "", resultat_data[1, 1])
    fagskole_navn_fra_fil <- trimws(toupper(fagskole_navn_fra_fil))
    
    # Finn matching i målfilen
    fagskole_navn_i_mal <- toupper(fagskole_navn[[1]])
    match_index <- which(fagskole_navn_i_mal == fagskole_navn_fra_fil)
    
    if (length(match_index) == 0) {
      cat("ADVARSEL: Fant ikke match for", fagskole_navn_fra_fil, "i målfilen. Hopper over denne filen.\n")
      next
    }
    
    gjeldende_rad <- match_index + resultat_rad - 1
    cat("Fant match:", fagskole_navn[[1]][match_index], "- skriver til rad", gjeldende_rad, "\n")
    
    # Hent 2024 verdier
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
    
    # Beregn nøkkeltall for 2024
    verdier_2024$resultat$driftsmargin <- 
      (verdier_2024$resultat$driftsresultat / verdier_2024$resultat$driftsinntekter) * 100
    verdier_2024$balanse$egenkapitalgrad <- 
      (verdier_2024$balanse$egenkapital / verdier_2024$balanse$totalkapital) * 100
    verdier_2024$balanse$finansieringsgrad2 <- 
      (verdier_2024$balanse$omlopsmidler / verdier_2024$balanse$kortsiktig_gjeld) * 100
    
    # Hent 2023 verdier
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
    
    # Beregn nøkkeltall for 2023
    verdier_2023$resultat$driftsmargin <- 
      (verdier_2023$resultat$driftsresultat / verdier_2023$resultat$driftsinntekter) * 100
    verdier_2023$balanse$egenkapitalgrad <- 
      (verdier_2023$balanse$egenkapital / verdier_2023$balanse$totalkapital) * 100
    verdier_2023$balanse$finansieringsgrad2 <- 
      (verdier_2023$balanse$omlopsmidler / verdier_2023$balanse$kortsiktig_gjeld) * 100
    
    # Skriv 2024 verdier til resultatarket
    writeData(mal_wb, "Tab_resultat", verdier_2024$resultat$driftsinntekter, 
             startRow = gjeldende_rad, startCol = kolonner_2024$resultat$driftsinntekter)
    writeData(mal_wb, "Tab_resultat", verdier_2024$resultat$driftskostnader, 
             startRow = gjeldende_rad, startCol = kolonner_2024$resultat$driftskostnader)
    writeData(mal_wb, "Tab_resultat", verdier_2024$resultat$driftsresultat, 
             startRow = gjeldende_rad, startCol = kolonner_2024$resultat$driftsresultat)
    writeData(mal_wb, "Tab_resultat", verdier_2024$resultat$arsresultat, 
             startRow = gjeldende_rad, startCol = kolonner_2024$resultat$arsresultat)
    writeData(mal_wb, "Tab_resultat", verdier_2024$resultat$driftsmargin, 
             startRow = gjeldende_rad, startCol = kolonner_2024$resultat$driftsmargin)
    
    # Skriv 2024 verdier til balansearket
    writeData(mal_wb, "Tab_balanse", verdier_2024$balanse$omlopsmidler, 
             startRow = gjeldende_rad, startCol = kolonner_2024$balanse$omlopsmidler)
    writeData(mal_wb, "Tab_balanse", verdier_2024$balanse$egenkapital, 
             startRow = gjeldende_rad, startCol = kolonner_2024$balanse$egenkapital)
    writeData(mal_wb, "Tab_balanse", verdier_2024$balanse$kortsiktig_gjeld, 
             startRow = gjeldende_rad, startCol = kolonner_2024$balanse$kortsiktig_gjeld)
    writeData(mal_wb, "Tab_balanse", verdier_2024$balanse$totalkapital, 
             startRow = gjeldende_rad, startCol = kolonner_2024$balanse$totalkapital)
    writeData(mal_wb, "Tab_balanse", verdier_2024$balanse$egenkapitalgrad, 
             startRow = gjeldende_rad, startCol = kolonner_2024$balanse$egenkapitalgrad)
    writeData(mal_wb, "Tab_balanse", verdier_2024$balanse$finansieringsgrad2, 
             startRow = gjeldende_rad, startCol = kolonner_2024$balanse$finansieringsgrad2)
    
    # Skriv 2023 verdier til resultatarket
    writeData(mal_wb, "Tab_resultat", verdier_2023$resultat$driftsinntekter, 
             startRow = gjeldende_rad, startCol = kolonner_2023$resultat$driftsinntekter)
    writeData(mal_wb, "Tab_resultat", verdier_2023$resultat$driftskostnader, 
             startRow = gjeldende_rad, startCol = kolonner_2023$resultat$driftskostnader)
    writeData(mal_wb, "Tab_resultat", verdier_2023$resultat$driftsresultat, 
             startRow = gjeldende_rad, startCol = kolonner_2023$resultat$driftsresultat)
    writeData(mal_wb, "Tab_resultat", verdier_2023$resultat$arsresultat, 
             startRow = gjeldende_rad, startCol = kolonner_2023$resultat$arsresultat)
    writeData(mal_wb, "Tab_resultat", verdier_2023$resultat$driftsmargin, 
             startRow = gjeldende_rad, startCol = kolonner_2023$resultat$driftsmargin)
    
    # Skriv 2023 verdier til balansearket
    writeData(mal_wb, "Tab_balanse", verdier_2023$balanse$omlopsmidler, 
             startRow = gjeldende_rad, startCol = kolonner_2023$balanse$omlopsmidler)
    writeData(mal_wb, "Tab_balanse", verdier_2023$balanse$egenkapital, 
             startRow = gjeldende_rad, startCol = kolonner_2023$balanse$egenkapital)
    writeData(mal_wb, "Tab_balanse", verdier_2023$balanse$kortsiktig_gjeld, 
             startRow = gjeldende_rad, startCol = kolonner_2023$balanse$kortsiktig_gjeld)
    writeData(mal_wb, "Tab_balanse", verdier_2023$balanse$totalkapital, 
             startRow = gjeldende_rad, startCol = kolonner_2023$balanse$totalkapital)
    writeData(mal_wb, "Tab_balanse", verdier_2023$balanse$egenkapitalgrad, 
             startRow = gjeldende_rad, startCol = kolonner_2023$balanse$egenkapitalgrad)
    writeData(mal_wb, "Tab_balanse", verdier_2023$balanse$finansieringsgrad2, 
             startRow = gjeldende_rad, startCol = kolonner_2023$balanse$finansieringsgrad2)
    
    cat(paste("\nSkrevet verdier for", basename(fil), 
              "til rad", gjeldende_rad, 
              "(", fagskole_navn[[1]][match_index], ")\n"))
  }
  
  saveWorkbook(mal_wb, output_fil, overwrite = TRUE)
  cat(paste("\nLagret resultater til:", output_fil, "\n"))
}