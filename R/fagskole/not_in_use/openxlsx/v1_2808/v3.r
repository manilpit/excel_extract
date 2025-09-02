behandle_regnskap_filer <- function(
  mappe_sti, 
  output_fil,
  resultat_rad = 12,
  balanse_rad = 12
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
  
  # Spør om fortsettelse
  cat("\nVil du fortsette med behandlingen av filene? (y/n): ")
  svar <- readline()
  if (tolower(svar) != "y") {
    return(invisible(NULL))
  }
  
  # Oppdaterte kolonneplasseringer
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

  # For hver fil, spør brukeren om hvilket fagskolenavn den tilhører
  for (fil_nr in seq_along(excel_filer)) {
    fil <- excel_filer[fil_nr]
    cat(paste("\nBehandler fil:", basename(fil), "\n"))
    
    # Vis tilgjengelige fagskolenavn og be om matching
    cat("\nTilgjengelige fagskolenavn:\n")
    for (i in seq_along(fagskole_navn[[1]])) {
      cat(i, ": ", fagskole_navn[[1]][i], "\n")
    }
    cat("\nVelg nummer for fagskolen denne filen tilhører: ")
    valgt_nr <- as.numeric(readline())
    
    if (is.na(valgt_nr) || valgt_nr < 1 || valgt_nr > length(fagskole_navn[[1]])) {
      cat("Ugyldig valg. Hopper over denne filen.\n")
      next
    }
    
    gjeldende_rad <- valgt_nr + resultat_rad - 1
    
    # Les data fra arkene
    resultat_data <- read.xlsx(fil, sheet = "Resultatregnskap")
    eiendeler_data <- read.xlsx(fil, sheet = "Balanse - eiendeler")
    gjeld_ek_data <- read.xlsx(fil, sheet = "Balanse - egenkapital og gjeld")
    
    # Hent 2024 verdier
    verdier_2024 <- list(
      resultat = list(
        driftsinntekter = as.numeric(resultat_data[12, 3]),    # C12
        driftskostnader = as.numeric(resultat_data[20, 3]),    # C20
        driftsresultat = as.numeric(resultat_data[24, 3]),     # C24
        arsresultat = as.numeric(resultat_data[33, 3]),        # C33
        driftsmargin = as.numeric(resultat_data[24, 3]) / as.numeric(resultat_data[12, 3]) * 100  # Beregnet
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
    
    # Beregn ytterligere nøkkeltall for 2024
    verdier_2024$balanse$egenkapitalgrad <- 
      (verdier_2024$balanse$egenkapital / verdier_2024$balanse$totalkapital) * 100
    verdier_2024$balanse$finansieringsgrad2 <- 
      (verdier_2024$balanse$omlopsmidler / verdier_2024$balanse$kortsiktig_gjeld) * 100
    
    # Hent 2023 verdier
    verdier_2023 <- list(
      resultat = list(
        driftsinntekter = as.numeric(resultat_data[12, 4]),    # D12
        driftskostnader = as.numeric(resultat_data[20, 4]),    # D20
        driftsresultat = as.numeric(resultat_data[24, 4]),     # D24
        arsresultat = as.numeric(resultat_data[33, 4]),        # D33
        driftsmargin = as.numeric(resultat_data[24, 4]) / as.numeric(resultat_data[12, 4]) * 100  # Beregnet
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
    
    # Beregn ytterligere nøkkeltall for 2023
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
              "(", fagskole_navn[[1]][valgt_nr], ")\n"))
  }
  
  saveWorkbook(mal_wb, output_fil, overwrite = TRUE)
  cat(paste("\nLagret resultater til:", output_fil, "\n"))
}