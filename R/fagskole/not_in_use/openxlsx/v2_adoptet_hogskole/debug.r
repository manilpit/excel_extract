# debug_navn.R - Script for å analysere hvordan skolenavnene ser ut i kildefilene

# Hjelpefunksjoner 
normalize <- function(x) {
  x <- toupper(iconv(as.character(x), from = "", to = "UTF-8"))
  x <- gsub("\\s+", " ", x)
  trimws(x)
}

# Funksjon for å sjekke navn i en enkelt fil
sjekk_navn_i_fil <- function(fil_sti) {
  fil_navn <- basename(fil_sti)
  cat("\n==== Fil:", fil_navn, "====\n")
  
  # Ekstraher ID fra filnavnet
  id <- sub("^(\\d+)_.*", "\\1", fil_navn)
  if (id != fil_navn) {
    cat("ID fra filnavn:", id, "\n")
  }
  
  # Les inn Excel-filen
  tryCatch({
    excel_data <- read.xlsx(fil_sti, sheet = 1, colNames = FALSE)
    
    # Vis de første radene for å se hva som finnes der
    cat("\nFørste 10 rader i første kolonne:\n")
    for (i in 1:min(10, nrow(excel_data))) {
      val <- as.character(excel_data[i, 1])
      cat("Rad", i, ":", val, "\n")
    }
    
    # Søk etter "Fagskolens navn:" i hele dokumentet (første 10 rader, alle kolonner)
    cat("\nSøker etter 'Fagskolens navn' eller lignende:\n")
    found_name <- FALSE
    for (r in 1:min(10, nrow(excel_data))) {
      for (c in 1:min(5, ncol(excel_data))) {
        val <- as.character(excel_data[r, c])
        if (is.na(val)) next
        
        if (grepl("Fagskolens\\s+navn|Navn\\s+på\\s+fagskole", val, ignore.case = TRUE)) {
          cat("Funnet i celle [", r, ",", c, "]:", val, "\n")
          # Forsøk å ekstrahere selve navnet
          name_part <- sub(".*[Nn]avn\\s*:?\\s*", "", val)
          name_part <- trimws(name_part)
          if (nchar(name_part) > 0 && name_part != val) {
            cat("  Ekstrahert navn:", name_part, "\n")
            found_name <- TRUE
          }
        }
      }
    }
    
    if (!found_name) {
      cat("Fant ikke noe tydelig skolenavn-mønster\n")
    }
    
    # Prøv å finne skolenavn basert på mønstergjenkjenning av typiske navn
    cat("\nSøker etter typiske skolenavn-mønstre (f.eks. 'Fagskole', 'Høyskole', 'Akademi'):\n")
    name_patterns <- c("Fagskole", "Høyskole", "Akademi", "Skole", "Institutt", "Stiftelse")
    for (r in 1:min(15, nrow(excel_data))) {
      for (c in 1:min(5, ncol(excel_data))) {
        val <- as.character(excel_data[r, c])
        if (is.na(val)) next
        
        for (pattern in name_patterns) {
          if (grepl(pattern, val, ignore.case = TRUE)) {
            cat("Mulig navn i celle [", r, ",", c, "]:", val, "\n")
            break
          }
        }
      }
    }
    
  }, error = function(e) {
    cat("FEIL ved lesing av fil:", conditionMessage(e), "\n")
  })
}

# Hovedfunksjon
finn_navn_i_alle_filer <- function(kilde_mappe, max_filer = 5) {
  # Hent alle Excel-filer i mappen
  filer <- list.files(kilde_mappe, pattern = "\\.xlsx$|\\.xls$", full.names = TRUE)
  
  if (length(filer) == 0) {
    stop("Ingen Excel-filer funnet i mappen:", kilde_mappe)
  }
  
  cat("Fant", length(filer), "Excel-filer i mappen\n")
  cat("Vil analysere de første", min(max_filer, length(filer)), "filene\n")
  
  # Analyser filene
  for (i in 1:min(max_filer, length(filer))) {
    sjekk_navn_i_fil(filer[i])
  }
  
  cat("\nVil du se flere filer? (Skriv antall eller 'n' for nei): ")
  answer <- readline()
  
  if (tolower(answer) != "n" && !is.na(as.numeric(answer))) {
    more_files <- as.numeric(answer)
    if (more_files > 0 && (max_filer + more_files) <= length(filer)) {
      for (i in (max_filer+1):min(max_filer+more_files, length(filer))) {
        sjekk_navn_i_fil(filer[i])
      }
    }
  }
  
  cat("\nFerdig med å analysere filer.\n")
}

# Kjør funksjonen med din mappe
kilde_mappe <- "//wsl.localhost/Ubuntu-24.04/home/manilpit/github/manilpit_github/excel_extract/data/fagskole"
finn_navn_i_alle_filer(kilde_mappe)