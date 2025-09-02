# Oppdatert debug-script som sjekker alle ark i hver Excel-fil

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
  
  # Les inn Excel-filen og sjekk arkene
  tryCatch({
    # Hent alle arknavn i Excel-filen
    ark_navn <- getSheetNames(fil_sti)
    cat("\nFant", length(ark_navn), "ark i filen:", paste(ark_navn, collapse=", "), "\n")
    
    # Gå gjennom hvert ark
    for (ark in ark_navn) {
      cat("\n--- Ark:", ark, "---\n")
      excel_data <- read.xlsx(fil_sti, sheet = ark, colNames = FALSE)
      
      # Vis de første radene for å se hva som finnes der
      cat("Første 5 rader i første kolonne:\n")
      for (i in 1:min(5, nrow(excel_data))) {
        val <- as.character(excel_data[i, 1])
        cat("Rad", i, ":", val, "\n")
      }
      
      # Søk etter "Fagskolens navn:" eller "Resultatregnskap" i hele dokumentet
      cat("\nSøker etter 'Fagskolens navn', 'Resultatregnskap' eller lignende:\n")
      found_name <- FALSE
      for (r in 1:min(10, nrow(excel_data))) {
        for (c in 1:min(5, ncol(excel_data))) {
          val <- as.character(excel_data[r, c])
          if (is.na(val)) next
          
          if (grepl("Fagskolens\\s+navn|Navn\\s+på\\s+fagskole|Resultatregnskap|Regnskap\\s+for", val, ignore.case = TRUE)) {
            cat("Funnet i celle [", r, ",", c, "]:", val, "\n")
            # Forsøk å ekstrahere selve navnet
            if (grepl("Navn", val, ignore.case = TRUE)) {
              name_part <- sub(".*[Nn]avn\\s*:?\\s*", "", val)
              name_part <- trimws(name_part)
              if (nchar(name_part) > 0 && name_part != val) {
                cat("  Ekstrahert navn:", name_part, "\n")
                found_name <- TRUE
              }
            }
            # Sjekk om det er en resultatregnskap-overskrift med skolenavn
            else if (grepl("Regnskap|Resultat", val, ignore.case = TRUE)) {
              # Sjekk om det er et skolenavn i samme rad, annen kolonne
              for (other_c in 1:min(5, ncol(excel_data))) {
                if (other_c != c) {
                  other_val <- as.character(excel_data[r, other_c])
                  if (!is.na(other_val) && nchar(trimws(other_val)) > 0) {
                    cat("  Mulig skolenavn i samme rad:", other_val, "\n")
                  }
                }
              }
              # Sjekk om det er et skolenavn i neste rad
              if (r < nrow(excel_data)) {
                next_row_val <- as.character(excel_data[r+1, c])
                if (!is.na(next_row_val) && nchar(trimws(next_row_val)) > 0) {
                  cat("  Mulig skolenavn i neste rad:", next_row_val, "\n")
                }
              }
            }
          }
        }
      }
      
      if (!found_name) {
        cat("Fant ikke noe tydelig skolenavn-mønster i dette arket\n")
      }
      
      # Søk etter data som kan være resultatregnskapet
      cat("\nSøker etter typiske resultatregnskap-kolonner:\n")
      found_result <- FALSE
      for (r in 1:min(20, nrow(excel_data))) {
        for (c in 1:min(10, ncol(excel_data))) {
          val <- as.character(excel_data[r, c])
          if (is.na(val)) next
          
          if (grepl("driftsinntekter|driftskostnader|driftsresultat", val, ignore.case = TRUE)) {
            cat("Fant mulig resultatpost i celle [", r, ",", c, "]:", val, "\n")
            found_result <- TRUE
            
            # Vis noen rader rundt denne posten for å se strukturen
            start_row <- max(1, r-2)
            end_row <- min(nrow(excel_data), r+2)
            cat("  Kontekst rundt posten:\n")
            for (context_r in start_row:end_row) {
              row_vals <- character(0)
              for (context_c in 1:min(5, ncol(excel_data))) {
                cell_val <- as.character(excel_data[context_r, context_c])
                if (!is.na(cell_val) && nchar(trimws(cell_val)) > 0) {
                  row_vals <- c(row_vals, cell_val)
                }
              }
              if (length(row_vals) > 0) {
                cat("    Rad", context_r, ":", paste(row_vals, collapse=" | "), "\n")
              }
            }
            
            # Vi hopper ut etter å ha funnet første treff
            break
          }
        }
        if (found_result) break
      }
      
      if (!found_result) {
        cat("Fant ikke noe som ligner på resultatregnskap i dette arket\n")
      }
    }
    
  }, error = function(e) {
    cat("FEIL ved lesing av fil:", conditionMessage(e), "\n")
  })
}

# Hovedfunksjon
finn_navn_i_alle_filer <- function(kilde_mappe, max_filer = 3) {
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