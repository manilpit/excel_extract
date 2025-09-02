# ------- Pakker -------
library(openxlsx)
library(data.table)

# ------- Hjelpere -------
normalize <- function(x) {
  x <- toupper(iconv(as.character(x), from = "", to = "UTF-8"))
  x <- gsub("\\s+", " ", x)
  trimws(x)
}

to_num <- function(x) {
  if (is.numeric(x)) return(x)
  x <- as.character(x)
  x <- gsub("\u00A0", "", x)     # non-breaking space
  x <- gsub("\\s", "", x)        # whitespace
  x <- gsub("\\.", "", x)        # tusenskilletegn '.'
  x <- sub(",", ".", x)          # desimal-komma -> punktum
  suppressWarnings(as.numeric(x))
}

# Ny funksjon for å finne riktig gjeld/egenkapital-ark
find_gjeld_sheet <- function(kilde_fil) {
  sheet_names <- getSheetNames(kilde_fil)
  
  # Prøv begge varianter
  gjeld_variants <- c(
    "Balanse -- gjeld og egenkapital",
    "Balanse - gjeld og egenkapital", 
    "Balanse -- egenkapital og gjeld",
    "Balanse - egenkapital og gjeld"
  )
  
  for (variant in gjeld_variants) {
    if (variant %in% sheet_names) {
      cat("Fant gjeld-ark:", variant, "\n")
      return(variant)
    }
  }
  
  # Hvis ingen eksakt match, prøv partial matching
  for (sheet in sheet_names) {
    if (grepl("gjeld.*egenkapital|egenkapital.*gjeld", sheet, ignore.case = TRUE)) {
      cat("Fant gjeld-ark med partial match:", sheet, "\n")
      return(sheet)
    }
  }
  
  stop("Kunne ikke finne gjeld/egenkapital-ark i filen. Tilgjengelige ark: ", paste(sheet_names, collapse = ", "))
}

# Ny funksjon for å finne riktig eiendeler-ark  
find_eiendeler_sheet <- function(kilde_fil) {
  sheet_names <- getSheetNames(kilde_fil)
  
  # Prøv begge varianter
  eiendeler_variants <- c(
    "Balanse -- eiendeler",
    "Balanse - eiendeler"
  )
  
  for (variant in eiendeler_variants) {
    if (variant %in% sheet_names) {
      cat("Fant eiendeler-ark:", variant, "\n")
      return(variant)
    }
  }
  
  # Hvis ingen eksakt match, prøv partial matching
  for (sheet in sheet_names) {
    if (grepl("eiendeler", sheet, ignore.case = TRUE)) {
      cat("Fant eiendeler-ark med partial match:", sheet, "\n")
      return(sheet)
    }
  }
  
  stop("Kunne ikke finne eiendeler-ark i filen. Tilgjengelige ark: ", paste(sheet_names, collapse = ", "))
}

match_institution_row <- function(wb_path, sheet, inst_name, start_row = 6, name_col = 1, id_col = 2) {
  tab <- read.xlsx(wb_path, sheet = sheet, colNames = FALSE)
  
  if (nrow(tab) < start_row) {
    stop("Målarket '", sheet, "' har for få rader (", nrow(tab), "), forventer minst ", start_row)
  }
  
  target <- normalize(inst_name)
  
  id_match <- regexpr("Fagskole\\s+(\\d+)", target)
  id_from_name <- NULL
  if (id_match > 0) {
    id_match_text <- regmatches(target, id_match)
    id_from_name <- sub("Fagskole\\s+(\\d+).*", "\\1", id_match_text)
  }
  
  names_vec <- normalize(tab[start_row:nrow(tab), name_col])
  
  cat("Leter etter:", target, "\n")
  
  idx <- which(names_vec == target)
  if (length(idx) >= 1) return(start_row + idx[1] - 1)
  
  if (nchar(target) > 10) {
    short_target <- substr(target, 1, 15)
    for (i in seq_along(names_vec)) {
      if (grepl(short_target, names_vec[i], fixed = TRUE)) {
        cat("Matchet på delstreng:", short_target, "->", tab[start_row + i - 1, name_col], "\n")
        return(start_row + i - 1)
      }
    }
  }
  
  if (!is.null(id_from_name) && ncol(tab) >= id_col) {
    id_vec <- as.character(tab[start_row:nrow(tab), id_col])
    for (i in seq_along(id_vec)) {
      if (!is.na(id_vec[i]) && grepl(paste0("\\b", id_from_name, "\\b"), id_vec[i])) {
        cat("Matchet på ID:", id_from_name, "->", tab[start_row + i - 1, name_col], "\n")
        return(start_row + i - 1)
      }
    }
  }
  
  dist <- adist(target, names_vec)
  best <- which.min(dist)
  if (length(best) == 1 && is.finite(dist[best]) && dist[best] <= 5) {
    message(sprintf("MERK: '%s' ikke funnet eksakt i '%s'; brukte nærmeste '%s' (avstand %d).",
                    inst_name, sheet, tab[start_row + best - 1, name_col], dist[best]))
    return(start_row + best - 1)
  }
  
  if (length(names_vec) > 0) {
    warning("Kunne ikke finne match for '", inst_name, "' i ", sheet, ". Bruker første rad (", tab[start_row, name_col], ") som fallback.")
    return(start_row)
  }
  
  NA_integer_
}

get_school_name_from_source <- function(src_file, sheet_name = "Resultatregnskap") {
  sheet_names <- getSheetNames(src_file)
  if (!(sheet_name %in% sheet_names)) {
    warning("Arket '", sheet_name, "' finnes ikke i filen. Bruker første ark.")
    sheet_name <- sheet_names[1]
  }
  
  src_df <- read.xlsx(src_file, sheet = sheet_name, colNames = FALSE)
  first_cell <- as.character(src_df[1, 1])
  file_name <- basename(src_file)
  id <- sub("^(\\d+)_.*", "\\1", file_name)
  
  if (grepl("Fagskolens navn:", first_cell, ignore.case = TRUE)) {
    school_name <- sub(".*Fagskolens navn:\\s*", "", first_cell, ignore.case = TRUE)
    school_name <- trimws(school_name)
    
    if (nchar(school_name) == 0) {
      if (nrow(src_df) >= 2) {
        org_nr_cell <- as.character(src_df[2, 1])
        if (grepl("Org.nr:", org_nr_cell, ignore.case = TRUE)) {
          org_nr <- sub(".*Org.nr:\\s*", "", org_nr_cell)
          org_nr <- trimws(org_nr)
          if (nchar(org_nr) > 0) {
            school_name <- paste("Fagskole", id, "(Org.nr:", org_nr, ")")
            return(school_name)
          }
        }
      }
      school_name <- paste("Fagskole", id)
    }
    return(school_name)
  } 
  else if (nchar(trimws(first_cell)) > 0 && !grepl("^Note|^Prinsipp", first_cell, ignore.case = TRUE)) {
    return(first_cell)
  }
  return(paste("Fagskole", id))
}

# Debugging funksjon
debug_sheet_content <- function(data, sheet_name, target_rows = c(12, 20, 33), target_cols = c(3, 4)) {
  cat("\n=== DEBUGGING", sheet_name, "===\n")
  cat("Ark dimensjoner:", nrow(data), "x", ncol(data), "\n")
  
  # Vis innholdet i målcellene
  for (row in target_rows) {
    for (col in target_cols) {
      if (row <= nrow(data) && col <= ncol(data)) {
        cell_val <- data[row, col]
        cat("Rad", row, "Kolonne", col, ":", class(cell_val), "=", cell_val, "\n")
      } else {
        cat("Rad", row, "Kolonne", col, ": UTENFOR GRENSER\n")
      }
    }
  }
  
  # Vis noen rader rundt målradene
  for (row in target_rows) {
    if (row <= nrow(data)) {
      cat("\nRAD", row, "KONTEKST:\n")
      start_row <- max(1, row - 2)
      end_row <- min(nrow(data), row + 2)
      for (r in start_row:end_row) {
        if (ncol(data) >= 5) {
          cat("  Rad", r, ":", paste(data[r, 1:5], collapse = " | "), "\n")
        }
      }
    }
  }
}

# ------- Hovedfunksjon - MED DEBUGGING -------
excel_overforing_fagskole <- function(
  kilde_fil,
  output_fil,
  resultat_ark = "Tab_resultat",
  balanse_ark = "Tab_balanse",
  resultat_rad = NULL,
  balanse_rad = NULL
) {
  
  # Få skolenavnet fra Resultatregnskap-arket
  inst_name <- get_school_name_from_source(kilde_fil, sheet_name = "Resultatregnskap")
  cat("Fant skolenavn:", inst_name, "\n")
  
  # Finn raden i målarket
  if (is.null(resultat_rad)) {
    resultat_rad <- match_institution_row(output_fil, resultat_ark, inst_name)
    if (is.na(resultat_rad)) {
      stop("Kunne ikke finne raden for '", inst_name, "' i ", resultat_ark)
    }
  }
  
  if (is.null(balanse_rad)) {
    balanse_rad <- match_institution_row(output_fil, balanse_ark, inst_name)
    if (is.na(balanse_rad)) {
      stop("Kunne ikke finne raden for '", inst_name, "' i ", balanse_ark)
    }
  }
  
  cat("Skriver til rad:", resultat_rad, "(", resultat_ark, ") og rad", balanse_rad, "(", balanse_ark, ")\n")

  # Les inn kildedata med FASTE POSISJONER og dynamisk ark-søk
  cat("Leser data fra Resultatregnskap...\n")
  resultat_data <- read.xlsx(kilde_fil, sheet = "Resultatregnskap", colNames = FALSE)
  
  # Finn riktige ark-navn
  eiendeler_sheet <- find_eiendeler_sheet(kilde_fil)
  gjeld_sheet <- find_gjeld_sheet(kilde_fil)
  
  cat("Leser data fra", eiendeler_sheet, "...\n")
  eiendeler_data <- read.xlsx(kilde_fil, sheet = eiendeler_sheet, colNames = FALSE)
  
  cat("Leser data fra", gjeld_sheet, "...\n")
  gjeld_data <- read.xlsx(kilde_fil, sheet = gjeld_sheet, colNames = FALSE)

  # DEBUG: Vis innholdet i arkene
  debug_sheet_content(resultat_data, "Resultatregnskap", c(12, 20, 33), c(3, 4))
  debug_sheet_content(eiendeler_data, eiendeler_sheet, c(35, 40, 46, 51), c(3, 4))
  debug_sheet_content(gjeld_data, gjeld_sheet, c(20, 47, 51), c(3, 4))

  # Hent verdier med FASTE POSISJONER
  cat("\n=== HENTER VERDIER FRA FASTE POSISJONER ===\n")
  
  # 2024 data
  driftsinntekter_2024 <- to_num(resultat_data[12, 3])  # C12
  driftskostnader_2024 <- to_num(resultat_data[20, 3])  # C20
  arsresultat_2024 <- to_num(resultat_data[33, 3])      # C33
  
  cat("Fra resultat - 2024: Driftsinntekter:", driftsinntekter_2024, 
      "Driftskostnader:", driftskostnader_2024, "Årsresultat:", arsresultat_2024, "\n")
  
  # Omløpsmidler 2024 - sum av flere celler
  om1_2024 <- to_num(eiendeler_data[35, 3])  # C35
  om2_2024 <- to_num(eiendeler_data[40, 3])  # C40
  om3_2024 <- to_num(eiendeler_data[46, 3])  # C46
  om4_2024 <- to_num(eiendeler_data[51, 3])  # C51
  cat("Fra eiendeler - 2024: C35:", om1_2024, "C40:", om2_2024, "C46:", om3_2024, "C51:", om4_2024, "\n")
  
  omlopsmidler_2024 <- sum(c(om1_2024, om2_2024, om3_2024, om4_2024), na.rm = TRUE)
  if (all(is.na(c(om1_2024, om2_2024, om3_2024, om4_2024)))) omlopsmidler_2024 <- NA
  
  egenkapital_2024 <- to_num(gjeld_data[20, 3])         # C20
  korts_gjeld_2024 <- to_num(gjeld_data[47, 3])         # C47
  totalkapital_2024 <- to_num(gjeld_data[51, 3])        # C51
  
  cat("Fra gjeld - 2024: Egenkapital:", egenkapital_2024, 
      "Kortsiktig gjeld:", korts_gjeld_2024, "Totalkapital:", totalkapital_2024, "\n")
  
  # 2023 data
  driftsinntekter_2023 <- to_num(resultat_data[12, 4])  # D12
  driftskostnader_2023 <- to_num(resultat_data[20, 4])  # D20
  arsresultat_2023 <- to_num(resultat_data[33, 4])      # D33
  
  cat("Fra resultat - 2023: Driftsinntekter:", driftsinntekter_2023, 
      "Driftskostnader:", driftskostnader_2023, "Årsresultat:", arsresultat_2023, "\n")
  
  # Omløpsmidler 2023 - sum av flere celler
  om1_2023 <- to_num(eiendeler_data[35, 4])  # D35
  om2_2023 <- to_num(eiendeler_data[40, 4])  # D40
  om3_2023 <- to_num(eiendeler_data[46, 4])  # D46
  om4_2023 <- to_num(eiendeler_data[51, 4])  # D51
  cat("Fra eiendeler - 2023: D35:", om1_2023, "D40:", om2_2023, "D46:", om3_2023, "D51:", om4_2023, "\n")
  
  omlopsmidler_2023 <- sum(c(om1_2023, om2_2023, om3_2023, om4_2023), na.rm = TRUE)
  if (all(is.na(c(om1_2023, om2_2023, om3_2023, om4_2023)))) omlopsmidler_2023 <- NA
  
  egenkapital_2023 <- to_num(gjeld_data[20, 4])         # D20
  korts_gjeld_2023 <- to_num(gjeld_data[47, 4])         # D47
  totalkapital_2023 <- to_num(gjeld_data[51, 4])        # D51
  
  cat("Fra gjeld - 2023: Egenkapital:", egenkapital_2023, 
      "Kortsiktig gjeld:", korts_gjeld_2023, "Totalkapital:", totalkapital_2023, "\n")

  # Beregn driftsresultat (siden det ikke er direkte i mappingen)
  driftsresultat_2024 <- ifelse(is.na(driftsinntekter_2024) | is.na(driftskostnader_2024), 
                                NA, driftsinntekter_2024 - driftskostnader_2024)
  driftsresultat_2023 <- ifelse(is.na(driftsinntekter_2023) | is.na(driftskostnader_2023), 
                                NA, driftsinntekter_2023 - driftskostnader_2023)

  # Beregn nøkkeltall
  safe_div <- function(num, den) {
    if (is.na(num) | is.na(den) | den == 0) return(NA_real_) else return(num / den)
  }
  
  driftsmargin_2024 <- safe_div(driftsresultat_2024, driftsinntekter_2024) * 100
  driftsmargin_2023 <- safe_div(driftsresultat_2023, driftsinntekter_2023) * 100
  egenkapitalgrad_2024 <- safe_div(egenkapital_2024, totalkapital_2024) * 100
  egenkapitalgrad_2023 <- safe_div(egenkapital_2023, totalkapital_2023) * 100
  finansieringsgrad2_2024 <- safe_div(omlopsmidler_2024, korts_gjeld_2024) * 100
  finansieringsgrad2_2023 <- safe_div(omlopsmidler_2023, korts_gjeld_2023) * 100

  cat("\n=== BEREGNEDE NØKKELTALL ===\n")
  cat("Driftsmargin 2024:", round(driftsmargin_2024, 1), "2023:", round(driftsmargin_2023, 1), "\n")
  cat("Egenkapitalgrad 2024:", round(egenkapitalgrad_2024, 1), "2023:", round(egenkapitalgrad_2023, 1), "\n")
  cat("Finansieringsgrad2 2024:", round(finansieringsgrad2_2024, 1), "2023:", round(finansieringsgrad2_2023, 1), "\n")

  # Åpne målfil og skriv med FASTE KOLONNER
  wb <- loadWorkbook(output_fil)
  
  # Funksjon for sikker skriving
  safe_write <- function(sheet, value, row, col, description) {
    if (!is.na(value)) {
      tryCatch({
        writeData(wb, sheet = sheet, x = as.vector(value), startRow = row, startCol = col)
        cat("*** SKREV", description, ":", value, "til rad", row, "kolonne", col, "***\n")
      }, error = function(e) {
        cat("FEIL ved skriving av", description, ":", conditionMessage(e), "\n")
      })
    } else {
      cat("Hopper over", description, "- verdi er NA\n")
    }
  }

  cat("\n=== SKRIVER TIL RESULTAT-ARK ===\n")
  # Resultat-ark - FASTE KOLONNER
  safe_write(resultat_ark, driftsinntekter_2024, resultat_rad, 4, "driftsinntekter 2024")  # D
  safe_write(resultat_ark, driftsinntekter_2023, resultat_rad, 3, "driftsinntekter 2023")  # C
  safe_write(resultat_ark, driftskostnader_2024, resultat_rad, 8, "driftskostnader 2024")  # H
  safe_write(resultat_ark, driftskostnader_2023, resultat_rad, 7, "driftskostnader 2023")  # G
  safe_write(resultat_ark, arsresultat_2024, resultat_rad, 16, "årsresultat 2024")         # P
  safe_write(resultat_ark, arsresultat_2023, resultat_rad, 15, "årsresultat 2023")         # O
  safe_write(resultat_ark, round(driftsmargin_2024, 1), resultat_rad, 20, "driftsmargin 2024")  # T (kolonne 20)
  safe_write(resultat_ark, round(driftsmargin_2023, 1), resultat_rad, 19, "driftsmargin 2023")  # S (kolonne 19)

  cat("\n=== SKRIVER TIL BALANSE-ARK ===\n")
  # Balanse-ark - FASTE KOLONNER
  safe_write(balanse_ark, omlopsmidler_2024, balanse_rad, 4, "omløpsmidler 2024")         # D
  safe_write(balanse_ark, omlopsmidler_2023, balanse_rad, 3, "omløpsmidler 2023")         # C
  safe_write(balanse_ark, egenkapital_2024, balanse_rad, 8, "egenkapital 2024")           # H
  safe_write(balanse_ark, egenkapital_2023, balanse_rad, 7, "egenkapital 2023")           # G
  safe_write(balanse_ark, korts_gjeld_2024, balanse_rad, 12, "kortsiktig gjeld 2024")     # L
  safe_write(balanse_ark, korts_gjeld_2023, balanse_rad, 11, "kortsiktig gjeld 2023")     # K
  safe_write(balanse_ark, totalkapital_2024, balanse_rad, 16, "totalkapital 2024")        # P
  safe_write(balanse_ark, totalkapital_2023, balanse_rad, 15, "totalkapital 2023")        # O
  safe_write(balanse_ark, round(egenkapitalgrad_2024, 1), balanse_rad, 20, "egenkapitalgrad 2024")      # T
  safe_write(balanse_ark, round(egenkapitalgrad_2023, 1), balanse_rad, 19, "egenkapitalgrad 2023")      # S
  safe_write(balanse_ark, round(finansieringsgrad2_2024, 1), balanse_rad, 24, "finansieringsgrad2 2024") # X
  safe_write(balanse_ark, round(finansieringsgrad2_2023, 1), balanse_rad, 23, "finansieringsgrad2 2023") # W

  # Lagre workbook
  tryCatch({
    saveWorkbook(wb, output_fil, overwrite = TRUE)
    cat("\n*** WORKBOOK LAGRET SUCCESSFULLY ***\n")
  }, error = function(e) {
    stop("Feil ved lagring av workbook: ", conditionMessage(e))
  })

  cat(sprintf("*** FERDIG: %s -> rad %d (%s) / rad %d (%s) ***\n",
              inst_name, resultat_rad, resultat_ark, balanse_rad, balanse_ark))
  
  invisible(list(
    name = inst_name,
    driftsinntekter_2024 = driftsinntekter_2024,
    driftsinntekter_2023 = driftsinntekter_2023,
    arsresultat_2024 = arsresultat_2024,
    arsresultat_2023 = arsresultat_2023,
    omlopsmidler_2024 = omlopsmidler_2024,
    omlopsmidler_2023 = omlopsmidler_2023,
    egenkapital_2024 = egenkapital_2024,
    egenkapital_2023 = egenkapital_2023
  ))
}