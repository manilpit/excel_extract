# ------- Pakker -------
library(openxlsx)
library(data.table)

# ------- Hjelpere -------
normalize <- function(x) {
  # Håndter NA-verdier
  if (is.na(x) || is.null(x)) return("")
  
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

# Funksjon for å finne rad basert på kode i kolonne E
find_row_by_code <- function(data, code) {
  if (ncol(data) < 5) {
    cat("ADVARSEL: Arket har mindre enn 5 kolonner, kan ikke søke i kolonne E\n")
    return(NA_integer_)
  }
  
  # Søk i kolonne E (kolonne 5)
  for (row in 1:nrow(data)) {
    cell_value <- data[row, 5]
    
    # Sikker håndtering av NA
    if (is.na(cell_value) || is.null(cell_value)) {
      next
    }
    
    cell_value_norm <- normalize(cell_value)
    
    if (cell_value_norm == code) {
      cat("Fant kode", code, "på rad", row, ":", cell_value, "\n")
      return(row)
    }
  }
  
  cat("Fant ikke kode:", code, "\n")
  return(NA_integer_)
}

# Funksjon for å finne riktig gjeld/egenkapital-ark
find_gjeld_sheet <- function(kilde_fil) {
  sheet_names <- getSheetNames(kilde_fil)
  
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
  
  for (sheet in sheet_names) {
    if (grepl("gjeld.*egenkapital|egenkapital.*gjeld", sheet, ignore.case = TRUE)) {
      cat("Fant gjeld-ark med partial match:", sheet, "\n")
      return(sheet)
    }
  }
  
  stop("Kunne ikke finne gjeld/egenkapital-ark i filen. Tilgjengelige ark: ", paste(sheet_names, collapse = ", "))
}

find_eiendeler_sheet <- function(kilde_fil) {
  sheet_names <- getSheetNames(kilde_fil)
  
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
  
  names_vec <- sapply(tab[start_row:nrow(tab), name_col], normalize)
  
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

# Debug funksjon for å vise kolonne E innhold
debug_column_e <- function(data, sheet_name) {
  cat("\n=== DEBUG KOLONNE E I", sheet_name, "===\n")
  if (ncol(data) < 5) {
    cat("Arket har kun", ncol(data), "kolonner - kan ikke sjekke kolonne E\n")
    return()
  }
  
  cat("Viser første 20 rader i kolonne E:\n")
  for (i in 1:min(20, nrow(data))) {
    cell_val <- data[i, 5]
    cat("Rad", i, ":", class(cell_val), "=", cell_val, "\n")
  }
}

# ------- Hovedfunksjon - MED BEDRE FEILHÅNDTERING -------
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

  # Les inn kildedata
  cat("Leser data fra Resultatregnskap...\n")
  resultat_data <- read.xlsx(kilde_fil, sheet = "Resultatregnskap", colNames = FALSE)
  
  eiendeler_sheet <- find_eiendeler_sheet(kilde_fil)
  gjeld_sheet <- find_gjeld_sheet(kilde_fil)
  
  cat("Leser data fra", eiendeler_sheet, "...\n")
  eiendeler_data <- read.xlsx(kilde_fil, sheet = eiendeler_sheet, colNames = FALSE)
  
  cat("Leser data fra", gjeld_sheet, "...\n")
  gjeld_data <- read.xlsx(kilde_fil, sheet = gjeld_sheet, colNames = FALSE)

  # Debug kolonne E innhold
  debug_column_e(resultat_data, "Resultatregnskap")
  debug_column_e(eiendeler_data, eiendeler_sheet)
  debug_column_e(gjeld_data, gjeld_sheet)

  # Søk etter kodene og hent verdier
  cat("\n=== SØKER ETTER KODER OG HENTER VERDIER ===\n")
  
  # Resultatregnskap
  re6_rad <- find_row_by_code(resultat_data, "RE.6")    # Driftsinntekter
  re12_rad <- find_row_by_code(resultat_data, "RE.12")  # Driftskostnader  
  re19_rad <- find_row_by_code(resultat_data, "RE.19")  # Årsresultat
  
  # Balanse - eiendeler (omløpsmidler)
  ei21_rad <- find_row_by_code(eiendeler_data, "EI.21")
  ei24_rad <- find_row_by_code(eiendeler_data, "EI.24")
  ei28_rad <- find_row_by_code(eiendeler_data, "EI.28")
  ei31_rad <- find_row_by_code(eiendeler_data, "EI.31")
  
  # Balanse - gjeld og egenkapital
  gk7_rad <- find_row_by_code(gjeld_data, "GK.7")     # Egenkapital
  gk26_rad <- find_row_by_code(gjeld_data, "GK.26")   # Kortsiktig gjeld
  gk28_rad <- find_row_by_code(gjeld_data, "GK.28")   # Totalkapital

  # Hent verdier for 2024 (kolonne C = kolonne 3)
  cat("\n=== HENTER 2024-VERDIER (KOLONNE C) ===\n")
  driftsinntekter_2024 <- if (!is.na(re6_rad)) to_num(resultat_data[re6_rad, 3]) else NA
  driftskostnader_2024 <- if (!is.na(re12_rad)) to_num(resultat_data[re12_rad, 3]) else NA
  arsresultat_2024 <- if (!is.na(re19_rad)) to_num(resultat_data[re19_rad, 3]) else NA
  
  # Omløpsmidler 2024 - sum av flere EI-koder
  om_values_2024 <- c()
  if (!is.na(ei21_rad)) om_values_2024 <- c(om_values_2024, to_num(eiendeler_data[ei21_rad, 3]))
  if (!is.na(ei24_rad)) om_values_2024 <- c(om_values_2024, to_num(eiendeler_data[ei24_rad, 3]))
  if (!is.na(ei28_rad)) om_values_2024 <- c(om_values_2024, to_num(eiendeler_data[ei28_rad, 3]))
  if (!is.na(ei31_rad)) om_values_2024 <- c(om_values_2024, to_num(eiendeler_data[ei31_rad, 3]))
  
  omlopsmidler_2024 <- if (length(om_values_2024) > 0) sum(om_values_2024, na.rm = TRUE) else NA
  if (all(is.na(om_values_2024))) omlopsmidler_2024 <- NA
  
  egenkapital_2024 <- if (!is.na(gk7_rad)) to_num(gjeld_data[gk7_rad, 3]) else NA
  korts_gjeld_2024 <- if (!is.na(gk26_rad)) to_num(gjeld_data[gk26_rad, 3]) else NA
  totalkapital_2024 <- if (!is.na(gk28_rad)) to_num(gjeld_data[gk28_rad, 3]) else NA

  # Hent verdier for 2023 (kolonne D = kolonne 4)
  cat("\n=== HENTER 2023-VERDIER (KOLONNE D) ===\n")
  driftsinntekter_2023 <- if (!is.na(re6_rad)) to_num(resultat_data[re6_rad, 4]) else NA
  driftskostnader_2023 <- if (!is.na(re12_rad)) to_num(resultat_data[re12_rad, 4]) else NA
  arsresultat_2023 <- if (!is.na(re19_rad)) to_num(resultat_data[re19_rad, 4]) else NA
  
  # Omløpsmidler 2023 - sum av flere EI-koder
  om_values_2023 <- c()
  if (!is.na(ei21_rad)) om_values_2023 <- c(om_values_2023, to_num(eiendeler_data[ei21_rad, 4]))
  if (!is.na(ei24_rad)) om_values_2023 <- c(om_values_2023, to_num(eiendeler_data[ei24_rad, 4]))
  if (!is.na(ei28_rad)) om_values_2023 <- c(om_values_2023, to_num(eiendeler_data[ei28_rad, 4]))
  if (!is.na(ei31_rad)) om_values_2023 <- c(om_values_2023, to_num(eiendeler_data[ei31_rad, 4]))
  
  omlopsmidler_2023 <- if (length(om_values_2023) > 0) sum(om_values_2023, na.rm = TRUE) else NA
  if (all(is.na(om_values_2023))) omlopsmidler_2023 <- NA
  
  egenkapital_2023 <- if (!is.na(gk7_rad)) to_num(gjeld_data[gk7_rad, 4]) else NA
  korts_gjeld_2023 <- if (!is.na(gk26_rad)) to_num(gjeld_data[gk26_rad, 4]) else NA
  totalkapital_2023 <- if (!is.na(gk28_rad)) to_num(gjeld_data[gk28_rad, 4]) else NA

  # Beregn driftsresultat
  driftsresultat_2024 <- if (!is.na(driftsinntekter_2024) && !is.na(driftskostnader_2024)) 
    driftsinntekter_2024 - driftskostnader_2024 else NA
  driftsresultat_2023 <- if (!is.na(driftsinntekter_2023) && !is.na(driftskostnader_2023)) 
    driftsinntekter_2023 - driftskostnader_2023 else NA

  # Debug output
  cat("\n=== HENTET VERDIER ===\n")
  cat("2024: Driftsinnt:", driftsinntekter_2024, "Driftskost:", driftskostnader_2024, "Driftsres:", driftsresultat_2024, "Årsres:", arsresultat_2024, "\n")
  cat("2024: Omløp:", omlopsmidler_2024, "Egenk:", egenkapital_2024, "Kortgjeld:", korts_gjeld_2024, "Total:", totalkapital_2024, "\n")
  cat("2023: Driftsinnt:", driftsinntekter_2023, "Driftskost:", driftskostnader_2023, "Driftsres:", driftsresultat_2023, "Årsres:", arsresultat_2023, "\n")
  cat("2023: Omløp:", omlopsmidler_2023, "Egenk:", egenkapital_2023, "Kortgjeld:", korts_gjeld_2023, "Total:", totalkapital_2023, "\n")

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

  # Åpne målfil og skriv
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
  # Følger mappingen: 2024 til kolonne D(4), 2023 til kolonne C(3)
  safe_write(resultat_ark, driftsinntekter_2024, resultat_rad, 4, "driftsinntekter 2024")  # D
  safe_write(resultat_ark, driftsinntekter_2023, resultat_rad, 3, "driftsinntekter 2023")  # C
  safe_write(resultat_ark, driftskostnader_2024, resultat_rad, 8, "driftskostnader 2024")  # H
  safe_write(resultat_ark, driftskostnader_2023, resultat_rad, 7, "driftskostnader 2023")  # G
  safe_write(resultat_ark, arsresultat_2024, resultat_rad, 16, "årsresultat 2024")         # P
  safe_write(resultat_ark, arsresultat_2023, resultat_rad, 15, "årsresultat 2023")         # O
  safe_write(resultat_ark, round(driftsmargin_2024, 1), resultat_rad, 20, "driftsmargin 2024")  # T
  safe_write(resultat_ark, round(driftsmargin_2023, 1), resultat_rad, 19, "driftsmargin 2023")  # S

  cat("\n=== SKRIVER TIL BALANSE-ARK ===\n")
  # Følger mappingen: 2024 til kolonne D(4), 2023 til kolonne C(3)
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