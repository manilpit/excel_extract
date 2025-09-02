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

make_header_strings <- function(df, header_rows) {
  cols <- ncol(df)
  headers <- character(cols)
  for (c in seq_len(cols)) {
    v <- df[header_rows, c]
    v <- v[!is.na(v) & nzchar(as.character(v))]
    headers[c] <- normalize(paste(v, collapse = " | "))
  }
  headers
}

find_cols_for_metric_year <- function(df, metric_row = 3, year_row = 4, metric_pattern, year) {
  # Finn først kolonnen for metrikken
  metric_cols <- which(grepl(metric_pattern, normalize(as.character(df[metric_row,])), perl = TRUE))
  
  if (length(metric_cols) == 0) {
    cat("\nFant ikke kolonne for:", metric_pattern, "\n")
    return(NA_integer_)
  }
  
  # For hver metrikk-kolonne, finn tilhørende år-kolonne
  for (base_col in metric_cols) {
    # Året skal være i en av de neste kolonnene
    year_candidates <- base_col + 0:3  # Sjekk basiskol + 3 neste
    for (col in year_candidates) {
      if (col <= ncol(df)) {
        if (grepl(paste0("^", year, "$"), normalize(as.character(df[year_row, col])))) {
          cat("\nFant", metric_pattern, "for", year, "i kolonne", col, "\n")
          return(col)
        }
      }
    }
  }
  
  cat("\nFant ikke år", year, "for", metric_pattern, "\n")
  return(NA_integer_)
}

detect_first_inst_row <- function(df, default_row = 6) {
  for (r in 1:min(10, nrow(df))) {
    cell <- normalize(df[r, 1])
    if (grepl("^INSTITUSJON", cell)) {
      return(r + 1L)
    }
  }
  default_row
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
  cat("Første 5 navn i målarket:", paste(head(names_vec, 5), collapse=", "), "...\n")
  
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

find_source_rows <- function(src_df) {
  find_row <- function(pat) {
    i <- which(grepl(pat, normalize(src_df[[1]]), perl = TRUE))
    if (length(i) == 0) {
      cat("\nFant ikke rad for:", pat, "\n")
      return(NA_integer_)
    }
    cat("\nFant rad", pat, ":", i[1], "\n")
    return(i[1])
  }
  
  rader <- list(
    driftsinntekter  = find_row("^DRIFTSINNTEKTER\\b|^SALGSINNTEKTER\\b"),
    driftskostnader  = find_row("^TOTALE?\\s+DRIFTSKOSTNADER\\b|^DRIFTSKOSTNADER\\b"),
    driftsresultat   = find_row("^DRIFTSRESULTAT\\b"),
    arsresultat      = find_row("^ÅRSRESULTAT\\b|^AARSRESULTAT\\b"),
    omlopsmidler     = find_row("^OML[ØO]PSMIDLER\\b|^A\\.\\s*OML[ØO]PSMIDLER\\b"),
    egenkapital      = find_row("^EGENKAPITAL\\b|^C\\.\\s*EGENKAPITAL\\b"),
    korts_gjeld      = find_row("^KORTSIKTIG\\s+GJELD\\b|^E\\.\\s*KORTSIKTIG\\s+GJELD\\b"),
    totalkapital     = find_row("^TOTALE?\\s+EIENDELER\\b|^TOTALKAPITAL\\b|^SUM\\s+EIENDELER\\b")
  )
  
  # Debug utskrift
  cat("\nFunnet rader:\n")
  print(rader)
  
  return(rader)
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
# ------- Hovedfunksjon -------
excel_overforing_fagskole <- function(
  kilde_fil,
  output_fil,
  resultat_ark = "Tab_resultat",
  balanse_ark = "Tab_balanse",
  resultat_rad = NULL,
  balanse_rad = NULL,
  header_rows_guess = 1:5,
  kilde_ark = "Resultatregnskap"
) {
  # Få skolenavnet fra Resultatregnskap-arket
  inst_name <- get_school_name_from_source(kilde_fil, sheet_name = kilde_ark)
  cat("Fant skolenavn:", inst_name, "\n")
  
  # Les inn kilde-arket
  tryCatch({
    src <- read.xlsx(kilde_fil, sheet = kilde_ark, colNames = FALSE)
    if (nrow(src) == 0) stop("Tomt kildeark: ", kilde_ark, " i ", kilde_fil)
  }, error = function(e) {
    stop("Kunne ikke lese arket '", kilde_ark, "' i ", kilde_fil, ": ", conditionMessage(e))
  })
  
  # Finn kilderader
  rmap <- find_source_rows(src)
  
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
  
  cat("OK:", inst_name, "->", "rad", resultat_rad, "(", resultat_ark, ")", "/", 
      "rad", balanse_rad, "(", balanse_ark, ")", "\n")

  # Finn år-kolonner
  find_year_col <- function(year) {
    for (r in 1:min(10, nrow(src))) {
      row_vals <- normalize(unlist(src[r, , drop = FALSE]))
      hits <- which(grepl(paste0("^", year, "$"), row_vals))
      if (length(hits) >= 1) return(hits[1])
    }
    if (year == "2024") 3L else if (year == "2023") 4L else NA_integer_
  }
  col2024 <- find_year_col("2024")
  col2023 <- find_year_col("2023")

  # Hent verdier fra kilden
  val <- function(row, col) if (is.na(row) || is.na(col)) NA_real_ else to_num(src[row, col])
  v2024 <- list(
    driftsinntekter  = val(rmap$driftsinntekter, col2024),
    driftskostnader  = val(rmap$driftskostnader, col2024),
    driftsresultat   = val(rmap$driftsresultat,  col2024),
    arsresultat      = val(rmap$arsresultat,     col2024),
    omlopsmidler     = val(rmap$omlopsmidler,    col2024),
    egenkapital      = val(rmap$egenkapital,     col2024),
    korts_gjeld      = val(rmap$korts_gjeld,     col2024),
    totalkapital     = val(rmap$totalkapital,    col2024)
  )
  v2023 <- list(
    driftsinntekter  = val(rmap$driftsinntekter, col2023),
    driftskostnader  = val(rmap$driftskostnader, col2023),
    driftsresultat   = val(rmap$driftsresultat,  col2023),
    arsresultat      = val(rmap$arsresultat,     col2023),
    omlopsmidler     = val(rmap$omlopsmidler,    col2023),
    egenkapital      = val(rmap$egenkapital,     col2023),
    korts_gjeld      = val(rmap$korts_gjeld,     col2023),
    totalkapital     = val(rmap$totalkapital,    col2023)
  )

  # Debug: Skriv ut kildeverdier
  cat("\nHentet verdier fra kildefil:\n")
  print(v2024)
  print(v2023)
# Etter at verdier er hentet
cat("\nVerdier som skal skrives til resultat-ark:\n")
cat("Rad:", resultat_rad, "\n")
cat("Driftsmargin 2024 (kolonne", c_res_dm_2024, "):", k2024$driftsmargin, "\n")
cat("Driftsmargin 2023 (kolonne", c_res_dm_2023, "):", k2023$driftsmargin, "\n")

# Før hver skriveoperasjon
tryCatch({
  writeData(wb, sheet = resultat_ark, 
           x = as.vector(round(k2024$driftsmargin, 1)), 
           startRow = resultat_rad, 
           startCol = c_res_dm_2024)
  cat("Skrev driftsmargin 2024 OK\n")
}, error = function(e) {
  cat("FEIL ved skriving av driftsmargin 2024:", conditionMessage(e), "\n")
})
  # Nøkkeltall
  safe_div <- function(num, den) ifelse(is.na(num) | is.na(den) | den == 0, NA_real_, num / den)
  k2024 <- list(
    driftsmargin       = safe_div(v2024$driftsresultat,  v2024$driftsinntekter) * 100,
    egenkapitalgrad    = safe_div(v2024$egenkapital,     v2024$totalkapital) * 100,
    finansieringsgrad2 = safe_div(v2024$omlopsmidler,    v2024$korts_gjeld) * 100
  )
  k2023 <- list(
    driftsmargin       = safe_div(v2023$driftsresultat,  v2023$driftsinntekter) * 100,
    egenkapitalgrad    = safe_div(v2023$egenkapital,     v2023$totalkapital) * 100,
    finansieringsgrad2 = safe_div(v2023$omlopsmidler,    v2023$korts_gjeld) * 100
  )

  # Debug: Skriv ut beregnede nøkkeltall
  cat("\nBeregnede nøkkeltall:\n")
  print(k2024)
  print(k2023)

  # Åpne målfil
  wb <- loadWorkbook(output_fil)
  if (is.null(wb)) {
    stop("Kunne ikke åpne workbook: ", output_fil)
  }
  cat("\nWorkbook åpnet successfully\n")

  # Les arkene som raw
  res_tab <- read.xlsx(output_fil, sheet = resultat_ark, colNames = FALSE)
  bal_tab <- read.xlsx(output_fil, sheet = balanse_ark,  colNames = FALSE)

  # Legg til denne debuggingen i hovedfunksjonen rett etter at res_tab og bal_tab leses inn:
cat("\nRå headers i resultat-ark (første 5 rader):\n")
for(i in 1:5) {
  cat("Rad", i, ":", paste(res_tab[i,], collapse=" | "), "\n")
}

cat("\nRå headers i balanse-ark (første 5 rader):\n")
for(i in 1:5) {
  cat("Rad", i, ":", paste(bal_tab[i,], collapse=" | "), "\n")
}
  # Finn første institusjonsrad
  res_first <- detect_first_inst_row(res_tab, default_row = 6)
  bal_first <- detect_first_inst_row(bal_tab, default_row = 6)

  # Header-strenger
  res_headers <- make_header_strings(res_tab, header_rows = 1:(res_first - 1))
  bal_headers <- make_header_strings(bal_tab, header_rows = 1:(bal_first - 1))

  # Debug: Skriv ut headers
  cat("\nFunnet headers i resultat-ark:\n")
  print(res_headers)
  cat("\nFunnet headers i balanse-ark:\n")
  print(bal_headers)

  # Finn kolonner
  # Finn kolonner med nye parametre
  c_res_dm_2024 <- find_cols_for_metric_year(res_tab, metric_pattern = "DRIFTSMARGIN", year = "2024")
  c_res_dm_2023 <- find_cols_for_metric_year(res_tab, metric_pattern = "DRIFTSMARGIN", year = "2023")
  c_res_di_2024 <- find_cols_for_metric_year(res_tab, metric_pattern = "DRIFTSINNTEKTER", year = "2024")
  c_res_di_2023 <- find_cols_for_metric_year(res_tab, metric_pattern = "DRIFTSINNTEKTER", year = "2023")

  c_bal_om_2024  <- find_cols_for_metric_year(bal_tab, metric_pattern = "OML[ØO]PSMIDLER", year = "2024")
  c_bal_om_2023  <- find_cols_for_metric_year(bal_tab, metric_pattern = "OML[ØO]PSMIDLER", year = "2023")
  c_bal_ek_2024  <- find_cols_for_metric_year(bal_tab, metric_pattern = "EGENKAPITAL(?!GRAD)", year = "2024")
  c_bal_ek_2023  <- find_cols_for_metric_year(bal_tab, metric_pattern = "EGENKAPITAL(?!GRAD)", year = "2023")
  c_bal_tot_2024 <- find_cols_for_metric_year(bal_headers, "TOTALKAPITAL|TOTALE?\\s+EIENDELER|SUM\\s+EIENDELER", "2024")
  c_bal_tot_2023 <- find_cols_for_metric_year(bal_headers, "TOTALKAPITAL|TOTALE?\\s+EIENDELER|SUM\\s+EIENDELER", "2023")
  c_bal_ekg_2024 <- find_cols_for_metric_year(bal_headers, "EGENKAPITALGRAD", "2024")
  c_bal_ekg_2023 <- find_cols_for_metric_year(bal_headers, "EGENKAPITALGRAD", "2023")
  c_bal_f2_2024  <- find_cols_for_metric_year(bal_headers, "FINANSIERINGSGRAD\\s*2|LIKVIDITETSGRAD", "2024")
  c_bal_f2_2023  <- find_cols_for_metric_year(bal_headers, "FINANSIERINGSGRAD\\s*2|LIKVIDITETSGRAD", "2023")

  # Debug: Skriv ut funne kolonner
  cat("\nFunnet kolonner:\n")
  cat("Resultat-ark:\n")
  cat("  Driftsmargin 2024:", c_res_dm_2024, "\n")
  cat("  Driftsmargin 2023:", c_res_dm_2023, "\n")
  cat("Balanse-ark:\n")
  cat("  Omløpsmidler 2024:", c_bal_om_2024, "\n")
  cat("  Omløpsmidler 2023:", c_bal_om_2023, "\n")

  # Sjekk om vi fant noen kolonner
  if (all(is.na(c(c_res_dm_2024, c_res_dm_2023, c_bal_om_2024, c_bal_om_2023)))) {
    stop("Ingen kolonner ble funnet i målarkene. Sjekk header-matching.")
  }

  # Debug: Skriv hvilke rader/kolonner vi skal skrive til
  cat("\nSkriver til rader/kolonner:\n")
  cat("Resultat-ark rad:", resultat_rad, "\n")
  cat("Balanse-ark rad:", balanse_rad, "\n")

  # Skriv verdier
  if (!is.na(c_res_dm_2024)) {
    cat("\nSkriver driftsmargin 2024:", round(k2024$driftsmargin, 1), "til rad", resultat_rad, "kolonne", c_res_dm_2024, "\n")
  #   writeData(wb, sheet = resultat_ark, x = round(k2024$driftsmargin, 1), startRow = resultat_rad, startCol = c_res_dm_2024)
    # Før skriving, konverter verdiene til vektorer
  writeData(wb, sheet = resultat_ark, x = as.vector(round(k2024$driftsmargin, 1)), 
  startRow = resultat_rad, startCol = c_res_dm_2024)
  }
  if (!is.na(c_res_dm_2023)) writeData(wb, sheet = resultat_ark, x = round(k2023$driftsmargin, 1), startRow = resultat_rad, startCol = c_res_dm_2023)
  if (!is.na(c_res_di_2024)) writeData(wb, sheet = resultat_ark, x = v2024$driftsinntekter, startRow = resultat_rad, startCol = c_res_di_2024)
  if (!is.na(c_res_di_2023)) writeData(wb, sheet = resultat_ark, x = v2023$driftsinntekter, startRow = resultat_rad, startCol = c_res_di_2023)

  if (!is.na(c_bal_om_2024))  writeData(wb, sheet = balanse_ark, x = v2024$omlopsmidler,     startRow = balanse_rad, startCol = c_bal_om_2024)
  if (!is.na(c_bal_om_2023))  writeData(wb, sheet = balanse_ark, x = v2023$omlopsmidler,     startRow = balanse_rad, startCol = c_bal_om_2023)
  if (!is.na(c_bal_ek_2024))  writeData(wb, sheet = balanse_ark, x = v2024$egenkapital,      startRow = balanse_rad, startCol = c_bal_ek_2024)
  if (!is.na(c_bal_ek_2023))  writeData(wb, sheet = balanse_ark, x = v2023$egenkapital,      startRow = balanse_rad, startCol = c_bal_ek_2023)
  if (!is.na(c_bal_kg_2024))  writeData(wb, sheet = balanse_ark, x = v2024$korts_gjeld,      startRow = balanse_rad, startCol = c_bal_kg_2024)
  if (!is.na(c_bal_kg_2023))  writeData(wb, sheet = balanse_ark, x = v2023$korts_gjeld,      startRow = balanse_rad, startCol = c_bal_kg_2023)
  if (!is.na(c_bal_tot_2024)) writeData(wb, sheet = balanse_ark, x = v2024$totalkapital,     startRow = balanse_rad, startCol = c_bal_tot_2024)
  if (!is.na(c_bal_tot_2023)) writeData(wb, sheet = balanse_ark, x = v2023$totalkapital,     startRow = balanse_rad, startCol = c_bal_tot_2023)
  if (!is.na(c_bal_ekg_2024)) writeData(wb, sheet = balanse_ark, x = round(k2024$egenkapitalgrad, 1),    startRow = balanse_rad, startCol = c_bal_ekg_2024)
  if (!is.na(c_bal_ekg_2023)) writeData(wb, sheet = balanse_ark, x = round(k2023$egenkapitalgrad, 1),    startRow = balanse_rad, startCol = c_bal_ekg_2023)
  if (!is.na(c_bal_f2_2024))  writeData(wb, sheet = balanse_ark, x = round(k2024$finansieringsgrad2, 1), startRow = balanse_rad, startCol = c_bal_f2_2024)
  if (!is.na(c_bal_f2_2023))  writeData(wb, sheet = balanse_ark, x = round(k2023$finansieringsgrad2, 1), startRow = balanse_rad, startCol = c_bal_f2_2023)

  # Lagre workbook
  tryCatch({
    saveWorkbook(wb, output_fil, overwrite = TRUE)
    cat("\nWorkbook lagret successfully\n")
  }, error = function(e) {
    stop("Feil ved lagring av workbook: ", conditionMessage(e))
  })

  cat(sprintf("OK: %s -> rad %d (%s) / rad %d (%s)\n",
              inst_name, resultat_rad, resultat_ark, balanse_rad, balanse_ark))
  
  # Returner verdier for inspeksjon
  invisible(list(
    name = inst_name,
    values_2024 = v2024,
    values_2023 = v2023,
    key_figures_2024 = k2024,
    key_figures_2023 = k2023
  ))
}