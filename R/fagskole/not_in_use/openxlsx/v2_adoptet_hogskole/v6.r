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

# Ny funksjon for å finne kolonner basert på metrikk og år i separate rader
find_metric_year_column <- function(df, metric_row = 3, year_row = 4, metric_pattern, year) {
  # Finn først kolonner som matcher metrikken
  metric_cols <- which(grepl(metric_pattern, normalize(as.character(df[metric_row,])), perl = TRUE))
  
  if (length(metric_cols) == 0) {
    cat("Fant ikke kolonne for metrikk:", metric_pattern, "\n")
    return(NA_integer_)
  }
  
  cat("Fant metrikk-kolonner:", metric_cols, "for", metric_pattern, "\n")
  
  # For hver metrikk-kolonne, sjekk de neste kolonnene for året
  for (base_col in metric_cols) {
    # Sjekk de neste 4 kolonnene fra metrikk-kolonnen
    year_candidates <- base_col:(base_col + 3)
    for (col in year_candidates) {
      if (col <= ncol(df)) {
        year_cell <- normalize(as.character(df[year_row, col]))
        cat("Sjekker kolonne", col, "for år", year, "- fant:", year_cell, "\n")
        if (grepl(paste0("^", year, "$"), year_cell)) {
          cat("MATCH: Fant", metric_pattern, "for", year, "i kolonne", col, "\n")
          return(col)
        }
      }
    }
  }
  
  cat("Fant ikke år", year, "for", metric_pattern, "\n")
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
    for (col in 1:min(3, ncol(src_df))) {  # Sjekk de første kolonnene
      i <- which(grepl(pat, normalize(src_df[[col]]), perl = TRUE))
      if (length(i) > 0) {
        cat("Fant rad", pat, ":", i[1], "i kolonne", col, "\n")
        return(i[1])
      }
    }
    cat("Fant ikke rad for:", pat, "\n")
    return(NA_integer_)
  }
  
  rader <- list(
    driftsinntekter  = find_row("^DRIFTSINNTEKTER\\b|^SALGSINNTEKTER\\b|^TOTALE?\\s+DRIFTSINNTEKTER\\b"),
    driftskostnader  = find_row("^TOTALE?\\s+DRIFTSKOSTNADER\\b|^DRIFTSKOSTNADER\\b"),
    driftsresultat   = find_row("^DRIFTSRESULTAT\\b"),
    arsresultat      = find_row("^ÅRSRESULTAT\\b|^AARSRESULTAT\\b"),
    omlopsmidler     = find_row("^OML[ØO]PSMIDLER\\b|^A\\.\\s*OML[ØO]PSMIDLER\\b"),
    egenkapital      = find_row("^EGENKAPITAL\\b|^C\\.\\s*EGENKAPITAL\\b"),
    korts_gjeld      = find_row("^KORTSIKTIG\\s+GJELD\\b|^E\\.\\s*KORTSIKTIG\\s+GJELD\\b"),
    totalkapital     = find_row("^TOTALE?\\s+EIENDELER\\b|^TOTALKAPITAL\\b|^SUM\\s+EIENDELER\\b|^SUMMEN\\s+EIENDELER\\b")
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

  # Finn år-kolonner i kildefilen
  find_year_col <- function(year) {
    for (r in 1:min(10, nrow(src))) {
      row_vals <- normalize(unlist(src[r, , drop = FALSE]))
      hits <- which(grepl(paste0("^", year, "$"), row_vals))
      if (length(hits) >= 1) {
        cat("Fant år", year, "i rad", r, "kolonne", hits[1], "\n")
        return(hits[1])
      }
    }
    cat("Fant ikke år", year, "- bruker fallback\n")
    if (year == "2024") 3L else if (year == "2023") 4L else NA_integer_
  }
  col2024 <- find_year_col("2024")
  col2023 <- find_year_col("2023")

  # Hent verdier fra kilden
  val <- function(row, col) {
    if (is.na(row) || is.na(col) || row > nrow(src) || col > ncol(src)) {
      return(NA_real_)
    }
    return(to_num(src[row, col]))
  }
  
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
  cat("2024:\n"); print(v2024)
  cat("2023:\n"); print(v2023)

  # Nøkkeltall
  safe_div <- function(num, den) {
    if (is.na(num) | is.na(den) | den == 0) {
      return(NA_real_)
    } else {
      return(num / den)
    }
  }
  
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
  cat("2024:\n"); print(k2024)
  cat("2023:\n"); print(k2023)

  # Åpne målfil
  wb <- loadWorkbook(output_fil)
  if (is.null(wb)) {
    stop("Kunne ikke åpne workbook: ", output_fil)
  }
  cat("\nWorkbook åpnet successfully\n")

  # Les arkene som raw
  res_tab <- read.xlsx(output_fil, sheet = resultat_ark, colNames = FALSE)
  bal_tab <- read.xlsx(output_fil, sheet = balanse_ark,  colNames = FALSE)

  # Debug: Skriv ut strukturen
  cat("\nAnalyserer målark-struktur...\n")
  cat("Resultat-ark rad 3:", paste(res_tab[3,], collapse=" | "), "\n")
  cat("Resultat-ark rad 4:", paste(res_tab[4,], collapse=" | "), "\n")
  cat("Balanse-ark rad 3:", paste(bal_tab[3,], collapse=" | "), "\n")
  cat("Balanse-ark rad 4:", paste(bal_tab[4,], collapse=" | "), "\n")

  # Finn kolonner med ny metode
  cat("\n=== SØKER ETTER KOLONNER I RESULTAT-ARK ===\n")
  c_res_dm_2024 <- find_metric_year_column(res_tab, 3, 4, "DRIFTSMARGIN", "2024")
  c_res_dm_2023 <- find_metric_year_column(res_tab, 3, 4, "DRIFTSMARGIN", "2023")
  c_res_di_2024 <- find_metric_year_column(res_tab, 3, 4, "TOTALE\\s+DRIFTSINNTEKTER|DRIFTSINNTEKTER", "2024")
  c_res_di_2023 <- find_metric_year_column(res_tab, 3, 4, "TOTALE\\s+DRIFTSINNTEKTER|DRIFTSINNTEKTER", "2023")

  cat("\n=== SØKER ETTER KOLONNER I BALANSE-ARK ===\n")
  c_bal_om_2024  <- find_metric_year_column(bal_tab, 3, 4, "OML[ØO]PSMIDLER", "2024")
  c_bal_om_2023  <- find_metric_year_column(bal_tab, 3, 4, "OML[ØO]PSMIDLER", "2023")
  c_bal_ek_2024  <- find_metric_year_column(bal_tab, 3, 4, "^EGENKAPITAL$", "2024")
  c_bal_ek_2023  <- find_metric_year_column(bal_tab, 3, 4, "^EGENKAPITAL$", "2023")
  c_bal_kg_2024  <- find_metric_year_column(bal_tab, 3, 4, "KORTSIKTIG\\s+GJELD", "2024")
  c_bal_kg_2023  <- find_metric_year_column(bal_tab, 3, 4, "KORTSIKTIG\\s+GJELD", "2023")
  c_bal_tot_2024 <- find_metric_year_column(bal_tab, 3, 4, "TOTALKAPITAL", "2024")
  c_bal_tot_2023 <- find_metric_year_column(bal_tab, 3, 4, "TOTALKAPITAL", "2023")
  c_bal_ekg_2024 <- find_metric_year_column(bal_tab, 3, 4, "EGENKAPITALGRAD", "2024")
  c_bal_ekg_2023 <- find_metric_year_column(bal_tab, 3, 4, "EGENKAPITALGRAD", "2023")
  c_bal_f2_2024  <- find_metric_year_column(bal_tab, 3, 4, "FINANSIERINGSGRAD\\s*2|LIKVIDITETSGRAD", "2024")
  c_bal_f2_2023  <- find_metric_year_column(bal_tab, 3, 4, "FINANSIERINGSGRAD\\s*2|LIKVIDITETSGRAD", "2023")

  # Debug: Skriv ut funne kolonner
  cat("\n=== FUNNET KOLONNER ===\n")
  cat("Resultat-ark:\n")
  cat("  Driftsmargin 2024:", c_res_dm_2024, "\n")
  cat("  Driftsmargin 2023:", c_res_dm_2023, "\n")
  cat("  Driftsinntekter 2024:", c_res_di_2024, "\n")
  cat("  Driftsinntekter 2023:", c_res_di_2023, "\n")
  cat("Balanse-ark:\n")
  cat("  Omløpsmidler 2024:", c_bal_om_2024, "\n")
  cat("  Omløpsmidler 2023:", c_bal_om_2023, "\n")
  cat("  Egenkapital 2024:", c_bal_ek_2024, "\n")
  cat("  Egenkapital 2023:", c_bal_ek_2023, "\n")

  # Funksjon for sikker skriving
  safe_write <- function(sheet, value, row, col, description) {
    if (!is.na(col) && !is.na(value)) {
      tryCatch({
        writeData(wb, sheet = sheet, x = as.vector(value), startRow = row, startCol = col)
        cat("*** SKREV", description, ":", value, "til rad", row, "kolonne", col, "***\n")
      }, error = function(e) {
        cat("FEIL ved skriving av", description, ":", conditionMessage(e), "\n")
      })
    } else {
      cat("Hopper over", description, "- kolonne:", col, "verdi:", value, "\n")
    }
  }

  cat("\n=== SKRIVER VERDIER ===\n")
  
  # Skriv verdier til resultat-ark
  safe_write(resultat_ark, round(k2024$driftsmargin, 1), resultat_rad, c_res_dm_2024, "driftsmargin 2024")
  safe_write(resultat_ark, round(k2023$driftsmargin, 1), resultat_rad, c_res_dm_2023, "driftsmargin 2023")
  safe_write(resultat_ark, v2024$driftsinntekter, resultat_rad, c_res_di_2024, "driftsinntekter 2024")
  safe_write(resultat_ark, v2023$driftsinntekter, resultat_rad, c_res_di_2023, "driftsinntekter 2023")

  # Skriv verdier til balanse-ark
  safe_write(balanse_ark, v2024$omlopsmidler, balanse_rad, c_bal_om_2024, "omløpsmidler 2024")
  safe_write(balanse_ark, v2023$omlopsmidler, balanse_rad, c_bal_om_2023, "omløpsmidler 2023")
  safe_write(balanse_ark, v2024$egenkapital, balanse_rad, c_bal_ek_2024, "egenkapital 2024")
  safe_write(balanse_ark, v2023$egenkapital, balanse_rad, c_bal_ek_2023, "egenkapital 2023")
  safe_write(balanse_ark, v2024$korts_gjeld, balanse_rad, c_bal_kg_2024, "kortsiktig gjeld 2024")
  safe_write(balanse_ark, v2023$korts_gjeld, balanse_rad, c_bal_kg_2023, "kortsiktig gjeld 2023")
  safe_write(balanse_ark, v2024$totalkapital, balanse_rad, c_bal_tot_2024, "totalkapital 2024")
  safe_write(balanse_ark, v2023$totalkapital, balanse_rad, c_bal_tot_2023, "totalkapital 2023")
  safe_write(balanse_ark, round(k2024$egenkapitalgrad, 1), balanse_rad, c_bal_ekg_2024, "egenkapitalgrad 2024")
  safe_write(balanse_ark, round(k2023$egenkapitalgrad, 1), balanse_rad, c_bal_ekg_2023, "egenkapitalgrad 2023")
  safe_write(balanse_ark, round(k2024$finansieringsgrad2, 1), balanse_rad, c_bal_f2_2024, "finansieringsgrad2 2024")
  safe_write(balanse_ark, round(k2023$finansieringsgrad2, 1), balanse_rad, c_bal_f2_2023, "finansieringsgrad2 2023")

  # Lagre workbook
  tryCatch({
    saveWorkbook(wb, output_fil, overwrite = TRUE)
    cat("\n*** WORKBOOK LAGRET SUCCESSFULLY ***\n")
  }, error = function(e) {
    stop("Feil ved lagring av workbook: ", conditionMessage(e))
  })

  cat(sprintf("*** FERDIG: %s -> rad %d (%s) / rad %d (%s) ***\n",
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