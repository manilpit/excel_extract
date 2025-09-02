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

# Slår sammen header-tekst i en kolonne (alle rader over første institusjonsrad),
# slik at vi kan søke på "DRIFTSMARGIN" + "2024" i en sammenslått streng.
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

# Finner kolonner for en gitt metrikk + år (basert på header-strenger)
# pattern_metric: f.eks. "DRIFTSMARGIN", "EGENKAPITALGRAD", "OML[ØO]PSMIDLER"
# year: "2024" / "2023"
find_cols_for_metric_year <- function(header_strings, pattern_metric, year) {
  hits <- which(grepl(pattern_metric, header_strings, perl = TRUE) &
                  grepl(paste0("\\b", year, "\\b"), header_strings))
  if (length(hits) == 0) return(NA_integer_)
  hits[1]
}

# Finn første rad med institusjonsnavn (vi forventer 'Institusjon' i header)
detect_first_inst_row <- function(df, default_row = 6) {
  # Se etter ordet "Institusjon" i de første 10 radene, i kolonne 1
  for (r in 1:min(10, nrow(df))) {
    cell <- normalize(df[r, 1])
    if (grepl("^INSTITUSJON", cell)) {
      # Antar at institusjonsnavnene starter i neste rad
      return(r + 1L)
    }
  }
  default_row
}

# Slår opp rad for en institusjon ved å matche mot kolonne A i målarket (rad start_row->)
match_institution_row <- function(wb_path, sheet, inst_name, start_row = 6, name_col = 1) {
  tab <- read.xlsx(wb_path, sheet = sheet, colNames = FALSE)
  names_vec <- normalize(tab[start_row:nrow(tab), name_col])
  target <- normalize(inst_name)

  # Eksakt match først
  idx <- which(names_vec == target)
  if (length(idx) >= 1) return(start_row + idx[1] - 1)

  # Nærmeste match (tolerer små forskjeller)
  dist <- adist(target, names_vec)
  best <- which.min(dist)
  if (length(best) == 1 && is.finite(dist[best]) && dist[best] <= 3) {
    message(sprintf("MERK: '%s' ikke funnet eksakt i '%s'; brukte nærmeste '%s' (avstand %d).",
                    inst_name, sheet, tab[start_row + best - 1, name_col], dist[best]))
    return(start_row + best - 1)
  }

  NA_integer_
}

# Finn rader for relevante etiketter i KILDE-arket (bruk regex som matcher dine kilder)
find_source_rows <- function(src_df) {
  find_row <- function(pat) {
    i <- which(grepl(pat, normalize(src_df[[1]]), perl = TRUE))
    if (length(i) == 0) return(NA_integer_) else i[1]
  }
  list(
    driftsinntekter  = find_row("^DRIFTSINNTEKTER\\b"),
    driftskostnader  = find_row("^TOTALE?\\s+DRIFTSKOSTNADER\\b|^DRIFTSKOSTNADER\\b"),
    driftsresultat   = find_row("^DRIFTSRESULTAT\\b"),
    arsresultat      = find_row("^ÅRSRESULTAT\\b|^AARSRESULTAT\\b"),
    omlopsmidler     = find_row("^OML[ØO]PSMIDLER\\b"),
    egenkapital      = find_row("^EGENKAPITAL\\b"),
    korts_gjeld      = find_row("^KORTSIKTIG\\s+GJELD\\b"),
    totalkapital     = find_row("^TOTALE?\\s+EIENDELER\\b|^TOTALKAPITAL\\b|^SUM\\s+EIENDELER\\b")
  )
}

# Henter fagskolens navn fra kilde (tilpass hvis ligger annet sted enn [1,1])
get_school_name_from_source <- function(src_df) {
  first <- as.character(src_df[1, 1])
  nm <- sub("^\\s*FAGSKOLENS NAVN:\\s*", "", normalize(first))
  nm <- ifelse(nchar(nm) == 0, normalize(first), nm)
  nm
}

# ------- Hovedfunksjon -------
excel_overforing_fagskole <- function(
  kilde_fil,
  output_fil,
  resultat_ark = "Tab_resultat",   # ev. "Tab resultat"
  balanse_ark  = "Tab_balanse",    # ev. "balanse"
  # Valgfritt: oppgi rader manuelt som i høgskole-scriptet (ellers finner vi automatisk)
  resultat_rad = NULL,
  balanse_rad  = NULL,
  # Hvis overskriftene i målarket ligger på topp-rader 1:5 (institusjoner fra rad 6)
  header_rows_guess = 1:5,
  # Hvilket ark i kildefilen data ligger på
  kilde_ark = 1
) {

  # --- Les kilden ---
  src <- read.xlsx(kilde_fil, sheet = kilde_ark, colNames = FALSE)
  if (nrow(src) == 0) stop("Tomt kildeark: ", kilde_fil)

  # Finn navn og kilderader
  inst_name <- get_school_name_from_source(src)
  rmap <- find_source_rows(src)

  # Finn år-kolonner ved å se etter 2024/2023 i de første 10 radene
  # (Hvis kildearkene er rene tabeller med overskrift-rad for år)
  find_year_col <- function(year) {
    for (r in 1:min(10, nrow(src))) {
      row_vals <- normalize(unlist(src[r, , drop = FALSE]))
      hits <- which(grepl(paste0("^", year, "$"), row_vals))
      if (length(hits) >= 1) return(hits[1])
    }
    # fallback: klassisk 2024=kol 3, 2023=kol 4
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

  # --- Åpne målfil ---
  wb <- loadWorkbook(output_fil)

  # Les arkene som raw for å finne header-blokk og kolonner:
  res_tab <- read.xlsx(output_fil, sheet = resultat_ark, colNames = FALSE)
  bal_tab <- read.xlsx(output_fil, sheet = balanse_ark,  colNames = FALSE)

  # Finn første institusjonsrad (normalt 6)
  res_first <- detect_first_inst_row(res_tab, default_row = 6)
  bal_first <- detect_first_inst_row(bal_tab, default_row = 6)

  # Header-strenger (alle rader over første inst-rad)
  res_headers <- make_header_strings(res_tab, header_rows = 1:(res_first - 1))
  bal_headers <- make_header_strings(bal_tab, header_rows = 1:(bal_first - 1))

  # Finn skrive-kolonner i RESULTAT
  c_res_dm_2024 <- find_cols_for_metric_year(res_headers, "DRIFTSMARGIN", "2024")
  c_res_dm_2023 <- find_cols_for_metric_year(res_headers, "DRIFTSMARGIN", "2023")
  # Disse 4 bruker vi hvis du også vil lagre komponentene (valgfritt)
  c_res_di_2024 <- find_cols_for_metric_year(res_headers, "TOTALE?\\s+DRIFTSINNTEKTER|DRIFTSINNTEKTER", "2024")
  c_res_di_2023 <- find_cols_for_metric_year(res_headers, "TOTALE?\\s+DRIFTSINNTEKTER|DRIFTSINNTEKTER", "2023")

  # Finn skrive-kolonner i BALANSE
  c_bal_om_2024  <- find_cols_for_metric_year(bal_headers, "OML[ØO]PSMIDLER", "2024")
  c_bal_om_2023  <- find_cols_for_metric_year(bal_headers, "OML[ØO]PSMIDLER", "2023")
  c_bal_ek_2024  <- find_cols_for_metric_year(bal_headers, "EGENKAPITAL(?!GRAD)", "2024")
  c_bal_ek_2023  <- find_cols_for_metric_year(bal_headers, "EGENKAPITAL(?!GRAD)", "2023")
  c_bal_kg_2024  <- find_cols_for_metric_year(bal_headers, "KORTSIKTIG\\s+GJELD", "2024")
  c_bal_kg_2023  <- find_cols_for_metric_year(bal_headers, "KORTSIKTIG\\s+GJELD", "2023")
  c_bal_tot_2024 <- find_cols_for_metric_year(bal_headers, "TOTALKAPITAL|TOTALE?\\s+EIENDELER|SUM\\s+EIENDELER", "2024")
  c_bal_tot_2023 <- find_cols_for_metric_year(bal_headers, "TOTALKAPITAL|TOTALE?\\s+EIENDELER|SUM\\s+EIENDELER", "2023")
  c_bal_ekg_2024 <- find_cols_for_metric_year(bal_headers, "EGENKAPITALGRAD", "2024")
  c_bal_ekg_2023 <- find_cols_for_metric_year(bal_headers, "EGENKAPITALGRAD", "2023")
  c_bal_f2_2024  <- find_cols_for_metric_year(bal_headers, "FINANSIERINGSGRAD\\s*2|LIKVIDITETSGRAD", "2024")
  c_bal_f2_2023  <- find_cols_for_metric_year(bal_headers, "FINANSIERINGSGRAD\\s*2|LIKVIDITETSGRAD", "2023")

  # Finn rad i målarket (hvis ikke gitt)
  if (is.null(resultat_rad)) {
    resultat_rad <- match_institution_row(output_fil, resultat_ark, inst_name, start_row = res_first)
    if (is.na(resultat_rad)) stop("Fant ikke rad for '", inst_name, "' i ", resultat_ark)
  }
  if (is.null(balanse_rad)) {
    # antar samme rad i begge ark
    balanse_rad <- match_institution_row(output_fil, balanse_ark, inst_name, start_row = bal_first)
    if (is.na(balanse_rad)) balanse_rad <- resultat_rad
  }

  # --- Skriv verdier ---
  # Resultat-arket: skriv driftsmargin (du kan også velge å skrive DI, DK, DR, ÅR dersom mal skal vise tall)
  if (!is.na(c_res_dm_2024)) writeData(wb, sheet = resultat_ark, x = round(k2024$driftsmargin, 1), startRow = resultat_rad, startCol = c_res_dm_2024)
  if (!is.na(c_res_dm_2023)) writeData(wb, sheet = resultat_ark, x = round(k2023$driftsmargin, 1), startRow = resultat_rad, startCol = c_res_dm_2023)

  # (Valgfritt – bare hvis du ønsker å fylle komponenter i resultatarket)
  if (!is.na(c_res_di_2024)) writeData(wb, sheet = resultat_ark, x = v2024$driftsinntekter, startRow = resultat_rad, startCol = c_res_di_2024)
  if (!is.na(c_res_di_2023)) writeData(wb, sheet = resultat_ark, x = v2023$driftsinntekter, startRow = resultat_rad, startCol = c_res_di_2023)

  # Balanse-arket: skriv komponenter + indikatorer
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

  saveWorkbook(wb, output_fil, overwrite = TRUE)

  cat(sprintf("OK: %s -> rad %d (%s) / rad %d (%s)\n",
              inst_name, resultat_rad, resultat_ark, balanse_rad, balanse_ark))
}