library(openxlsx)
library(data.table)

# ------- Hjelpere -------
normalize <- function(x) {
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

# Liste over gyldige fagskolenumre som skal behandles
valid_fagskole_ids <- c(
  '197', # AOF Østlandet
  '226', # Campus BLÅ Fagskole AS
  '23',  # Centric IT Academy
  '17',  # Designinstituttet
  '89',  # Det tverrfaglige kunstinstitutt
  '232', # Din Kompetanse fagskole
  '157', # Einar Granum Kunstfagskole
  '15',  # Fabrikken Asker Kunstfagskole
  '236', # Fagskolen Diakonova
  '174', # Fagskolen Essens AS
  '85',  # Fagskolen for bokbransjen
  '237', # Fagskolen GET Academy AS
  '40',  # Fagskolen Kristiania
  '61',  # Fagskolen Tirna
  '216', # Folkeuniversitetets Fagskole AS
  '91',  # Frelsesarmeens offisersskole AS
  '220', # Gokstad akademiet
  '36',  # Hald internasjonale skole
  '218', # KBT-fagskole
  '142', # Kunstfagskolen i Bergen
  '92',  # Kunstskolen i Stavanger AS
  '111', # Lukas høyere yrkesfagskole
  '209', # MedLearn AS
  '104', # Norsk Hestesenter
  '29',  # TISIP Fagskole
  '182'  # Ytre kunstfagskole
)

# Funksjon for å sjekke om en fil skal behandles basert på fagskolenummeret
should_process_file <- function(filename) {
  fagskole_id <- get_fagskole_id_from_filename(filename)
  if (is.na(fagskole_id)) {
    cat("ADVARSEL: Kunne ikke finne fagskolenummer fra filnavn:", basename(filename), "\n")
    return(FALSE)
  }
  
  fagskole_id_str <- as.character(fagskole_id)
  if (fagskole_id_str %in% valid_fagskole_ids) {
    cat("Fil med fagskolenummer", fagskole_id_str, "er på listen og vil bli behandlet\n")
    return(TRUE)
  } else {
    cat("Fil med fagskolenummer", fagskole_id_str, "er IKKE på listen og vil bli hoppet over\n")
    return(FALSE)
  }
}

# Funksjon for å finne rad basert på kode i kolonne E
find_row_by_code <- function(data, code) {
  if (ncol(data) < 5) {
    cat("ADVARSEL: Arket har mindre enn 5 kolonner, kan ikke søke i kolonne E\n")
    return(NA_integer_)
  }
  
  for (row in 1:nrow(data)) {
    cell_value <- data[row, 5]
    if (is.na(cell_value) || is.null(cell_value)) {
      next
    }
    
    cell_value_norm <- normalize(cell_value)
    if (cell_value_norm == code) {
      cat("Fant kode", code, "på rad", row, "\n")
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

# Hent fagskole-ID fra filnavn
get_fagskole_id_from_filename <- function(filename) {
  # Forventet format: "NNN_..." hvor NNN er fagskolenummeret
  id_match <- regexpr("^(\\d+)_", basename(filename))
  if (id_match > 0) {
    id_text <- regmatches(basename(filename), id_match)
    return(as.integer(sub("^(\\d+)_.*", "\\1", id_text)))
  }
  return(NA_integer_)
}

# Hardkodet mapping mellom fagskolenummer og rad i målfilen
get_target_rows <- function(fagskole_id) {
  # Hardkodet mapping basert på din liste
  row_mapping <- list(
    '197' = list(resultat = 7, balanse = 7),    # AOF Østlandet
    '226' = list(resultat = 8, balanse = 8),    # Campus BLÅ Fagskole AS
    '23' = list(resultat = 9, balanse = 9),     # Centric IT Academy
    '17' = list(resultat = 10, balanse = 10),   # Designinstituttet
    '89' = list(resultat = 11, balanse = 11),   # Det tverrfaglige kunstinstitutt
    '232' = list(resultat = 12, balanse = 12),  # Din Kompetanse fagskole
    '157' = list(resultat = 13, balanse = 13),  # Einar Granum Kunstfagskole
    '15' = list(resultat = 14, balanse = 14),   # Fabrikken Asker Kunstfagskole
    '236' = list(resultat = 15, balanse = 15),  # Fagskolen Diakonova
    '174' = list(resultat = 28, balanse = 28),  # Fagskolen Essens AS
    '85' = list(resultat = 16, balanse = 16),   # Fagskolen for bokbransjen
    '237' = list(resultat = 17, balanse = 17),  # Fagskolen GET Academy AS
    '40' = list(resultat = 18, balanse = 18),   # Fagskolen Kristiania
    '61' = list(resultat = 6, balanse = 6),     # Fagskolen Tirna
    '216' = list(resultat = 19, balanse = 19),  # Folkeuniversitetets Fagskole AS
    '91' = list(resultat = 20, balanse = 20),   # Frelsesarmeens offisersskole AS
    '220' = list(resultat = 21, balanse = 21),  # Gokstad akademiet
    '36' = list(resultat = 22, balanse = 22),   # Hald internasjonale skole
    '218' = list(resultat = 23, balanse = 23),  # KBT-fagskole
    '142' = list(resultat = 24, balanse = 24),  # Kunstfagskolen i Bergen
    '92' = list(resultat = 25, balanse = 25),   # Kunstskolen i Stavanger AS
    '111' = list(resultat = 26, balanse = 26),  # Lukas høyere yrkesfagskole
    '209' = list(resultat = 27, balanse = 27),  # MedLearn AS
    '104' = list(resultat = 29, balanse = 29),  # Norsk Hestesenter
    '29' = list(resultat = 30, balanse = 30),   # TISIP Fagskole
    '182' = list(resultat = 31, balanse = 31)   # Ytre kunstfagskole
  )
  
  # Konverterer til character for sikker sammenligning
  fagskole_id_str <- as.character(fagskole_id)
  
  if (fagskole_id_str %in% names(row_mapping)) {
    return(row_mapping[[fagskole_id_str]])
  } else {
    warning(sprintf("Fagskole-ID %s finnes ikke i mappingen. Ingen hardkodet rad.", fagskole_id))
    return(NULL)
  }
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

# ------- Hovedfunksjon - MED BEVART FORMATERING -------
excel_overforing_fagskole <- function(
  kilde_fil,
  output_fil,
  resultat_ark = "Tab_resultat",
  balanse_ark = "Tab_balanse",
  resultat_rad = NULL,
  balanse_rad = NULL
) {
  
  # Sjekk om denne filen skal behandles basert på fagskolenummeret
  if (!should_process_file(kilde_fil)) {
    cat("HOPPER OVER FIL:", basename(kilde_fil), "- ikke på listen over fagskoler som skal behandles\n")
    return(invisible(NULL))
  }
  
  # Få fagskole-ID fra filnavnet
  fagskole_id <- get_fagskole_id_from_filename(kilde_fil)
  cat("Fant fagskole-ID:", fagskole_id, "\n")
  
  # Få skolenavnet fra Resultatregnskap-arket
  inst_name <- get_school_name_from_source(kilde_fil, sheet_name = "Resultatregnskap")
  cat("Fant skolenavn:", inst_name, "\n")
  
  # Hent hardkodede rader for denne fagskole-ID
  target_rows <- get_target_rows(fagskole_id)
  
  if (is.null(target_rows)) {
    stop("Kunne ikke finne hardkodede rader for fagskolenummer ", fagskole_id, 
         ". Dette burde ikke skje da filen allerede er sjekket mot gyldig liste.")
  }
  
  # Bruk hardkodede rader
  if (is.null(resultat_rad)) {
    resultat_rad <- target_rows$resultat
    cat("Bruker hardkodet rad for resultat:", resultat_rad, "\n")
  }
  
  if (is.null(balanse_rad)) {
    balanse_rad <- target_rows$balanse
    cat("Bruker hardkodet rad for balanse:", balanse_rad, "\n")
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
  cat("2024: Driftsinnt:", driftsinntekter_2024, "Driftskost:", driftskostnader_2024, "Årsres:", arsresultat_2024, "\n")
  cat("2024: Omløp:", omlopsmidler_2024, "Egenk:", egenkapital_2024, "Kortgjeld:", korts_gjeld_2024, "Total:", totalkapital_2024, "\n")
  cat("2023: Driftsinnt:", driftsinntekter_2023, "Driftskost:", driftskostnader_2023, "Årsres:", arsresultat_2023, "\n")
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

  # Åpne EKSISTERENDE workbook for å bevare formatering
  cat("\n=== ÅPNER EKSISTERENDE WORKBOOK (BEVARER FORMATERING) ===\n")
  wb <- loadWorkbook(output_fil)
  
  # Funksjon for sikker skriving som bevarer formatering
  safe_write_preserve_format <- function(sheet, value, row, col, description) {
    if (!is.na(value)) {
      tryCatch({
        # Bruk writeData med kun verdien, ikke hele data.frame
        writeData(wb, sheet = sheet, x = value, startRow = row, startCol = col)
        cat("*** SKREV", description, ":", value, "til rad", row, "kolonne", col, "***\n")
      }, error = function(e) {
        cat("FEIL ved skriving av", description, ":", conditionMessage(e), "\n")
      })
    } else {
      cat("Hopper over", description, "- verdi er NA\n")
    }
  }

  cat("\n=== SKRIVER TIL RESULTAT-ARK (BEVARER FORMATERING) ===\n")
  # Følger mappingen: 2024 til kolonne D(4), 2023 til kolonne C(3)
  safe_write_preserve_format(resultat_ark, driftsinntekter_2024, resultat_rad, 4, "driftsinntekter 2024")
  safe_write_preserve_format(resultat_ark, driftsinntekter_2023, resultat_rad, 3, "driftsinntekter 2023")
  safe_write_preserve_format(resultat_ark, driftskostnader_2024, resultat_rad, 8, "driftskostnader 2024")
  safe_write_preserve_format(resultat_ark, driftskostnader_2023, resultat_rad, 7, "driftskostnader 2023")
  safe_write_preserve_format(resultat_ark, arsresultat_2024, resultat_rad, 16, "årsresultat 2024")
  safe_write_preserve_format(resultat_ark, arsresultat_2023, resultat_rad, 15, "årsresultat 2023")
  safe_write_preserve_format(resultat_ark, round(driftsmargin_2024, 1), resultat_rad, 20, "driftsmargin 2024")
  safe_write_preserve_format(resultat_ark, round(driftsmargin_2023, 1), resultat_rad, 19, "driftsmargin 2023")

  cat("\n=== SKRIVER TIL BALANSE-ARK (BEVARER FORMATERING) ===\n")
  # Følger mappingen: 2024 til kolonne D(4), 2023 til kolonne C(3)
  safe_write_preserve_format(balanse_ark, omlopsmidler_2024, balanse_rad, 4, "omløpsmidler 2024")
  safe_write_preserve_format(balanse_ark, omlopsmidler_2023, balanse_rad, 3, "omløpsmidler 2023")
  safe_write_preserve_format(balanse_ark, egenkapital_2024, balanse_rad, 8, "egenkapital 2024")
  safe_write_preserve_format(balanse_ark, egenkapital_2023, balanse_rad, 7, "egenkapital 2023")
  safe_write_preserve_format(balanse_ark, korts_gjeld_2024, balanse_rad, 12, "kortsiktig gjeld 2024")
  safe_write_preserve_format(balanse_ark, korts_gjeld_2023, balanse_rad, 11, "kortsiktig gjeld 2023")
  safe_write_preserve_format(balanse_ark, totalkapital_2024, balanse_rad, 16, "totalkapital 2024")
  safe_write_preserve_format(balanse_ark, totalkapital_2023, balanse_rad, 15, "totalkapital 2023")
  safe_write_preserve_format(balanse_ark, round(egenkapitalgrad_2024, 1), balanse_rad, 20, "egenkapitalgrad 2024")
  safe_write_preserve_format(balanse_ark, round(egenkapitalgrad_2023, 1), balanse_rad, 19, "egenkapitalgrad 2023")
  safe_write_preserve_format(balanse_ark, round(finansieringsgrad2_2024, 1), balanse_rad, 24, "finansieringsgrad2 2024")
  safe_write_preserve_format(balanse_ark, round(finansieringsgrad2_2023, 1), balanse_rad, 23, "finansieringsgrad2 2023")

  # Lagre workbook (bevarer original formatering)
  tryCatch({
    saveWorkbook(wb, output_fil, overwrite = TRUE)
    cat("\n*** WORKBOOK LAGRET SUCCESSFULLY (FORMATERING BEVART) ***\n")
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

# Eksempel på bruk for prosessering av flere filer
process_multiple_files <- function(kilde_filer, output_fil, resultat_ark = "Tab_resultat", balanse_ark = "Tab_balanse") {
  processed_count <- 0
  skipped_count <- 0
  
  for (kilde_fil in kilde_filer) {
    cat("\n\n=========================================================\n")
    cat("PROSESSERER FIL:", basename(kilde_fil), "\n")
    cat("=========================================================\n\n")
    
    result <- tryCatch({
      excel_overforing_fagskole(kilde_fil, output_fil, resultat_ark, balanse_ark)
      TRUE
    }, error = function(e) {
      cat("FEIL ved prosessering av", basename(kilde_fil), ":", conditionMessage(e), "\n")
      FALSE
    })
    
    if (isTRUE(result)) {
      processed_count <- processed_count + 1
    } else if (is.null(result)) {
      skipped_count <- skipped_count + 1
    }
  }
  
  cat("\n\n=========================================================\n")
  cat("PROSESSERING FULLFØRT\n")
  cat("Antall filer behandlet:", processed_count, "\n")
  cat("Antall filer hoppet over:", skipped_count, "\n")
  cat("Totalt antall filer:", length(kilde_filer), "\n")
  cat("=========================================================\n")
}