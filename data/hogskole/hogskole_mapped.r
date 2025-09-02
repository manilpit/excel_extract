# h_mapped.r - For bruk med source()

# Funksjon for å utføre Excel-overføring
excel_overføring <- function(
  kilde_fil, 
  output_fil, 
  resultat_rad, 
  balanse_rad, 
  institusjon = NULL,  # Nå er dette valgfritt, scriptet vil forsøke å finne det automatisk
  resultat_målark = "Tab resultat",
  eiendeler_målark = "balanse"
) {
  # Last inn nødvendige pakker
  suppressPackageStartupMessages(library(openxlsx))
  
  # 1. Sjekk at filer eksisterer
  if (!file.exists(kilde_fil)) {
    stop(paste("Kildefilen", kilde_fil, "finnes ikke!"), call.=FALSE)
  }
  
  if (!file.exists(output_fil)) {
    cat(paste("Målfilen", output_fil, "finnes ikke. Vil opprette ny fil.\n"))
  }
  
  # 2. Les inn kildefil og finn nøkkelverdier
  cat(paste("Leser inn kildefil:", kilde_fil, "\n"))
  
  # Hent liste over ark i kildefilen
  kilde_wb <- loadWorkbook(kilde_fil)
  tilgjengelige_ark <- names(kilde_wb)
  cat("Tilgjengelige ark i kildefilen:\n")
  for (i in seq_along(tilgjengelige_ark)) {
    cat(paste(" ", i, ":", tilgjengelige_ark[i], "\n"))
  }
  
  # Definer arknavnene i kildefilen
  resultat_ark <- "Resultatregnskap"
  eiendeler_ark <- "Balanse - eiendeler"
  gjeld_ek_ark <- "Balanse - gjeld og egenkapital"
  
  # Sjekk at nødvendige ark finnes i kildefilen
  if (!(resultat_ark %in% tilgjengelige_ark)) {
    stop(paste("Arknavnet", resultat_ark, "finnes ikke i kildefilen!"), call.=FALSE)
  }
  if (!(eiendeler_ark %in% tilgjengelige_ark)) {
    stop(paste("Arknavnet", eiendeler_ark, "finnes ikke i kildefilen!"), call.=FALSE)
  }
  if (!(gjeld_ek_ark %in% tilgjengelige_ark)) {
    stop(paste("Arknavnet", gjeld_ek_ark, "finnes ikke i kildefilen!"), call.=FALSE)
  }
  
  # Les inn data fra ark i kildefilen
  resultat_data <- read.xlsx(kilde_fil, sheet = resultat_ark)
  eiendeler_data <- read.xlsx(kilde_fil, sheet = eiendeler_ark)
  gjeld_ek_data <- read.xlsx(kilde_fil, sheet = gjeld_ek_ark)
  
  # Hjelpefunksjon for å finne verdi i en dataramme (hopper over lønnskostnader)
  finn_verdi <- function(data, nøkkelord) {
    # Hopp over lønnskostnader - de håndteres separat
    if (nøkkelord == "Lønnskostnader") {
      cat(" * Hopper over lønnskostnader i standard søk\n")
      return(NA)
    }
    
    # Standard søkemetode for andre nøkkelord
    for (rad in 1:nrow(data)) {
      for (kol in 1:ncol(data)) {
        if (!is.na(data[rad, kol]) && is.character(data[rad, kol])) {
          if (tolower(trimws(data[rad, kol])) == tolower(trimws(nøkkelord)) || 
              grepl(tolower(trimws(nøkkelord)), tolower(trimws(data[rad, kol])), fixed = TRUE)) {
            
            # Sjekk neste kolonner for en tallverdi
            for (offset in 1:5) {
              if (kol + offset <= ncol(data)) {
                verdi <- data[rad, kol + offset]
                if (is.numeric(verdi) || (!is.na(verdi) && is.character(verdi) && 
                    grepl("^\\s*-?[0-9.,]+\\s*$", verdi))) {
                  
                  # Konverter til tall
                  if (is.character(verdi)) {
                    verdi <- as.numeric(gsub("[^0-9\\.-]", "", gsub(",", ".", verdi)))
                  }
                  
                  cat(paste(" * Fant", nøkkelord, "=", verdi, "\n"))
                  return(verdi)
                }
              }
            }
          }
        }
      }
    }
    cat(paste(" * Fant ikke verdi for", nøkkelord, "\n"))
    return(NA)
  }
  
  # Spesialfunksjon for å finne lønnskostnader direkte fra kolonne 3
  finn_lonnskostnader <- function(data) {
    for (rad in 1:nrow(data)) {
      for (kol in 1:ncol(data)) {
        if (!is.na(data[rad, kol]) && is.character(data[rad, kol])) {
          if (grepl("lønnskostnad", tolower(data[rad, kol]), fixed = TRUE)) {
            cat(paste("Fant lønnskostnader i rad", rad, "kolonne", kol, "\n"))
            
            # Direkte hent fra kolonne 3 som vi vet har lønnskostnader
            if (3 <= ncol(data)) {
              lonnsverdi <- data[rad, 3]  # Direkte fra kolonne 3
              
              # Sjekk om dette er et tall
              if (is.numeric(lonnsverdi) || (!is.na(lonnsverdi) && is.character(lonnsverdi) && 
                  grepl("^\\s*-?[0-9.,]+\\s*$", lonnsverdi))) {
                
                # Konverter til tall
                if (is.character(lonnsverdi)) {
                  lonnsverdi <- as.numeric(gsub("[^0-9\\.-]", "", gsub(",", ".", lonnsverdi)))
                }
                
                cat(paste(" * Fant Lønnskostnader =", lonnsverdi, "(fra kolonne 3)\n"))
                return(lonnsverdi)
              } else {
                cat(paste(" * Verdi i kolonne 3 ikke numerisk:", lonnsverdi, "\n"))
              }
            }
          }
        }
      }
    }
    
    cat(" * Kunne ikke finne lønnskostnader i kolonne 3\n")
    return(NA)
  }
  
  # Hjelpefunksjon for å finne tekst i en dataramme
  finn_tekst <- function(data, søkeord) {
    for (rad in 1:nrow(data)) {
      for (kol in 1:ncol(data)) {
        if (!is.na(data[rad, kol]) && is.character(data[rad, kol])) {
          if (grepl(tolower(søkeord), tolower(data[rad, kol]), fixed = TRUE)) {
            return(trimws(data[rad, kol]))
          }
        }
      }
    }
    return(NULL)
  }
  
  # Finn institusjonsnavn hvis ikke spesifisert
  if (is.null(institusjon)) {
    # Prøv å finne institusjonsnavn fra resultatregnskapet (vanligvis i overskriften)
    institusjon_kandidater <- c()
    
    # Let i resultatarket
    for (rad in 1:5) {  # Sjekk de første radene
      if (rad <= nrow(resultat_data)) {
        for (kol in 1:ncol(resultat_data)) {
          if (!is.na(resultat_data[rad, kol]) && is.character(resultat_data[rad, kol])) {
            tekst <- trimws(resultat_data[rad, kol])
            if (nchar(tekst) > 3 && !grepl("^[0-9]", tekst) && 
                !grepl("^sum", tolower(tekst)) && 
                !grepl("^resultat", tolower(tekst))) {
              institusjon_kandidater <- c(institusjon_kandidater, tekst)
            }
          }
        }
      }
    }
    
    # Søk etter spesifikke mønstre i filnavnet
    if (length(institusjon_kandidater) == 0) {
      # Prøv å hente fra filnavnet
      fil_base <- basename(kilde_fil)
      if (grepl("_", fil_base)) {
        deler <- strsplit(fil_base, "_")[[1]]
        if (length(deler) >= 3) {
          # Vanlig format: 8208_2024_VID.xlsx
          mulig_institutt <- gsub("\\.xlsx$", "", deler[3])
          institusjon_kandidater <- c(institusjon_kandidater, mulig_institutt)
        }
      }
    }
    
    # Mapper fra kortnavn til fullt navn
    institusjon_mapping <- list(
      "VID" = "VID vitenskapelige høgskole",
      "NLA" = "NLA Høgskolen",
      "MF" = "MF vitenskapelig høyskole",
      "FIH" = "Fjellhaug Internasjonale Høgskole",
      "ATH" = "Ansgar Teologiske Høgskole",
      "HLT" = "Høyskolen for Ledelse og Teologi",
      "LDH" = "Lovisenberg diakonale høgskole",
      "DMMH" = "Dronning Mauds Minne Høgskole",
      "BDH" = "Barratt Due musikkinstitutt, høyskoleavdelingen",
      "BIMM" = "Bergen Internasjonale Filmskole"
    )
    
    # Velg det beste kandidatnavnet
    if (length(institusjon_kandidater) > 0) {
      institusjon_kort <- institusjon_kandidater[1]
      
      # Sjekk om vi kan mappe kortnavn til fullt navn
      if (institusjon_kort %in% names(institusjon_mapping)) {
        institusjon <- institusjon_mapping[[institusjon_kort]]
      } else {
        institusjon <- institusjon_kort
      }
    } else {
      # Fallback til standard
      institusjon <- "Ukjent høgskole/institusjon"
    }
    
    cat(paste(" * Fant institusjon:", institusjon, "\n"))
  }
  
  # Definer standard målark og kolonner
  gjeld_ek_målark <- eiendeler_målark  # Samme som eiendeler_målark
  resultat_kolonner <- c(4, 8, 12, 16, 20)  # Kolonner for Sum driftsinntekter, Lønnskostnader, etc.
  eiendeler_kolonner <- c(4, 8, 12)        # Kolonner for anleggsmidler, omløpsmidler, sum eiendeler
  gjeld_ek_kolonner <- c(16, 20, 24, 28, 32, 36)  # Kolonner for gjeld/EK verdier
  
  # Vis konfigurasjon
  cat("\n== EXCEL OVERFØRING ==\n")
  cat(paste("Institusjon:", institusjon, "\n"))
  cat(paste("Kildefil:", kilde_fil, "\n"))
  cat(paste("Målfil:", output_fil, "\n"))
  cat(paste("Resultatark:", resultat_målark, "på rad", resultat_rad, "\n"))
  cat(paste("Balanseark:", eiendeler_målark, "på rad", balanse_rad, "\n"))
  cat("=====================\n\n")
  
  # Finn verdi for anleggsmidler ved å summere komponenter
  finn_anleggsmidler <- function() {
    # Prøv direkte søk først
    direkte <- finn_verdi(eiendeler_data, "Sum anleggsmidler")
    if (!is.na(direkte)) return(direkte)
    
    # Beregn fra komponenter
    cat("Beregner anleggsmidler fra komponenter...\n")
    sum_immaterielle <- finn_verdi(eiendeler_data, "Sum immaterielle eiendeler")
    sum_varige <- finn_verdi(eiendeler_data, "Sum varige driftsmidler")
    sum_finansielle <- finn_verdi(eiendeler_data, "Sum finansielle anleggsmidler")
    
    # Utskrift for debugging
    cat(paste(" * Sum immaterielle eiendeler =", sum_immaterielle, "\n"))
    cat(paste(" * Sum varige driftsmidler =", sum_varige, "\n"))
    cat(paste(" * Sum finansielle anleggsmidler =", sum_finansielle, "\n"))
    
    # Samle komponenter og fjerne NA-verdier
    komponenter <- c(sum_immaterielle, sum_varige, sum_finansielle)
    komponenter <- komponenter[!is.na(komponenter)]
    
    if (length(komponenter) > 0) {
      anlegg <- sum(komponenter)
      cat(paste(" * Beregnet anleggsmidler =", anlegg, "\n"))
      return(anlegg)
    }
    
    # Prøv differanse mellom sum eiendeler og omløpsmidler
    sum_eiendeler <- finn_verdi(eiendeler_data, "SUM EIENDELER")
    omlop <- finn_omlopsmidler()
    
    if (!is.na(sum_eiendeler) && !is.na(omlop)) {
      anlegg <- sum_eiendeler - omlop
      cat(paste(" * Beregnet anleggsmidler som differanse =", anlegg, "\n"))
      return(anlegg)
    }
    
    return(NA)
  }
  
  # Finn verdi for omløpsmidler (spesialtilfelle)
  finn_omlopsmidler <- function() {
    # Prøv direkte søk først
    direkte <- finn_verdi(eiendeler_data, "Omløpsmidler")
    if (!is.na(direkte)) return(direkte)
    
    direkte <- finn_verdi(eiendeler_data, "Sum omløpsmidler")
    if (!is.na(direkte)) return(direkte)
    
    # Prøv å regne ut fra komponenter
    cat("Beregner omløpsmidler fra komponenter...\n")
    sum_varer <- finn_verdi(eiendeler_data, "Sum varer")
    sum_fordringer <- finn_verdi(eiendeler_data, "Sum fordringer")
    sum_investeringer <- finn_verdi(eiendeler_data, "Sum investeringer")
    sum_bank <- finn_verdi(eiendeler_data, "Sum bankinnskudd, kontanter og lignende")
    
    # Utskrift for debugging
    cat(paste(" * Sum varer =", sum_varer, "\n"))
    cat(paste(" * Sum fordringer =", sum_fordringer, "\n"))
    cat(paste(" * Sum investeringer =", sum_investeringer, "\n"))
    cat(paste(" * Sum bankinnskudd =", sum_bank, "\n"))
    
    komponenter <- c(sum_varer, sum_fordringer, sum_investeringer, sum_bank)
    komponenter <- komponenter[!is.na(komponenter)]
    
    if (length(komponenter) > 0) {
      omlop <- sum(komponenter)
      cat(paste(" * Beregnet omløpsmidler =", omlop, "\n"))
      return(omlop)
    }
    
    # Prøv differanse mellom sum eiendeler og anleggsmidler
    sum_eiendeler <- finn_verdi(eiendeler_data, "SUM EIENDELER")
    direkte_anlegg <- finn_verdi(eiendeler_data, "Sum anleggsmidler")
    
    if (!is.na(sum_eiendeler) && !is.na(direkte_anlegg)) {
      omlop <- sum_eiendeler - direkte_anlegg
      cat(paste(" * Beregnet omløpsmidler som differanse =", omlop, "\n"))
      return(omlop)
    }
    
    return(NA)
  }
  
  # 3. Finn verdier fra kildefil
  cat("Henter verdier fra kildefil...\n")
  
  # For resultatregnskapet - vi hopper over lønnskostnader for nå
  driftsinntekter <- finn_verdi(resultat_data, "Sum driftsinntekter")
  # Lønnskostnader håndteres separat
  driftskostnader <- finn_verdi(resultat_data, "Sum driftskostnader")
  driftsresultat <- finn_verdi(resultat_data, "Driftsresultat")
  arsresultat <- finn_verdi(resultat_data, "Årsresultat")
  
  # For eiendeler
  anleggsmidler <- finn_anleggsmidler()  # Nå med summeringslogikk
  omlopsmidler <- finn_omlopsmidler()
  sum_eiendeler <- finn_verdi(eiendeler_data, "SUM EIENDELER")
  
  # For gjeld og egenkapital
  opptjent_ek <- finn_verdi(gjeld_ek_data, "Sum opptjent egenkapital")
  sum_ek <- finn_verdi(gjeld_ek_data, "Sum egenkapital")
  avsetning_forpliktelser <- finn_verdi(gjeld_ek_data, "Sum avsetning for forpliktelser")
  langsiktig_gjeld <- finn_verdi(gjeld_ek_data, "Sum annen langsiktig gjeld")
  kortsiktig_gjeld <- finn_verdi(gjeld_ek_data, "Sum kortsiktig gjeld")
  sum_gjeld_ek <- finn_verdi(gjeld_ek_data, "SUM EGENKAPITAL OG GJELD")
  
  # 4. Last inn målfilen
  cat(paste("\nLaster målfil:", output_fil, "\n"))
  
  # Sjekk om målfil eksisterer, ellers opprett ny
  if (file.exists(output_fil)) {
    # Prøv å laste inn eksisterende fil med feilhåndtering
    tryCatch({
      wb <- loadWorkbook(output_fil)
      
      # List opp alle ark i målfilen for å verifisere navn
      målark <- names(wb)
      cat("Ark i målfilen:\n")
      for (i in seq_along(målark)) {
        cat(paste(" ", i, ":", målark[i], "\n"))
      }
      
      # Sjekk om arkene finnes
      if (!(resultat_målark %in% målark)) {
        stop(paste("FEIL: Arket", resultat_målark, "finnes ikke i målfilen!"), call.=FALSE)
      }
      
      if (!(eiendeler_målark %in% målark)) {
        lignende_ark <- grep("balanse", målark, ignore.case = TRUE, value = TRUE)
        if (length(lignende_ark) > 0) {
          cat(paste("Fant lignende balanseark:", paste(lignende_ark, collapse = ", "), "\n"))
          cat("Bruk eiendeler_målark parameter for å spesifisere riktig arknavn.\n")
        }
        stop(paste("FEIL: Arket", eiendeler_målark, "finnes ikke i målfilen!"), call.=FALSE)
      }
    }, error = function(e) {
      cat(paste("FEIL ved lesing av målfil:", e$message, "\n"))
      cat("Oppretter ny målfil...\n")
      wb <- createWorkbook()
      addWorksheet(wb, resultat_målark)
      addWorksheet(wb, eiendeler_målark)
    })
  } else {
    # Opprett ny arbeidsbok
    wb <- createWorkbook()
    addWorksheet(wb, resultat_målark)
    addWorksheet(wb, eiendeler_målark)
    cat("Opprettet ny målfil med ark:", resultat_målark, "og", eiendeler_målark, "\n")
  }
  
  # 5. Skriv resultatverdier til målfilen (uten lønnskostnader)
  cat("\n-- SKRIVER VERDIER TIL MÅLFIL (UNNTATT LØNNSKOSTNADER) --\n")
  cat(paste("Skriver til", resultat_målark, "på rad", resultat_rad, "...\n"))
  
  resultat_verdier <- c(driftsinntekter, NA, driftskostnader, driftsresultat, arsresultat)
  resultat_nøkkelord <- c("Sum driftsinntekter", "Lønnskostnader", "Sum driftskostnader", 
                        "Driftsresultat", "Årsresultat")
  verdier_skrevet <- 0
  
  for (i in 1:length(resultat_verdier)) {
    # Ikke skriv NA-verdier
    if (!is.na(resultat_verdier[i])) {
      # Skriv verdien direkte til cellen
      writeData(wb, sheet = resultat_målark, x = resultat_verdier[i], 
               startRow = resultat_rad, startCol = resultat_kolonner[i])
      cat(paste(" * Skrev", resultat_nøkkelord[i], "=", resultat_verdier[i], "til", 
               resultat_målark, "rad", resultat_rad, "kolonne", resultat_kolonner[i], "\n"))
      verdier_skrevet <- verdier_skrevet + 1
    }
  }
  
  # 6. Skriv balanseverdier til målfilen
  cat(paste("\nSkriver til", eiendeler_målark, "på rad", balanse_rad, "...\n"))
  
  # Skriv eiendeler-verdier
  eiendeler_verdier <- c(anleggsmidler, omlopsmidler, sum_eiendeler)
  eiendeler_nøkkelord <- c("Sum anleggsmidler", "Sum omløpsmidler", "SUM EIENDELER")
  
  for (i in 1:length(eiendeler_verdier)) {
    # Ikke skriv NA-verdier
    if (!is.na(eiendeler_verdier[i])) {
      # Skriv verdien direkte til cellen med feilhåndtering
      tryCatch({
        writeData(wb, sheet = eiendeler_målark, x = eiendeler_verdier[i], 
                 startRow = balanse_rad, startCol = eiendeler_kolonner[i])
        cat(paste(" * Skrev", eiendeler_nøkkelord[i], "=", eiendeler_verdier[i], "til", 
                 eiendeler_målark, "rad", balanse_rad, "kolonne", eiendeler_kolonner[i], "\n"))
        verdier_skrevet <- verdier_skrevet + 1
      }, error = function(e) {
        cat(paste("FEIL ved skriving av", eiendeler_nøkkelord[i], "til balanseark:", e$message, "\n"))
      })
    }
  }
  
  # Skriv gjeld/EK-verdier
  gjeld_ek_verdier <- c(opptjent_ek, sum_ek, avsetning_forpliktelser, 
                       langsiktig_gjeld, kortsiktig_gjeld, sum_gjeld_ek)
  gjeld_ek_nøkkelord <- c("Sum opptjent egenkapital", "Sum egenkapital", 
                        "Sum avsetning for forpliktelser", "Sum annen langsiktig gjeld", 
                        "Sum kortsiktig gjeld", "SUM EGENKAPITAL OG GJELD")
  
  for (i in 1:length(gjeld_ek_verdier)) {
    # Ikke skriv NA-verdier
    if (!is.na(gjeld_ek_verdier[i])) {
      # Skriv verdien direkte til cellen med feilhåndtering
      tryCatch({
        writeData(wb, sheet = gjeld_ek_målark, x = gjeld_ek_verdier[i], 
                 startRow = balanse_rad, startCol = gjeld_ek_kolonner[i])
        cat(paste(" * Skrev", gjeld_ek_nøkkelord[i], "=", gjeld_ek_verdier[i], "til", 
                 gjeld_ek_målark, "rad", balanse_rad, "kolonne", gjeld_ek_kolonner[i], "\n"))
        verdier_skrevet <- verdier_skrevet + 1
      }, error = function(e) {
        cat(paste("FEIL ved skriving av", gjeld_ek_nøkkelord[i], "til balanseark:", e$message, "\n"))
      })
    }
  }
  
  # TRINN 2: SEPARAT BEHANDLING AV LØNNSKOSTNADER
  cat("\n== SEPARAT BEHANDLING AV LØNNSKOSTNADER ==\n")
  
  # Finn lønnskostnader direkte fra kolonne 3
  lonnskostnader <- finn_lonnskostnader(resultat_data)
  
  if (!is.na(lonnskostnader)) {
    # Skriv lønnskostnader til målfilen
    writeData(wb, sheet = resultat_målark, x = lonnskostnader, 
             startRow = resultat_rad, startCol = resultat_kolonner[2])  # Kolonne 2 er for lønnskostnader
    cat(paste(" * Skrev Lønnskostnader =", lonnskostnader, "til", 
             resultat_målark, "rad", resultat_rad, "kolonne", resultat_kolonner[2], "\n"))
    verdier_skrevet <- verdier_skrevet + 1
  } else {
    cat(" * ADVARSEL: Kunne ikke finne lønnskostnader!\n")
  }
  
  # 7. Lagre målfilen
  tryCatch({
    saveWorkbook(wb, output_fil, overwrite = TRUE)
    cat(paste("\nLagret", verdier_skrevet, "verdier til filen:", output_fil, "\n"))
  }, error = function(e) {
    cat(paste("FEIL ved lagring av Excel-filen:", e$message, "\n"))
  })
  
  # 8. Skriv sammendrag
  cat("\n==== SAMMENDRAG ====\n")
  cat(paste("Institusjon:", institusjon, "\n"))
  cat(paste("Verdier funnet og skrevet:", verdier_skrevet, "\n"))
  cat(paste("Kildefil:", kilde_fil, "\n"))
  cat(paste("Målfil:", output_fil, "\n"))
  cat("=====================\n")
  
  return(invisible(verdier_skrevet))
}

# # Eksempel på bruk (kan fjernes/kommenteres ut)
# excel_overføring(
#   kilde_fil = "8202_2024_høyskolen_197978.xlsx", 
#   output_fil = "bruksanvisning_1.xlsx", 
#   resultat_rad = 12, 
#   balanse_rad = 12,
#   resultat_målark = "Tab resultat",
#   eiendeler_målark = "balanse"
# )