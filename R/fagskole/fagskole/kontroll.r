
#Final version_v4:
# kjor_script.R
# Dette er kontrollskriptet som setter stier og kjører hovedfunksjonen
library(activeDir)
set_wd_to_current()

# Last inn hovedfunksjoner
source("v6.r")

# Sett stier
kilde_mappe <- "//wsl.localhost/Ubuntu-24.04/home/manilpit/github/manilpit_github/excel_extract/data/fagskole"  # Endre til din mappe med kildefiler
output_fil <- "//wsl.localhost/Ubuntu-24.04/home/manilpit/github/manilpit_github/excel_extract/resultat/Kontroll_private_fagskoler_2025.xlsx"  # Endre til din målfil

# Funksjon for å kjøre på alle filer
kjor_for_alle_filer <- function(kilde_mappe, output_fil) {
  # Hent alle Excel-filer i mappen
  filer <- list.files(kilde_mappe, pattern = "\\.xlsx$|\\.xls$", full.names = TRUE)
  
  if (length(filer) == 0) {
    stop("Ingen Excel-filer funnet i mappen:", kilde_mappe)
  }
  
  # Vis antall filer som skal behandles
  cat("Behandler", length(filer), "filer fra", kilde_mappe, "til", output_fil, "\n")
  
  # Løp gjennom hver fil og kjør overføringsfunksjonen
  for (fil in filer) {
    tryCatch({
      cat("Behandler fil:", basename(fil), "... ")
      excel_overforing_fagskole(
        kilde_fil = fil,
        output_fil = output_fil,
        resultat_ark = "Tab_resultat",
        balanse_ark = "Tab_balanse"
      )
    }, error = function(e) {
      cat("FEIL:", conditionMessage(e), "\n")
    })
  }
  
  cat("\nFerdig! Alle filer er behandlet.\n")
}

# Kjør prosesseringen
kjor_for_alle_filer(kilde_mappe, output_fil)
