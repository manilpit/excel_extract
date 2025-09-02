# Prosesser alle kildefiler i en mappe:
process_folder <- function(kilde_mappe, output_fil,
  resultat_ark = "Tab_resultat",
  balanse_ark  = "Tab_balanse") {
xlsx <- list.files(kilde_mappe, pattern = "\\.xlsx$", full.names = TRUE)
for (f in xlsx) {
cat("Behandler:", basename(f), "\n")
try({
excel_overforing_fagskole(
kilde_fil   = f,
output_fil  = output_fil,
resultat_ark = resultat_ark,
balanse_ark  = balanse_ark
# Hvis du vil tvinge manuell rad (som i høgskole-scriptet), oppgi resultat_rad/balanse_rad her.
)
}, silent = TRUE)
}
}
