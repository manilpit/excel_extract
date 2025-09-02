
source("v5.r")
arknavnene <- behandle_regnskap_filer(
  mappe_sti = "//wsl.localhost/Ubuntu-24.04/home/manilpit/github/manilpit_github/excel_extract/data/fagskole",
  output_fil = "//wsl.localhost/Ubuntu-24.04/home/manilpit/github/manilpit_github/excel_extract/R/fagskole/mal/Kontroll_private_fagskoler_2025.xlsx",
  resultat_rad = 6,
  balanse_rad = 6
)
