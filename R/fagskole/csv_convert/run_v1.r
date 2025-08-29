library(activeDir)
set_wd_to_current()

source("v1_convert.r")

behandle_regnskap_filer(
  mappe_sti = "//wsl.localhost/Ubuntu-24.04/home/manilpit/github/manilpit_github/excel_extract/data/fagskole",
  output_fil = "//wsl.localhost/Ubuntu-24.04/home/manilpit/github/manilpit_github/excel_extract/R/fagskole/mal/Kontroll_private_fagskoler_2025.xlsx",
  resultat_rad = 6,
  balanse_rad = 6
)