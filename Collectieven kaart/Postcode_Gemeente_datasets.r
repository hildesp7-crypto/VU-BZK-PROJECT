# ---- setup ----
# Packages: dplyr, tidyr, readxl, writexl, cbsodataR
# install.packages("cbsodataR")  # once
library(dplyr)
library(tidyr)
library(readxl)
library(writexl)
install.packages("cbsodataR")
library(cbsodataR)

rm(locaties)  # remove the old, too-small object from the session (error if absent is fine)

data_dir    <- "/Users/hildespaan/Documents/1 Werk & onderzoek/Werk/Vrije Universiteit Amsterdam/BZK Project/MAP BZK/Data "
source_file <- file.path(data_dir, "data_toevoegen.xlsx")
file.exists(source_file)  # should be TRUE

# ---- inlezen ----
locaties_raw <- read_excel(source_file, sheet = 1) |>
  mutate(Relatienaam = trimws(gsub("\\s+", " ", Relatienaam)),  # removes line breaks in names
         KvK = na_if(trimws(as.character(KvK)), ""))
nrow(locaties_raw)

# ---- dubbelingen bepalen ----
# same KvK number = same organisation; rows without KvK: same name (case-insensitive)
locaties_raw <- locaties_raw |>
  mutate(dedup_key = if_else(!is.na(KvK),
                             paste0("kvk_", KvK),
                             paste0("naam_", tolower(Relatienaam))))

sum(is.na(locaties_raw$KvK))  # rows deduplicated on name instead of KvK

# ---- uitdraai dubbelingen ----
duplicaten <- locaties_raw |>
  group_by(dedup_key) |>
  mutate(n_in_groep = n(),
         status = if_else(row_number() == 1, "behouden", "verwijderd")) |>
  ungroup() |>
  dplyr::filter(n_in_groep > 1) |>
  mutate(reden = if_else(startsWith(dedup_key, "kvk_"),
                         "zelfde KvK-nummer", "zelfde naam (geen KvK)")) |>
  arrange(dedup_key, desc(status == "behouden")) |>
  select(dedup_key, status, reden, n_in_groep, everything())

cat("Duplicate groups:", n_distinct(duplicaten$dedup_key),
    "| rows involved:", nrow(duplicaten),
    "| rows dropped:", sum(duplicaten$status == "verwijderd"), "\n")
count(duplicaten, reden)

write_xlsx(duplicaten, file.path(data_dir, "260922_duplicaten.xlsx"))

# ---- ontdubbelen ----
locaties <- locaties_raw |>
  distinct(dedup_key, .keep_all = TRUE) |>
  select(-dedup_key)

cat("Rows dropped as duplicates:", nrow(locaties_raw) - nrow(locaties), "\n")
cat("Rows remaining:", nrow(locaties), "\n")

# ---- gemeentenamen van CBS ----
# take the most recent "Gebieden in Nederland" table and read its region labels
toc <- cbs_get_toc(Language = "nl")
gebieden_id <- toc |>
  dplyr::filter(grepl("^Gebieden in Nederland", Title)) |>
  arrange(desc(Period)) |>
  slice(1) |>
  pull(Identifier)
gebieden_id  # check which year is used

gemeenten <- cbs_get_meta(gebieden_id)$RegioS |>
  transmute(gm_code = trimws(Key), gemeente = trimws(Title)) |>
  dplyr::filter(startsWith(gm_code, "GM"))

# ---- aantallen per gemeente ----
# rows with several codes ("GM0505, GM0523") count once in each gemeente
per_gemeente <- locaties |>
  separate_rows(gm_code, sep = ",\\s*") |>
  mutate(gm_code = na_if(trimws(gm_code), "")) |>
  count(gm_code, name = "n_organisaties") |>
  left_join(gemeenten, by = "gm_code") |>
  select(gemeente, gm_code, n_organisaties) |>
  arrange(desc(n_organisaties))

# checks: missing codes, and codes CBS doesn't recognise (e.g. merged gemeenten)
per_gemeente |> dplyr::filter(is.na(gm_code))
per_gemeente |> dplyr::filter(!is.na(gm_code), is.na(gemeente))

write_xlsx(per_gemeente, file.path(data_dir, "260922_aantal_per_gemeente.xlsx"))

# ---- aantallen per postcode ----
locaties <- locaties |>
  mutate(pc_clean = toupper(gsub("\\s", "", Postcode)),
         pc4 = if_else(grepl("^[0-9]{4}", pc_clean), substr(pc_clean, 1, 4), NA_character_),
         pc6 = if_else(grepl("^[0-9]{4}[A-Z]{2}$", pc_clean), pc_clean, NA_character_))

cat("With PC4:", sum(!is.na(locaties$pc4)),
    "| with full PC6:", sum(!is.na(locaties$pc6)),
    "| no postcode:", sum(is.na(locaties$pc4)), "\n")

per_pc4 <- locaties |>
  dplyr::filter(!is.na(pc4)) |>
  count(pc4, name = "n_organisaties") |>
  arrange(desc(n_organisaties))

per_pc6 <- locaties |>
  dplyr::filter(!is.na(pc6)) |>
  count(pc6, name = "n_organisaties") |>
  arrange(desc(n_organisaties))

write_xlsx(per_pc4, file.path(data_dir, "260922_aantal_per_pc4.xlsx"))
write_xlsx(per_pc6, file.path(data_dir, "260922_aantal_per_pc6.xlsx"))
