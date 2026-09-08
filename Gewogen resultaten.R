knitr::opts_chunk$set(echo = TRUE, message = FALSE, warning = FALSE)

library(here)
library(readxl)
library(dplyr)
library(tidyr)
library(purrr)
library(stringr)
library(janitor)
library(survey)
library(srvyr)
library(knitr)

pad_steekproef <- here("/Users/hildespaan/Documents/6 Overig/Steekproef_herhaling.xlsx")

ruwe_data <- read_excel(pad_steekproef, sheet = "Volledige steekproef") |>
  clean_names() |>
  filter(!is.na(koepel_organisatie))  # trailing empty rows in the sheet

nrow(ruwe_data)

# The second sheet should contain exactly the rows flagged as burgerinitiatief.
controle_bc <- read_excel(pad_steekproef, sheet = "Burgercollectieven") |>
  clean_names() |>
  filter(!is.na(koepel_organisatie))

vlag_namen <- ruwe_data |> filter(burgerinitiatief == "Ja") |> pull(naam_organisatie)

c(
  n_vlag_in_volledige = length(vlag_namen),
  n_tabblad_bc        = nrow(controle_bc),
  verschil            = length(setdiff(controle_bc$naam_organisatie, vlag_namen)) +
                        length(setdiff(vlag_namen, controle_bc$naam_organisatie))
)

# Map the koepel labels as they appear in the workbook onto the eight strata of Table B2.1.
stratum_lookup <- tribble(
  ~koepel_organisatie,        ~stratum,
  "Cooplink",                 "Cooplink",
  "LSA",                      "LSA Bewoners",
  "MAEX",                     "MAEX",
  "NL Zorgt voor elkaar",     "NLZVE",
  "NCR",                      "NCR",
  "Collectieve Kracht",       "CollectieveKracht",
  "Groninger Dorpen",         "LVKK",
  "BOKD",                     "LVKK",
  "DKK Gelderland",           "LVKK",
  "Dorpswerk Noord-Holland",  "LVKK",
  "ZHVKK",                    "LVKK",
  "VKKNB",                    "LVKK",
  "LEM1",                     "Energie Samen",
  "LEM2",                     "Energie Samen"
)

# Sampling frame sizes from Table B2.1 (total frame 14.660). Replace for the 2026 frame.
populatie <- tribble(
  ~stratum,             ~n_populatie,
  "Cooplink",                    149,
  "LSA Bewoners",               7800,
  "MAEX",                       2404,
  "NLZVE",                      1469,
  "NCR",                         243,
  "CollectieveKracht",            73,
  "LVKK",                       1817,
  "Energie Samen",               705
)

ruwe_data <- ruwe_data |>
  mutate(koepel_organisatie = str_squish(koepel_organisatie)) |>
  left_join(stratum_lookup, by = "koepel_organisatie")

# Any label that did not match is a spelling variant that must be added to the lookup.
ruwe_data |> filter(is.na(stratum)) |> count(koepel_organisatie)

gewichten <- ruwe_data |>
  count(stratum, name = "n_steekproef") |>
  left_join(populatie, by = "stratum") |>
  mutate(
    gewicht = n_populatie / n_steekproef,
    # Normalised weight averages 1 over the sample: use when you want proportions without
    # population-sized standard errors.
    gewicht_genormaliseerd = gewicht / (sum(n_populatie) / sum(n_steekproef))
  )

gewichten |> kable(digits = 1)

ruwe_data <- ruwe_data |>
  left_join(
    gewichten |> select(stratum, n_populatie, n_steekproef, gewicht, gewicht_genormaliseerd),
    by = "stratum"
  )

# Sanity check: the weighted N must reproduce the frame size.
c(
  ongewogen_n = nrow(ruwe_data),
  gewogen_n   = sum(ruwe_data$gewicht),
  kader_n     = sum(populatie$n_populatie)
)

w <- ruwe_data$gewicht
# Kish's effective sample size: a blunt but honest summary of the precision loss caused by
# the very unequal weights (LSA about 430, CollectieveKracht about 8).
c(
  deff_kish   = length(w) * sum(w^2) / sum(w)^2,
  n_effectief = sum(w)^2 / sum(w^2)
)

# The website column holds either a URL or "Onbekend/Onbekend"; one entry lacks the http
# prefix, so match www. as well. Social media columns mix a bare "Nee" with
# "Ja (actief augustus 2026)" and "Ja (niet actief)": separate presence from activity
# instead of letting every month-label become its own category.
data_hercodeerd <- ruwe_data |>
  mutate(
    heeft_website = if_else(
      str_detect(str_to_lower(str_squish(website)), "^(https?://|www\\.)"),
      "Ja", "Nee"
    ),
    across(
      c(linkedin_pagina, instagram_pagina, facebook_pagina),
      ~ case_when(
        str_starts(str_to_lower(str_squish(.x)), "ja")  ~ "Ja",
        str_starts(str_to_lower(str_squish(.x)), "nee") ~ "Nee",
        TRUE                                            ~ "Onbekend"
      ),
      .names = "heeft_{.col}"
    ),
    across(
      c(instagram_pagina, facebook_pagina),
      ~ case_when(
        str_detect(str_to_lower(.x), "niet actief") ~ "Ja, inactief",
        str_starts(str_to_lower(.x), "ja")          ~ "Ja, actief",
        str_starts(str_to_lower(.x), "nee")         ~ "Nee",
        TRUE                                        ~ "Onbekend"
      ),
      .names = "status_{.col}"
    )
  ) |>
  rename(
    heeft_linkedin  = heeft_linkedin_pagina,
    heeft_instagram = heeft_instagram_pagina,
    heeft_facebook  = heeft_facebook_pagina
  ) |>
  mutate(
    online_aanwezig = if_else(
      heeft_website == "Ja" | heeft_linkedin == "Ja" |
        heeft_instagram == "Ja" | heeft_facebook == "Ja",
      "Ja", "Nee"
    )
)

# `actief` and `werkgebied` carry free-text elaborations that split categories meaning the
# same thing: "Landelijk" vs "Landelijk (Nederland)", four flavours of "Regionaal",
# "Opgeheven (gestopt in 2026)" vs "Opgeheven".
data_hercodeerd <- data_hercodeerd |>
  mutate(
    actief_hercodeerd = case_when(
      str_starts(actief, "Actief")      ~ "Actief",
      str_starts(actief, "Opgeheven")   ~ "Opgeheven",
      str_starts(actief, "Inactief")    ~ "Opgeheven",
      str_starts(actief, "Onduidelijk") ~ "Onduidelijk",
      TRUE                              ~ "Onbekend"
    ),
    werkgebied_hercodeerd = case_when(
      str_starts(werkgebied, "Lokaal")     ~ "Lokaal",
      str_starts(werkgebied, "Regionaal")  ~ "Regionaal",
      str_starts(werkgebied, "Landelijk")  ~ "Landelijk",
      str_detect(werkgebied, "nternation") ~ "Internationaal",
      TRUE                                 ~ "Onbekend"
    ),
    # Treat explicit "Onbekend" as missing in the criterion variables, so weighted
    # percentages are of the cases where the criterion could actually be assessed.
    across(
      starts_with("criterium_"),
      ~ na_if(str_squish(.x), "Onbekend"),
      .names = "{.col}_bekend"
    )
  )

# Check every recode against its source column before trusting it.
data_hercodeerd |> count(heeft_website, heeft_url = str_starts(website, "http"))
data_hercodeerd |> count(actief, actief_hercodeerd) |> arrange(actief_hercodeerd)
data_hercodeerd |> count(werkgebied, werkgebied_hercodeerd) |> arrange(werkgebied_hercodeerd)
data_hercodeerd |> count(facebook_pagina, heeft_facebook, status_facebook_pagina) |> head(10)

# Stratified sample without replacement; fpc uses the frame size per stratum.
design_alle <- data_hercodeerd |>
  as_survey_design(ids = 1, strata = stratum, weights = gewicht, fpc = n_populatie)

# The subset keeps the weights of the full sample: recomputing N/n here would erase the
# selection on "is a burgercollectief", which is precisely the quantity being estimated.
# Use subset() rather than filtering the data first, so the design information survives.
design_bc <- design_alle |> filter(burgerinitiatief == "Ja")

c(
  n_alle        = nrow(data_hercodeerd),
  n_bc          = sum(data_hercodeerd$burgerinitiatief == "Ja"),
  gewogen_alle  = sum(data_hercodeerd$gewicht),
  gewogen_bc    = sum(data_hercodeerd$gewicht[data_hercodeerd$burgerinitiatief == "Ja"])
)

# Logit CIs rather than the normal approximation: proportions near zero otherwise get
# impossible negative lower bounds.
gewogen_tabel <- function(design, var) {
  ruwe_n <- design$variables |>
    count(categorie = as.character(.data[[var]]), name = "n_ongewogen")

  design |>
    mutate(cat_ = as.character(.data[[var]])) |>
    group_by(cat_) |>
    summarise(
      aandeel = survey_prop(vartype = "ci", proportion = TRUE, na.rm = TRUE),
      aantal  = survey_total(vartype = "ci", na.rm = TRUE)
    ) |>
    rename(categorie = cat_) |>
    left_join(ruwe_n, by = "categorie") |>
    transmute(
      variabele = var,
      categorie,
      n_ongewogen,
      pct       = round(100 * aandeel, 1),
      pct_ci    = sprintf("%.1f–%.1f", 100 * aandeel_low, 100 * aandeel_upp),
      aantal    = round(aantal),
      aantal_ci = sprintf("%.0f–%.0f", aantal_low, aantal_upp)
    ) |>
    arrange(desc(pct))
}

# Free-text fields, identifiers and the raw versions of recoded columns are not
# distributions worth tabulating.
uitsluiten <- c(
  "naam_organisatie", "naam_zoals_in_kvk_register", "kvk_beschrijving", "adres", "plaats",
  "website", "link_naar_jaarverslag", "missie_doelstelling_activiteiten_website",
  "rsin_nummer", "btw_nummer", "kvk_nummer_2026", "oprichting", "anbi_sinds",
  "gewicht", "gewicht_genormaliseerd", "stratum", "koepel_organisatie", "steekproef",
  "doelgroep", "actief", "werkgebied",
  "linkedin_pagina", "instagram_pagina", "facebook_pagina"
)

kandidaten <- data_hercodeerd |>
  select(where(~ is.character(.x) || is.factor(.x))) |>
  select(-any_of(uitsluiten)) |>
  names()

# Anything with more than 15 distinct values is free text, not a category.
kandidaten <- kandidaten[map_int(kandidaten, ~ n_distinct(data_hercodeerd[[.x]])) <= 15]
kandidaten

tabellen_alle <- map(kandidaten, ~ gewogen_tabel(design_alle, .x)) |>
  bind_rows() |>
  rename_with(~ paste0(.x, "_alle"), c(n_ongewogen, pct, pct_ci, aantal, aantal_ci))

# burgerinitiatief is constant in the subset, so it has no distribution there.
kandidaten_bc <- setdiff(kandidaten, "burgerinitiatief")

tabellen_bc <- map(kandidaten_bc, ~ gewogen_tabel(design_bc, .x)) |>
  bind_rows() |>
  rename_with(~ paste0(.x, "_bc"), c(n_ongewogen, pct, pct_ci, aantal, aantal_ci))

vergelijking <- tabellen_alle |>
  full_join(tabellen_bc, by = c("variabele", "categorie")) |>
  arrange(variabele, desc(pct_alle))

vergelijking

# Keep `variabele` in every printed table: with results='asis' the cat() heading only shows
# up in the knitted document, so running this chunk in the console would otherwise give a
# stack of tables with no way to tell which variable each one belongs to.
vergelijking |>
  group_split(variabele) |>
  walk(function(tab) {
    cat("\n\n## ", tab$variabele[1], "\n\n")
    print(kable(tab, caption = paste("Variabele:", tab$variabele[1])))
  })

# Named list, so a single variable can be pulled up interactively without hunting through
# the full comparison table.
tabellen <- vergelijking |>
  group_split(variabele) |>
  set_names(map_chr(vergelijking |> group_split(variabele), ~ .x$variabele[1]))

names(tabellen)

# Example: one variable at a time.
tabellen[["heeft_website"]]
tabellen[["werkgebied_hercodeerd"]]

# Doelgroep is a multiple-response field (semicolon-separated), so it is not a distribution:
# percentages are of initiatives and sum to more than 100.
doelgroep_verdeling <- function(df, label) {
  df |>
    separate_rows(doelgroep, sep = ";") |>
    mutate(doelgroep = str_squish(doelgroep)) |>
    filter(doelgroep != "", !is.na(doelgroep)) |>
    group_by(doelgroep) |>
    summarise(n_ongewogen = n(), aantal = sum(gewicht), .groups = "drop") |>
    mutate(populatie = label, pct = round(100 * aantal / sum(df$gewicht), 1)) |>
    arrange(desc(aantal))
}

bind_rows(
  doelgroep_verdeling(data_hercodeerd, "alle initiatieven"),
  doelgroep_verdeling(filter(data_hercodeerd, burgerinitiatief == "Ja"), "burgercollectieven")
)

readr::write_csv(vergelijking, here("output", "gewogen_verdelingen.csv"))
readr::write_csv(data_hercodeerd, here("output", "steekproef_gewogen_hercodeerd.csv"))
readr::write_csv(gewichten, here("output", "gewichten_per_koepel.csv"))




