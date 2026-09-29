"""
02_geocoderen.py
=================
Zet adres/postcode/plaats om naar breedtegraad/lengtegraad (lat/lon), zodat
de gegevens op een kaart gezet kunnen worden.

Twee bronnen, automatisch gekozen per rij:
  1. PDOK Locatieserver — gratis, geen API-key nodig, beste dekking voor
     Europees Nederland (BAG-adressen).
  2. Nominatim (OpenStreetMap) — fallback voor adressen die PDOK niet vindt,
     en voor Caribisch Nederland (Bonaire, Sint Eustatius, Saba), dat niet
     in de BAG/PDOK zit.

Resultaten worden gecachet in geocode_cache.json, zodat je het script
opnieuw kan draaien zonder alles opnieuw op te vragen.

VOOR GEBRUIK:
  1. Zorg dat 01_dedupliceren.py al is gedraaid (of pas INPUT_FILE aan).
  2. Vul COLUMN_MAP hieronder in. "postcode" mag op None staan als je geen
     aparte postcode-kolom hebt (bijv. omdat de postcode al in het adres
     verwerkt zit).
  3. python 02_geocoderen.py
"""

import json
import time
from pathlib import Path

import pandas as pd
import requests

# ----------------------------------------------------------------------------
INPUT_FILE = "steekproef_herhaling_ontdubbeld.xlsx"
OUTPUT_FILE = "Dataset_met_coordinaten.xlsx"
CACHE_FILE = "geocode_cache.json"

# "postcode" mag None zijn als je geen aparte postcode-kolom hebt.
COLUMN_MAP = {
    "adres": "Adres",
    "postcode": "Postcode",
    "plaats": "Plaats",
}

NOMINATIM_USER_AGENT = "burgerinitiatieven-kaart (contact: spaanhilde@gmail.com)"
# ----------------------------------------------------------------------------


def load_data(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]  # spaties in kolomnamen weghalen

    verplicht = ["adres", "plaats"]
    missing = [COLUMN_MAP[k] for k in verplicht if COLUMN_MAP[k] not in df.columns]
    if COLUMN_MAP.get("postcode") and COLUMN_MAP["postcode"] not in df.columns:
        print(f"Let op: postcode-kolom '{COLUMN_MAP['postcode']}' niet gevonden, "
              "geocoderen gebeurt zonder aparte postcode.")
        COLUMN_MAP["postcode"] = None
    if missing:
        raise SystemExit(
            "Kolommen niet gevonden in het Excel-bestand: "
            f"{missing}\nControleer COLUMN_MAP bovenin dit script tegen de "
            f"echte kolomnamen: {list(df.columns)}"
        )
    return df


def load_cache() -> dict:
    if Path(CACHE_FILE).exists():
        return json.loads(Path(CACHE_FILE).read_text())
    return {}


def save_cache(cache: dict):
    Path(CACHE_FILE).write_text(json.dumps(cache, ensure_ascii=False, indent=2))


def build_query(row) -> str:
    delen = []
    if pd.notna(row[COLUMN_MAP["adres"]]):
        delen.append(str(row[COLUMN_MAP["adres"]]))
    if COLUMN_MAP.get("postcode") and pd.notna(row[COLUMN_MAP["postcode"]]):
        delen.append(str(row[COLUMN_MAP["postcode"]]))
    if pd.notna(row[COLUMN_MAP["plaats"]]):
        delen.append(str(row[COLUMN_MAP["plaats"]]))
    return ", ".join(d for d in delen if d.strip())


def geocode_pdok(query: str):
    try:
        resp = requests.get(
            "https://api.pdok.nl/bzk/locatieserver/search/v3_1/free",
            params={"q": query, "rows": 1, "fq": "type:adres"},
            timeout=10,
        )
        resp.raise_for_status()
        docs = resp.json().get("response", {}).get("docs", [])
        if not docs:
            return None
        centroid = docs[0].get("centroide_ll")  # "POINT(lon lat)"
        if not centroid:
            return None
        lon, lat = centroid.replace("POINT(", "").replace(")", "").split()
        return float(lat), float(lon)
    except Exception:
        return None


def geocode_nominatim(query: str):
    try:
        resp = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={"q": query, "format": "json", "limit": 1},
            headers={"User-Agent": NOMINATIM_USER_AGENT},
            timeout=10,
        )
        resp.raise_for_status()
        results = resp.json()
        if not results:
            return None
        return float(results[0]["lat"]), float(results[0]["lon"])
    except Exception:
        return None


def geocode_row(query: str, cache: dict):
    if not query:
        return None, None, "geen adres"
    if query in cache:
        lat, lon, bron = cache[query]
        return lat, lon, bron

    result = geocode_pdok(query)
    bron = "pdok"
    if result is None:
        time.sleep(1)  # Nominatim vraagt max 1 request/seconde
        result = geocode_nominatim(query)
        bron = "nominatim" if result else "niet gevonden"

    lat, lon = result if result else (None, None)
    cache[query] = (lat, lon, bron)
    return lat, lon, bron


def main():
    df = load_data(INPUT_FILE)
    cache = load_cache()

    lats, lons, bronnen = [], [], []
    for i, row in df.iterrows():
        query = build_query(row)
        lat, lon, bron = geocode_row(query, cache)
        lats.append(lat)
        lons.append(lon)
        bronnen.append(bron)
        if (i + 1) % 25 == 0:
            print(f"{i + 1}/{len(df)} verwerkt...")
            save_cache(cache)  # tussentijds opslaan

    df["latitude"] = lats
    df["longitude"] = lons
    df["geocode_bron"] = bronnen
    save_cache(cache)

    niet_gevonden = df["latitude"].isna().sum()
    print(f"Klaar. {len(df) - niet_gevonden}/{len(df)} adressen gevonden, "
          f"{niet_gevonden} niet gevonden (controleer 'geocode_bron' = 'niet gevonden').")

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Opgeslagen als: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
