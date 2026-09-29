"""
06_kaart_exacte_postcode.py
============================
Vlakkenkaart per EXACTE postcode (6 tekens, bijv. "1018PL") — net als
04_kaarten_maken.py een choropleth (gekleurde gebieden), maar dan tot op
het niveau van de individuele postcode i.p.v. PC4 (4 cijfers) of gemeente.

Belangrijk om te weten: er bestaan in Nederland GEEN officiële,
door de overheid gepubliceerde grenzen per exacte 6-cijferige postcode
(dat is meestal maar een klein stukje straat). PDOK/CBS levert alleen
vlakken tot op PC4-niveau. Om een postcode toch als vlak te kunnen tonen,
berekent dit script zelf vlakken met een Voronoi-diagram: elke postcode
krijgt het gebied dat dichter bij die postcode ligt dan bij enige andere
postcode in je bestand, geknipt op de landsgrens van Nederland. Dat is een
gangbare, wiskundig nette manier om van punten naar vlakken te gaan (zo
maakt men bijv. ook verzorgingsgebied-kaarten), maar het zijn dus
BEREKENDE nabijheidsgebieden, geen officiële postcodegrenzen. In dichte
gebieden (binnenstad) worden de vlakken vanzelf klein, in dunbevolkte
gebieden groot — dat is een prettige bijkomstigheid, geen toeval.

VOOR GEBRUIK:
  1. Vul INPUT_FILE en COLUMN_MAP hieronder in. Standaard staat dit al
     goed voor Locaties_per_postcode.xlsx (kolommen: Postcode, kvknr).
  2. pip install pandas openpyxl requests scipy numpy shapely
  3. python 06_kaart_exacte_postcode.py
  4. Open kaart_exacte_postcode.html in je browser.

Let op: dit script geocodeert elke postcode via de gratis PDOK
Locatieserver — dat zijn evenveel verzoeken als je unieke postcodes
hebt. Voor ~20.000 postcodes duurt dat de EERSTE keer ca. 15 minuten.
Resultaten worden lokaal gecachet (postcode_coords_cache.json), dus een
volgende run (bijv. na een nieuwe Excel-export) is vrijwel meteen klaar
— alleen nieuwe/gewijzigde postcodes worden opnieuw opgezocht.
"""

import json
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import numpy as np
import pandas as pd
import requests
from scipy.spatial import Voronoi
from shapely.geometry import Polygon, shape, mapping
from shapely.ops import transform, unary_union

# ----------------------------------------------------------------------------
INPUT_FILE = "Locaties_per_postcode.xlsx"
OUTPUT_FILE = "kaart_exacte_postcode.html"
NIET_GEVONDEN_FILE = "exacte_postcodes_niet_gevonden.xlsx"
AANTALLEN_FILE = "exacte_postcodes_aantallen.xlsx"
CACHE_FILE = "postcode_coords_cache.json"
NL_GRENS_CACHE_FILE = "nl_landsgrens_cache.json"  # voorkomt opnieuw ophalen+samenvoegen van 342 gemeentegrenzen

# Kolomnamen zoals ze in jouw dataset heten. 'aantal' mag een kolom zijn die
# al een telling bevat (zoals 'kvknr' in Locaties_per_postcode.xlsx) — dan
# worden waarden bij dubbele postcodes gesómmeerd. Staat hier in plaats
# daarvan een KvK-nummer-kolom, zet dan AANTAL_IS_AL_EEN_TELLING op False;
# dan wordt per postcode het aantal unieke KvK-nummers geteld.
COLUMN_MAP = {
    "postcode": "Postcode",
    "aantal": "kvknr",
}
AANTAL_IS_AL_EEN_TELLING = True

PDOK_LOCATIESERVER_URL = "https://api.pdok.nl/bzk/locatieserver/search/v3_1/free"
GEMEENTE_ITEMS_URL = "https://api.pdok.nl/kadaster/brk-bestuurlijke-gebieden/ogc/v1/collections/gemeentegebied/items"
GEOCODE_WORKERS = 12  # gelijktijdige verzoeken naar PDOK — hoger = sneller maar zwaarder voor hun (gratis) dienst

KLEUREN = ["#fee5d9", "#fcae91", "#fb6a4a", "#de2d26", "#a50f15"]  # licht -> donker
SIMPLIFY_TOLERANTIE = 0.00008  # graden, ~9m — zie 04_kaarten_maken.py voor uitleg
COORDINATEN_DECIMALEN = 5
# ----------------------------------------------------------------------------

POSTCODE_PATRONA = re.compile(r"^\d{4}[A-Z]{2}$")


def naar_geldige_postcode(waarde) -> str:
    ruw = str(waarde).strip().upper().replace(" ", "")
    return ruw if POSTCODE_PATRONA.match(ruw) else ""


# ----------------------------------------------------------------------------
# Geocoderen (met lokale cache, zodat een 2e run niet alles opnieuw doet)
# ----------------------------------------------------------------------------

def laad_json_cache(pad_str: str) -> dict:
    pad = Path(pad_str)
    if pad.exists():
        with pad.open(encoding="utf-8") as f:
            return json.load(f)
    return {}


def bewaar_json_cache(pad_str: str, data):
    with open(pad_str, "w", encoding="utf-8") as f:
        json.dump(data, f)


def geocodeer_postcode(postcode: str):
    try:
        resp = requests.get(
            PDOK_LOCATIESERVER_URL,
            params={"q": postcode, "fq": "type:postcode", "rows": 1},
            timeout=15,
        )
        resp.raise_for_status()
        docs = resp.json().get("response", {}).get("docs", [])
        if not docs:
            return None
        match = re.match(r"POINT\(([-\d.]+) ([-\d.]+)\)", docs[0]["centroide_ll"])
        if not match:
            return None
        lon, lat = float(match.group(1)), float(match.group(2))
        return (lat, lon)
    except (requests.RequestException, KeyError, ValueError):
        return None


def geocodeer_alles(postcodes: list, cache: dict) -> dict:
    nog_te_doen = [pc for pc in postcodes if pc not in cache]
    print(f"{len(postcodes)} unieke postcodes, {len(nog_te_doen)} nog niet in cache — geocoderen via PDOK...")
    if nog_te_doen:
        t0 = time.time()
        klaar = 0
        with ThreadPoolExecutor(max_workers=GEOCODE_WORKERS) as executor:
            toekomsten = {executor.submit(geocodeer_postcode, pc): pc for pc in nog_te_doen}
            for toekomst in as_completed(toekomsten):
                pc = toekomsten[toekomst]
                resultaat = toekomst.result()
                cache[pc] = list(resultaat) if resultaat else None
                klaar += 1
                if klaar % 500 == 0 or klaar == len(nog_te_doen):
                    verstreken = time.time() - t0
                    print(f"  {klaar}/{len(nog_te_doen)} gedaan ({verstreken:.0f}s, ~{klaar / verstreken:.1f}/s)")
                if klaar % 2000 == 0:
                    bewaar_json_cache(CACHE_FILE, cache)
        bewaar_json_cache(CACHE_FILE, cache)
        print(f"Geocoderen klaar in {time.time() - t0:.0f}s.")
    return {pc: tuple(cache[pc]) for pc in postcodes if cache.get(pc) is not None}


def haal_ogc_features(items_url: str, omschrijving: str) -> list:
    features = []
    url = f"{items_url}?f=json&limit=1000"
    pagina = 1
    while url and pagina <= 30:
        resp = requests.get(url, timeout=60)
        resp.raise_for_status()
        data = resp.json()
        features.extend(data.get("features", []))
        print(f"  [{omschrijving}] pagina {pagina}: {len(data.get('features', []))} opgehaald (totaal {len(features)})")
        url = next((link["href"] for link in data.get("links", []) if link.get("rel") == "next"), None)
        pagina += 1
    return features


def haal_nl_landsgrens_op():
    """Unie van alle gemeentegrenzen = de landsgrens van Nederland. Wordt
    lokaal gecachet, want het samenvoegen van 342 vlakken kost even tijd en
    hoeft maar één keer per run/dataset."""
    cache_pad = Path(NL_GRENS_CACHE_FILE)
    if cache_pad.exists():
        print(f"Landsgrens uit cache geladen ({NL_GRENS_CACHE_FILE}).")
        with cache_pad.open(encoding="utf-8") as f:
            return shape(json.load(f))

    print("Landsgrens berekenen (unie van alle gemeentegrenzen bij PDOK)...")
    features = haal_ogc_features(GEMEENTE_ITEMS_URL, "gemeentegrenzen")
    t0 = time.time()
    polygonen = [shape(f["geometry"]) for f in features]
    nl = unary_union(polygonen)
    print(f"  klaar in {time.time() - t0:.1f}s ({len(features)} gemeenten samengevoegd)")
    with cache_pad.open("w", encoding="utf-8") as f:
        json.dump(mapping(nl), f)
    return nl


# ----------------------------------------------------------------------------
# Voronoi: van punten naar vlakken
# ----------------------------------------------------------------------------

def bouw_voronoi_vlakken(postcodes: list, coords: dict, nl_grens) -> list:
    """Geeft een lijst (postcode, shapely-vlak) terug: één vlak per
    postcode, geknipt op de Nederlandse landsgrens."""
    geldige_pcs = [pc for pc in postcodes if pc in coords]
    punten = np.array([[coords[pc][1], coords[pc][0]] for pc in geldige_pcs])  # (lon, lat)

    minx, miny, maxx, maxy = nl_grens.bounds
    pad = 2.0
    # Verre hulppunten rondom, zodat scipy alleen begrensde (gesloten)
    # Voronoi-cellen teruggeeft voor de échte postcodepunten — zonder deze
    # trucs krijgen randpunten open/oneindige cellen.
    hulppunten = np.array([
        [minx - pad, miny - pad], [minx - pad, maxy + pad],
        [maxx + pad, miny - pad], [maxx + pad, maxy + pad],
        [(minx + maxx) / 2, miny - pad * 3], [(minx + maxx) / 2, maxy + pad * 3],
        [minx - pad * 3, (miny + maxy) / 2], [maxx + pad * 3, (miny + maxy) / 2],
    ])
    alle_punten = np.vstack([punten, hulppunten])

    print(f"Voronoi-diagram berekenen voor {len(geldige_pcs)} postcodes...")
    t0 = time.time()
    vor = Voronoi(alle_punten)
    print(f"  klaar in {time.time() - t0:.1f}s")

    print("Vlakken knippen op de landsgrens en vereenvoudigen...")
    t0 = time.time()

    def afronden(x, y, z=None):
        return (round(x, COORDINATEN_DECIMALEN), round(y, COORDINATEN_DECIMALEN))

    resultaten = []
    for i, pc in enumerate(geldige_pcs):
        region = vor.regions[vor.point_region[i]]
        if not region or -1 in region or len(region) < 3:
            continue  # onbegrensde cel (zou door de hulppunten niet meer moeten voorkomen)
        cel = Polygon([vor.vertices[v] for v in region])
        geknipt = cel.intersection(nl_grens)
        if geknipt.is_empty:
            continue
        vereenvoudigd = geknipt.simplify(SIMPLIFY_TOLERANTIE, preserve_topology=True)
        if vereenvoudigd.is_empty:
            continue
        afgerond = transform(afronden, vereenvoudigd)
        resultaten.append((pc, afgerond))
        if (i + 1) % 5000 == 0:
            print(f"  {i + 1}/{len(geldige_pcs)}...")
    print(f"  klaar in {time.time() - t0:.0f}s, {len(resultaten)} vlakken.")
    return resultaten


def bereken_breaks(waarden: list, aantal_klassen: int) -> list:
    waarden = sorted(w for w in waarden if w > 0)
    if not waarden:
        return list(range(1, aantal_klassen))
    breaks = sorted(set(
        waarden[min(int(len(waarden) * i / aantal_klassen), len(waarden) - 1)]
        for i in range(1, aantal_klassen)
    ))
    while len(breaks) < aantal_klassen - 1:
        breaks.append(breaks[-1] + 1)
    return breaks


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="UTF-8">
<title>__TITEL__</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
<style>
  html, body { margin:0; padding:0; height:100%; font-family: system-ui, Arial, sans-serif; }
  #map { position:absolute; top:0; left:0; right:0; bottom:0; background:#fff; }
  #titel { position:absolute; top:10px; left:50px; z-index:1000; background:#fff; padding:8px 14px; border-radius:6px; box-shadow:0 1px 4px rgba(0,0,0,.3); font-size:14px; font-weight:bold; }
  .legenda { background:#fff; padding:10px 12px; border-radius:6px; box-shadow:0 1px 4px rgba(0,0,0,.3); font-size:12px; line-height:1.6; max-width:220px; }
  .legenda span { display:inline-block; width:14px; height:14px; margin-right:6px; vertical-align:middle; border:1px solid #999; }
</style>
</head>
<body>
<div id="titel">__TITEL__</div>
<div id="map"></div>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script>
const GEOJSON = __GEOJSON__;
const BREAKS = __BREAKS__;
const KLEUREN = __KLEUREN__;

const map = L.map('map', { preferCanvas: true }).setView([52.1, 5.3], 7);
// Bewust GEEN achtergrondkaart (tegelserver) — na drie providers die
// allemaal om een andere reden niet betrouwbaar bleken (osm.org blokkeert
// dit soort verspreide bestanden, CARTO vraagt alsnog een API-key, en
// PDOK's dekking liep niet ver genoeg door rond Nederland), is een effen
// achtergrond de simpelste, altijd-werkende oplossing: geen enkele
// afhankelijkheid van een externe dienst meer.

function kleurVoor(aantal) {
  for (let i = 0; i < BREAKS.length; i++) {
    if (aantal <= BREAKS[i]) return KLEUREN[i];
  }
  return KLEUREN[KLEUREN.length - 1];
}

const laag = L.geoJSON(GEOJSON, {
  style: (feature) => ({
    fillColor: kleurVoor(feature.properties.aantal || 0),
    fillOpacity: 0.8,
    color: '#999',
    weight: 0.5,
  }),
  onEachFeature: (feature, layer) => {
    // Tienduizenden vlakken: klik-i.p.v.-hover, dat is veel lichter voor de browser.
    layer.bindPopup(`<strong>${feature.properties.postcode}</strong><br>${feature.properties.aantal || 0} initiatieven`);
    layer.on('click', () => layer.setStyle({ weight: 2, color: '#000' }));
    layer.on('popupclose', () => layer.setStyle({ weight: 0.5, color: '#999' }));
  },
}).addTo(map);

function pasKaartAanOpInhoud() {
  map.invalidateSize();
  map.fitBounds(laag.getBounds());
}
function wachtTotZichtbaarEnFit(pogingenOver) {
  const grootte = map.getSize();
  if (grootte.x > 0 && grootte.y > 0) {
    pasKaartAanOpInhoud();
  } else if (pogingenOver > 0) {
    requestAnimationFrame(() => wachtTotZichtbaarEnFit(pogingenOver - 1));
  }
}
pasKaartAanOpInhoud();
wachtTotZichtbaarEnFit(300);
window.addEventListener('resize', pasKaartAanOpInhoud);

const legenda = L.control({ position: 'bottomright' });
legenda.onAdd = () => {
  const div = L.DomUtil.create('div', 'legenda');
  let html = '<strong>Aantal initiatieven</strong><br>';
  let vorige = 0;
  BREAKS.forEach((b, i) => {
    html += `<span style="background:${KLEUREN[i]}"></span>${vorige}${vorige === b ? '' : ' - ' + b}<br>`;
    vorige = b + 1;
  });
  html += `<span style="background:${KLEUREN[KLEUREN.length - 1]}"></span>${vorige}+`;
  html += '<br><small>Vlakken zijn berekende nabijheidsgebieden per postcode (Voronoi), geen officiële postcodegrenzen.</small>';
  div.innerHTML = html;
  return div;
};
legenda.addTo(map);
</script>
</body>
</html>
"""


def main():
    df = pd.read_excel(INPUT_FILE)
    df.columns = [str(c).strip() for c in df.columns]
    ontbrekend = [c for c in COLUMN_MAP.values() if c not in df.columns]
    if ontbrekend:
        raise SystemExit(
            f"Kolommen niet gevonden: {ontbrekend}\n"
            f"Beschikbaar in je bestand: {list(df.columns)}\n"
            f"Pas COLUMN_MAP bovenin het script aan."
        )
    print(f"Ingelezen: {len(df)} rijen uit {INPUT_FILE}")

    df["_postcode"] = df[COLUMN_MAP["postcode"]].apply(naar_geldige_postcode)
    ongeldig = (df["_postcode"] == "").sum()
    if ongeldig:
        print(f"Let op: {ongeldig} rijen hebben geen herkenbare 6-tekens postcode en worden overgeslagen.")
    df = df[df["_postcode"] != ""]

    if AANTAL_IS_AL_EEN_TELLING:
        aantallen = df.groupby("_postcode")[COLUMN_MAP["aantal"]].sum()
    else:
        aantallen = df.groupby("_postcode")[COLUMN_MAP["aantal"]].nunique()
    aantallen = aantallen.astype(int).to_dict()

    cache = laad_json_cache(CACHE_FILE)
    coords = geocodeer_alles(list(aantallen.keys()), cache)

    niet_gevonden_geocodeerd = [pc for pc in aantallen if pc not in coords]

    nl_grens = haal_nl_landsgrens_op()
    vlakken = bouw_voronoi_vlakken(list(aantallen.keys()), coords, nl_grens)

    gebruikte_pcs = {pc for pc, _ in vlakken}
    niet_gevonden = [
        {"postcode": pc, "aantal": aantallen[pc],
         "reden": "kon niet gegeocodeerd worden" if pc in niet_gevonden_geocodeerd else "viel buiten de landsgrens na knippen"}
        for pc in aantallen if pc not in gebruikte_pcs
    ]
    if niet_gevonden:
        pd.DataFrame(niet_gevonden).sort_values("aantal", ascending=False).to_excel(NIET_GEVONDEN_FILE, index=False)
        print(f"Let op: {len(niet_gevonden)} postcodes staan niet op de kaart. Zie {NIET_GEVONDEN_FILE}.")

    features = [
        {"type": "Feature", "properties": {"postcode": pc, "aantal": aantallen[pc]}, "geometry": mapping(geom)}
        for pc, geom in vlakken
    ]

    pd.DataFrame([
        {"postcode": f["properties"]["postcode"], "aantal": f["properties"]["aantal"]}
        for f in features
    ]).sort_values("aantal", ascending=False).to_excel(AANTALLEN_FILE, index=False)
    print(f"Volledige telling weggeschreven naar {AANTALLEN_FILE}.")

    breaks = bereken_breaks([f["properties"]["aantal"] for f in features], len(KLEUREN))

    html = HTML_TEMPLATE
    html = html.replace("__GEOJSON__", json.dumps({"type": "FeatureCollection", "features": features}, ensure_ascii=False, separators=(",", ":")))
    html = html.replace("__BREAKS__", json.dumps(breaks))
    html = html.replace("__KLEUREN__", json.dumps(KLEUREN))
    html = html.replace("__TITEL__", "Aantal initiatieven per exacte postcode (berekende vlakken)")
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    import os
    grootte_mb = os.path.getsize(OUTPUT_FILE) / (1024 * 1024)
    print(f"Klaar! {OUTPUT_FILE} ({grootte_mb:.1f} MB) — open in je browser.")


if __name__ == "__main__":
    main()
