"""
05_kaart_postcode.py
=====================
Bouwt een choropleth-kaart (gekleurde postcodevlakken) van het aantal
initiatieven per postcodegebied (PC4, de 4 cijfers), op basis van jouw
postcode-aggregaatbestand (kolommen: postcode + "kvknr" = aantal unieke
KvK-nummers).

Haalt de officiële postcodegrenzen (CBS Postcode4) automatisch op bij PDOK
(geen download nodig), koppelt ze aan jouw aantallen, en bouwt één
HTML-bestand met een kleurenkaart + legenda + hover-popups.

Er zijn ~4.000 postcodegebieden in Nederland — het ophalen gebeurt in
pagina's en kan daardoor een minuutje duren. Het resultaat wordt daarna in
één (groter) HTML-bestand gezet.

Deze versie heeft nog GEEN filters (fonds/organisatietype/opgericht in/KvK) —
dat vraagt om de volledige, gedetailleerde dataset i.p.v. dit
aggregaatbestand. Zodra deze basiskaart werkt, bouwen we die uitgebreide
versie als vervolgstap (zelfde als bij de gemeentekaart).

BELANGRIJK: dit script kon niet end-to-end getest worden met jouw echte data
(vertrouwelijk) of een live verbinding met de PDOK-service (de omgeving hier
heeft geen toegang tot dat soort externe sites). Werkt het niet meteen goed?
Stuur de foutmelding + de velden die het script print, dan pas ik het gericht
aan.

VOOR GEBRUIK:
  1. Vul COLUMN_MAP hieronder in.
  2. pip install pandas openpyxl requests
  3. python 05_kaart_postcode.py
  4. Open kaart_postcodes.html in je browser.
"""

import json
import re

import pandas as pd
import requests

# ----------------------------------------------------------------------------
INPUT_FILE = "aantal_per_postcode.xlsx"
OUTPUT_FILE = "kaart_postcodes.html"
NIET_GEVONDEN_RAPPORT = "postcodes_niet_gevonden.xlsx"

# Kolomnamen zoals ze in jouw postcode-aggregaatbestand heten.
COLUMN_MAP = {
    "postcode": "postcode",  # 4-cijferige postcode, of 6-tekens (script pakt dan zelf de eerste 4 cijfers)
    "aantal": "kvknr",       # kolom met het aantal unieke KvK-nummers per postcodegebied
}

POSTCODE4_ITEMS_URL = "https://api.pdok.nl/cbs/postcode4/ogc/v1/collections/postcode4/items"

KLEUREN = ["#fee5d9", "#fcae91", "#fb6a4a", "#de2d26", "#a50f15"]  # licht -> donker
# ----------------------------------------------------------------------------


def haal_postcodegrenzen_op():
    """Haalt alle PC4-postcodegebieden op bij de CBS/PDOK OGC API (gepagineerd)."""
    features = []
    url = f"{POSTCODE4_ITEMS_URL}?f=json&limit=1000"
    pagina = 1
    while url and pagina <= 30:  # veiligheidslimiet tegen oneindig doorpaginenren
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        batch = data.get("features", [])
        features.extend(batch)
        print(f"  Pagina {pagina}: {len(batch)} postcodegebieden opgehaald (totaal {len(features)})")
        url = next((link["href"] for link in data.get("links", []) if link.get("rel") == "next"), None)
        pagina += 1

    if not features:
        raise SystemExit(
            "Kon geen postcodegrenzen ophalen bij PDOK. Check je internetverbinding, "
            "of download handmatig een GeoJSON via "
            "https://www.pdok.nl/introductie/-/article/cbs-postcode4 "
            "en pas dit script aan om dat lokale bestand in te lezen."
        )
    print(f"Postcodegrenzen opgehaald: {len(features)} gebieden.")
    return features


def bepaal_property(features: list, kandidaten: list, omschrijving: str):
    beschikbaar = list(features[0]["properties"].keys())
    for kandidaat in kandidaten:
        if kandidaat in beschikbaar:
            print(f"{omschrijving}-property gevonden: '{kandidaat}'")
            return kandidaat
    print(f"Let op: geen {omschrijving}-property gevonden. Beschikbare velden: {beschikbaar}")
    return None


def naar_pc4(waarde) -> str:
    """Pakt de eerste 4 cijfers uit een postcode, of "" als er geen 4
    cijfers in staan."""
    match = re.search(r"\d{4}", str(waarde))
    return match.group() if match else ""


def main():
    df = pd.read_excel(INPUT_FILE)
    df.columns = [str(c).strip() for c in df.columns]
    if COLUMN_MAP["postcode"] not in df.columns or COLUMN_MAP["aantal"] not in df.columns:
        raise SystemExit(
            f"Kolommen niet gevonden. Beschikbaar in je bestand: {list(df.columns)}\n"
            f"Pas COLUMN_MAP bovenin het script aan."
        )

    features = haal_postcodegrenzen_op()

    pc_kandidaten = ["postcode4", "pc4", "postcode", "statcode"]
    jaar_kandidaten = ["jaar", "year", "statjaar"]
    pc_property = bepaal_property(features, pc_kandidaten, "postcode")
    if pc_property is None:
        raise SystemExit(
            "Kon geen postcode-veld vinden — pas de kandidatenlijst in het script "
            "aan naar het juiste veld uit de bovenstaande lijst met beschikbare velden."
        )
    jaar_property = bepaal_property(features, jaar_kandidaten, "jaar (voor filtering op meest recente editie)")

    # Als de dataset meerdere jaren per postcode bevat, alleen de meest
    # recente editie per postcode aanhouden.
    if jaar_property:
        nieuwste = {}
        for feature in features:
            pc = str(feature["properties"].get(pc_property, "")).zfill(4)
            jaar = feature["properties"].get(jaar_property, 0)
            if pc not in nieuwste or jaar > nieuwste[pc]["properties"].get(jaar_property, 0):
                nieuwste[pc] = feature
        features = list(nieuwste.values())
        print(f"Na filteren op meest recente editie: {len(features)} postcodegebieden.")

    aantallen = {}
    for _, row in df.iterrows():
        pc = naar_pc4(row[COLUMN_MAP["postcode"]])
        if pc:
            aantallen[pc] = row[COLUMN_MAP["aantal"]]

    gebruikte_pcs = set()
    for feature in features:
        pc = str(feature["properties"].get(pc_property, "")).zfill(4)
        aantal = aantallen.get(pc)
        feature["properties"]["aantal"] = int(aantal) if aantal is not None else 0
        if aantal is not None:
            gebruikte_pcs.add(pc)

    # Postcodes uit jouw bestand die niet aan een bekend Nederlands
    # postcodegebied gekoppeld konden worden (bijv. buiten Nederland, of
    # een niet-bestaande/verouderde postcode) -> apart rapporteren.
    niet_gevonden = [
        {"postcode_in_bestand": pc, "aantal": aantal}
        for pc, aantal in aantallen.items()
        if pc not in gebruikte_pcs
    ]
    if niet_gevonden:
        pd.DataFrame(niet_gevonden).to_excel(NIET_GEVONDEN_RAPPORT, index=False)
        print(f"Let op: {len(niet_gevonden)} postcodes konden niet gekoppeld worden. "
              f"Zie {NIET_GEVONDEN_RAPPORT}.")

    waarden = sorted(f["properties"]["aantal"] for f in features if f["properties"]["aantal"] > 0)
    if waarden:
        n = len(KLEUREN)
        breaks = sorted(set(waarden[min(int(len(waarden) * i / n), len(waarden) - 1)] for i in range(1, n)))
        while len(breaks) < n - 1:
            breaks.append(breaks[-1] + 1)
    else:
        breaks = [1, 2, 3, 4]

    html = HTML_TEMPLATE
    html = html.replace("__GEOJSON__", json.dumps({"type": "FeatureCollection", "features": features}, ensure_ascii=False))
    html = html.replace("__BREAKS__", json.dumps(breaks))
    html = html.replace("__KLEUREN__", json.dumps(KLEUREN))
    html = html.replace("__TITEL__", "Aantal initiatieven per postcodegebied (PC4)")
    html = html.replace("__NAAM_PROPERTY__", json.dumps(pc_property))

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Klaar! Open {OUTPUT_FILE} in je browser.")


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="UTF-8">
<title>__TITEL__</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
<style>
  html, body { margin:0; padding:0; height:100%; font-family: system-ui, Arial, sans-serif; }
  #map { position:absolute; top:0; left:0; right:0; bottom:0; }
  #titel { position:absolute; top:10px; left:50px; z-index:1000; background:#fff; padding:8px 14px; border-radius:6px; box-shadow:0 1px 4px rgba(0,0,0,.3); font-size:14px; font-weight:bold; }
  .legenda { background:#fff; padding:10px 12px; border-radius:6px; box-shadow:0 1px 4px rgba(0,0,0,.3); font-size:12px; line-height:1.6; }
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
const NAAM_PROPERTY = __NAAM_PROPERTY__;

const map = L.map('map');
L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
  attribution: '&copy; OpenStreetMap-bijdragers'
}).addTo(map);

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
    const naam = NAAM_PROPERTY ? (feature.properties[NAAM_PROPERTY] || '(onbekend)') : '(onbekend)';
    layer.bindTooltip(`<strong>${naam}</strong><br>${feature.properties.aantal || 0} initiatieven`, { sticky: true });
    layer.on('mouseover', () => layer.setStyle({ weight: 2, color: '#000' }));
    layer.on('mouseout', () => layer.setStyle({ weight: 0.5, color: '#999' }));
  },
}).addTo(map);

map.fitBounds(laag.getBounds());

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
  div.innerHTML = html;
  return div;
};
legenda.addTo(map);
</script>
</body>
</html>
"""


if __name__ == "__main__":
    main()
