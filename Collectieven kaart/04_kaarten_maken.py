"""
04_kaarten_maken.py

Verantwoording
--------------
Dit script is geschreven met AI-assistentie (Claude). De keuzes over wat
er geteld en getekend wordt zijn van mij, evenals de controle op de
uitkomsten. Het script schrijft daarom bij elke run de volledige telling
weg naar gemeenten_aantallen.xlsx en postcodes_aantallen.xlsx (hetzelfde voor missende aantallen); die zijn
regel voor regel te vergelijken met de brontelling in Excel. De
initiatieven die niet gekoppeld konden worden staan apart in
plaatsnamen_niet_gevonden.xlsx, gemeenten_niet_gevonden.xlsx,
postcodes_niet_gevonden.xlsx en postcode4_niet_gevonden.xlsx, zodat
zichtbaar is wat er buiten de kaart valt en waarom.

Bij deze versie valt 1,8% (gemeentekaart) resp. 2,6% (postcodekaart) buiten beeld; de uitsplitsing staat in de vier niet-gevonden-bestanden.

Het script bouwt twee kaarten van het aantal burgerinitiatieven in Nederland
(kaart_gemeenten.html en kaart_postcodes.html). Beide zijn gewoon een
los HTML-bestand met Leaflet erin, dus die kun je mailen en iemand kan
'm openen zonder dat er iets geinstalleerd hoeft te worden (als het goed is nu ook 
zonder watermerk).

1) kaart_gemeenten.html
Input is Locaties_per_gemeente.xlsx, met een telling per PLAATSNAAM
(niet per gemeente). Het script zoekt zelf per plaatsnaam de gemeente
op via de PDOK Locatieserver (api.pdok.nl/bzk/locatieserver, zie
https://www.pdok.nl/restful-locatieserver voor de documentatie).
Kostte me wat uitzoekwerk omdat gewone spreektaal-namen als "Den Haag"
niet als zodanig in de officiële (BAG-)naamgeving voorkomen (dat is
's-Gravenhage) en er ook plaatsen zijn die met een provincie-afkorting
worden aangeduid ("Etten Gld") om ze te onderscheiden van een
gelijknamig dorp elders. Beide gevallen worden hieronder apart
afgevangen.

2) kaart_postcodes.html
Input is Locaties_per_postcode.xlsx, telling per volledige (6-tekens)
postcode. Hieruit worden twee lagen gebouwd die op zoomniveau wisselen:
uitgezoomd zie je gekleurde PC4-vlakken (postcodegebied, de eerste 4
cijfers), inzoomen (vanaf ZOOM_DREMPEL_PUNTEN) geeft losse puntjes per
exacte postcode. Beide lagen komen uit dezelfde brontelling, dus een
vlak en de puntjes erin tellen altijd precies bij elkaar op.

Grenzen komen automatisch van PDOK (gemeentegrenzen via de Kadaster
BRK, postcodegrenzen via CBS/PDOK). Dit hoeft dus niet los gedownload te
worden, dat gebeurt bij het draaien van het script.

Een paar dingen die niet meteen voor de hand liggen maar wel expres
zijn:
- PDOK stuurt per postcodevlak een stuk of 140 CBS-kenmerken mee
  (inkomen, afstand tot school/huisarts, etc.) waar ik niks mee doe
  voor deze kaart - die gaan er meteen af, blijft alleen de postcode +
  het aantal over. En van elke postcode zit er in de PDOK-data ook nog
  een stapel oudere jaargangen bij; alleen de nieuwste wordt gebruikt.
- De grenzen van PDOK zijn nauwkeuriger dan je op een scherm ooit kan
  zien (submeter-precisie), en dat kost onnodig veel MB aan
  coördinaten. Daarom worden de lijnen vereenvoudigd  en coördinaten afgerond.
- Een gemeente heeft bestuurlijk ook zeggenschap over water binnen de
  grens (IJsselmeer, Waddenzee...), dus zonder ingrijpen kleurt dat
  water gewoon mee met het aantal initiatieven van die gemeente. Als
  truc gebruik ik de unie van alle PC4-postcodevlakken als knipmasker:
  postcodes bestaan nu eenmaal niet op open water.
- De postcodekaart heeft ~4000 vlakken en ~20.000 puntjes, en dat werd
  merkbaar traag met Leaflets standaardinstellingen. Met
  preferCanvas: true (canvas-rendering i.p.v. losse SVG-elementen per
  vlak) gaat dat een stuk soepeler.
- Er zit bewust geen achtergrondkaart (tegelserver, zoals OpenStreetMap
  of CARTO) in. Een effen achtergrond werkt overal, zonder internetverbinding nodig te hebben
  op het moment dat iemand de kaart opent.


Om te draaien:
1. Vul hieronder INPUT_FILE_GEMEENTE/COLUMN_MAP_GEMEENTE en
   INPUT_FILE_POSTCODE/COLUMN_MAP_POSTCODE in.
2. pip install pandas openpyxl requests shapely
3. python 04_kaarten_maken.py
4. Open kaart_gemeenten.html en kaart_postcodes.html in de browser.
"""

import json
import math
import re
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import pandas as pd
import requests
from shapely.geometry import Polygon, shape, mapping
from shapely.ops import unary_union

# ----------------------------------------------------------------------------
# Postcodekaart. Input is een telling per EXACTE (6-tekens) postcode; daaruit
# bouw ik zowel de PC4-vlakken (optellen tot 4 cijfers) als de punten per
# exacte postcode, dus dezelfde bron voor beide lagen en het klopt dus
# automatisch met elkaar.
INPUT_FILE_POSTCODE = "Locaties_per_postcode.xlsx"
COLUMN_MAP_POSTCODE = {
    "postcode": "Postcode",
    "aantal": "kvknr",  # let op: is hier gewoon een telling, geen echt KvK-nummer
}
AANTAL_IS_AL_EEN_TELLING_POSTCODE = True
POSTCODE_COORDS_CACHE_FILE = "postcode_coords_cache.json"
# Zoomniveau waarop de vlakken wisselen naar puntjes. Lager getal = eerder
# puntjes (maar drukker in beeld, want dan zie je nog meerdere PC4-gebieden
# tegelijk). 11 komt ongeveer overeen met inzoomen op een hele stad/plaats.
ZOOM_DREMPEL_PUNTEN = 11

# Gemeentekaart. Dit is een los bestand met een telling per PLAATSNAAM (dus
# niet per gemeente), en het script zoekt zelf de gemeente erbij op via PDOK.
INPUT_FILE_GEMEENTE = "Locaties_per_gemeente.xlsx"
COLUMN_MAP_GEMEENTE = {
    "plaatsnaam": "plnaam",
    "aantal": "kvknr",  # ook hier: al een telling, zie AANTAL_IS_AL_EEN_TELLING hieronder
}
AANTAL_IS_AL_EEN_TELLING_GEMEENTE = True
PLAATSNAAM_CACHE_FILE = "plaatsnaam_gemeente_cache.json"

OUTPUT_GEMEENTE = "kaart_gemeenten.html"
OUTPUT_POSTCODE = "kaart_postcodes.html"
NIET_GEVONDEN_GEMEENTE = "gemeenten_niet_gevonden.xlsx"
NIET_GEVONDEN_POSTCODE = "postcodes_niet_gevonden.xlsx"
# Schrijft precies weg wat er op de kaart komt (per gemeente/postcode het
# aantal). Fijn om naast een eigen Excel-telling te leggen als er ergens een
# raar getal op de kaart lijkt te staan.
AANTALLEN_GEMEENTE = "gemeenten_aantallen.xlsx"
AANTALLEN_POSTCODE = "postcodes_aantallen.xlsx"

# PDOK-endpoints (OGC API Features, dus gewoon GeoJSON terug). Documentatie
# staat op https://www.pdok.nl -> zoek op "bestuurlijke gebieden" resp.
# "cbs postcode4"; de Locatieserver (voor het geocoderen) staat op
# https://www.pdok.nl/restful-locatieserver.
GEMEENTE_ITEMS_URL = "https://api.pdok.nl/kadaster/brk-bestuurlijke-gebieden/ogc/v1/collections/gemeentegebied/items"
POSTCODE4_ITEMS_URL = "https://api.pdok.nl/cbs/postcode4/ogc/v1/collections/postcode4/items"
PDOK_LOCATIESERVER_URL = "https://api.pdok.nl/bzk/locatieserver/search/v3_1/free"
GEOCODE_WORKERS = 12  # aantal tegelijk lopende requests bij het opzoeken/geocoderen

KLEUREN_GEMEENTE = ["#eff3ff", "#bdd7e7", "#6baed6", "#3182bd", "#08519c"]  # licht naar donker
KLEUREN_POSTCODE = ["#fee5d9", "#fcae91", "#fb6a4a", "#de2d26", "#a50f15"]  # licht naar donker

# Tolerantie voor het vereenvoudigen van de lijnen, in graden (1 graad is
# ongeveer 111 km, dus dit is zo'n 9 meter). Heb dit even zitten testen: met
# een grovere waarde voor de gemeentekaart bleef er bij het knippen op het
# bewoond-gebied-masker (zie verderop) net iets te veel water over bij de
# Waddenzee, dus nu staat 'ie op dezelfde fijne waarde als de postcodekaart.
SIMPLIFY_TOLERANTIE_GEMEENTE = 0.00008
SIMPLIFY_TOLERANTIE_POSTCODE = 0.00008
COORDINATEN_DECIMALEN = 5  # ongeveer 1,1 m nauwkeurig, ruim genoeg voor op een webkaart

# Een gemeente heeft bestuurlijk ook zeggenschap over het water binnen de
# grens (IJsselmeer, Waddenzee, etc.), dus zonder ingrijpen kleurt dat gewoon
# mee met het aantal van die gemeente. Postcodegebieden (PC4) bestaan alleen
# uit bewoond land, dus de UNIE van alle PC4-vlakken samen is precies het
# bewoonde deel van Nederland. Dat gebruik ik als knipmasker voor elk
# gemeentevlak, vóór het tekenen. Zo krijg je dezelfde schone waterlijn als
# op de postcodekaart, zonder dat ik zelf per meer/zee iets hoef te doen.
BEWOOND_GEBIED_CACHE_FILE = "bewoond_gebied_cache.json"
# Extra fijne tolerantie (~3 m) omdat dit masker bepaalt waar de kustlijn
# precies komt te liggen. Een grovere vereenvoudiging liet bij de Waddenzee
# net iets te veel water binnen de gemeentegrens staan.
BEWOOND_GEBIED_SIMPLIFY_TOLERANTIE = 0.00003
# ----------------------------------------------------------------------------


def normaliseer(waarde) -> str:
    return str(waarde).strip().upper()


def naar_pc4(waarde) -> str:
    match = re.search(r"\d{4}", str(waarde))
    return match.group() if match else ""


POSTCODE6_PATROON = re.compile(r"^\d{4}[A-Z]{2}$")


def naar_geldige_postcode6(waarde) -> str:
    """Zet om naar '1234AB'-formaat. Geeft '' terug als het geen geldige
    6-tekens postcode blijkt te zijn."""
    ruw = str(waarde).strip().upper().replace(" ", "")
    return ruw if POSTCODE6_PATROON.match(ruw) else ""


def geocodeer_postcode6(postcode: str):
    """Vraagt bij de PDOK Locatieserver (zie PDOK_LOCATIESERVER_URL
    hierboven) het middelpunt van een postcode op. Geeft (lat, lon)
    terug, of None als 'ie niet gevonden wordt."""
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
        # WKT/GeoJSON schrijven coördinaten als (lon, lat), dus x eerst -
        # precies andersom dan je als mens zou verwachten. Zie ook
        # RFC 7946 (GeoJSON-spec, https://datatracker.ietf.org/doc/html/rfc7946),
        # sectie 3.1.1. Hier draai ik het om naar (lat, lon) voor Leaflet.
        lon, lat = float(match.group(1)), float(match.group(2))
        return (lat, lon)
    except (requests.RequestException, KeyError, ValueError):
        return None


def geocodeer_postcodes6(postcodes: list, cache: dict) -> dict:
    nog_te_doen = [pc for pc in postcodes if pc not in cache]
    print(f"{len(postcodes)} unieke exacte postcodes, {len(nog_te_doen)} nog niet in cache - geocoderen via PDOK...")
    if nog_te_doen:
        t0 = time.time()
        klaar = 0
        with ThreadPoolExecutor(max_workers=GEOCODE_WORKERS) as executor:
            toekomsten = {executor.submit(geocodeer_postcode6, pc): pc for pc in nog_te_doen}
            for toekomst in as_completed(toekomsten):
                pc = toekomsten[toekomst]
                resultaat = toekomst.result()
                cache[pc] = list(resultaat) if resultaat else None
                klaar += 1
                if klaar % 500 == 0 or klaar == len(nog_te_doen):
                    verstreken = time.time() - t0
                    print(f"  {klaar}/{len(nog_te_doen)} gedaan ({verstreken:.0f}s, ~{klaar / max(verstreken,0.01):.1f}/s)")
                if klaar % 2000 == 0:
                    bewaar_json_cache(POSTCODE_COORDS_CACHE_FILE, cache)
        bewaar_json_cache(POSTCODE_COORDS_CACHE_FILE, cache)
        print(f"Geocoderen klaar in {time.time() - t0:.0f}s.")
    return {pc: tuple(cache[pc]) for pc in postcodes if cache.get(pc) is not None}


def haal_ogc_features(items_url: str, omschrijving: str) -> list:
    """Haalt alle features op van een PDOK OGC API Features-endpoint.
    Dat soort endpoints geeft niet alles in één keer terug maar in
    pagina's van 1000, met een 'next'-link naar de volgende - die volg
    ik hier net zolang tot er geen meer is. (Zie ook: OGC API - Features,
    https://ogcapi.ogc.org/features/ - zo heet die standaard.)"""
    features = []
    url = f"{items_url}?f=json&limit=1000"
    pagina = 1
    while url and pagina <= 30:  # 30 pagina's is ruim genoeg, dit is een noodstop
        # Retry bij netwerkgehik (bijv. laptop die net uit slaapstand komt)
        # anders loopt een download die soms wel minuten duurt zomaar stuk op
        # één klein hapertje.
        for poging in range(3):
            try:
                resp = requests.get(url, timeout=60)
                resp.raise_for_status()
                data = resp.json()
                break
            except (requests.RequestException, ValueError):
                if poging == 2:
                    raise
                print(f"  [{omschrijving}] pagina {pagina}: netwerkfout, nieuwe poging...")
                time.sleep(3)
        batch = data.get("features", [])
        features.extend(batch)
        print(f"  [{omschrijving}] pagina {pagina}: {len(batch)} opgehaald (totaal {len(features)})")
        url = next((link["href"] for link in data.get("links", []) if link.get("rel") == "next"), None)
        pagina += 1
    return features


def bepaal_property(features: list, kandidaten: list, omschrijving: str):
    beschikbaar = list(features[0]["properties"].keys())
    for kandidaat in kandidaten:
        if kandidaat in beschikbaar:
            print(f"{omschrijving}-property gevonden: '{kandidaat}'")
            return kandidaat
    print(f"Let op: geen {omschrijving}-property gevonden. Beschikbare velden: {beschikbaar}")
    return None


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


def _perp_afstand(p, a, b) -> float:
    (x, y), (x1, y1), (x2, y2) = p, a, b
    dx, dy = x2 - x1, y2 - y1
    if dx == 0 and dy == 0:
        return math.hypot(x - x1, y - y1)
    t = ((x - x1) * dx + (y - y1) * dy) / (dx * dx + dy * dy)
    t = max(0.0, min(1.0, t))
    projx, projy = x1 + t * dx, y1 + t * dy
    return math.hypot(x - projx, y - projy)


def _rdp(punten: list, tolerantie: float) -> list:
   
    n = len(punten)
    if n < 3:
        return punten
    keep = [False] * n
    keep[0] = keep[-1] = True
    stack = [(0, n - 1)]
    while stack:
        start_i, end_i = stack.pop()
        if end_i - start_i < 2:
            continue
        start, end = punten[start_i], punten[end_i]
        dmax, index = -1.0, -1
        for i in range(start_i + 1, end_i):
            d = _perp_afstand(punten[i], start, end)
            if d > dmax:
                dmax, index = d, i
        if dmax > tolerantie and index != -1:
            keep[index] = True
            stack.append((start_i, index))
            stack.append((index, end_i))
    return [p for p, k in zip(punten, keep) if k]


def _vereenvoudig_ring(ring: list, tolerantie: float) -> list:
    vereenvoudigd = _rdp(ring, tolerantie)
   
    if len(vereenvoudigd) < 4:
        return ring
    return vereenvoudigd


def _rond_af(punt: list) -> list:
    return [round(punt[0], COORDINATEN_DECIMALEN), round(punt[1], COORDINATEN_DECIMALEN)]


def vereenvoudig_geometrie(geometry: dict, tolerantie: float) -> dict:
    """Werkt op een GeoJSON Polygon of MultiPolygon: haalt overbodige
    punten uit elke ring en rondt de coördinaten af."""
    type_ = geometry.get("type")

    def doe_ring(ring):
        vereenvoudigd = _vereenvoudig_ring(ring, tolerantie)
        return [_rond_af(p) for p in vereenvoudigd]

    if type_ == "Polygon":
        nieuwe_coords = [doe_ring(ring) for ring in geometry["coordinates"]]
    elif type_ == "MultiPolygon":
        nieuwe_coords = [[doe_ring(ring) for ring in polygon] for polygon in geometry["coordinates"]]
    else:
        return geometry
    return {"type": type_, "coordinates": nieuwe_coords}


# ----------------------------------------------------------------------------

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
const GEBRUIK_CLICK_POPUP = __GEBRUIK_CLICK_POPUP__;

// setView hier is eigenlijk maar een gok voor het beginscherm. Zonder dit
// zet fitBounds() verderop (op een kaart die nog nooit een view heeft gehad)
// wel het midden goed maar de zoom niet - vreemde Leaflet-eigenaardigheid,
// kostte me even puzzelen. fitBounds() corrigeert het toch naar de echte
// data zodra de kaart geladen is.
const map = L.map('map', { preferCanvas: true }).setView([52.1, 5.3], 7);
// Geen achtergrondkaart (tegelserver) - zie de toelichting bovenin het
// script. Dit gaat als los bestand rond, en tegeldiensten als OSM/CARTO
// staan dat gebruik niet zomaar toe.

function kleurVoor(aantal) {
  for (let i = 0; i < BREAKS.length; i++) {
    if (aantal <= BREAKS[i]) return KLEUREN[i];
  }
  return KLEUREN[KLEUREN.length - 1];
}

function labelVoor(feature) {
  const naam = NAAM_PROPERTY ? (feature.properties[NAAM_PROPERTY] || '(onbekend)') : '(onbekend)';
  return `<strong>${naam}</strong><br>${feature.properties.aantal || 0} initiatieven`;
}

const laag = L.geoJSON(GEOJSON, {
  style: (feature) => ({
    fillColor: kleurVoor(feature.properties.aantal || 0),
    fillOpacity: 0.8,
    color: '#666',
    weight: 1,
  }),
  onEachFeature: (feature, layer) => {
    if (GEBRUIK_CLICK_POPUP) {
      // Klikken i.p.v. hoveren: bij duizenden vlakken tegelijk werd hoveren
      // (continu mousemove volgen) merkbaar hakkelig in de browser.
      layer.bindPopup(labelVoor(feature));
      layer.on('click', () => layer.setStyle({ weight: 3, color: '#000' }));
      layer.on('popupclose', () => layer.setStyle({ weight: 1, color: '#666' }));
    } else {
      layer.bindTooltip(labelVoor(feature), { sticky: true });
      layer.on('mouseover', () => layer.setStyle({ weight: 3, color: '#000' }));
      layer.on('mouseout', () => layer.setStyle({ weight: 1, color: '#666' }));
    }
  },
}).addTo(map);

function pasKaartAanOpInhoud() {
  map.invalidateSize();
  map.fitBounds(laag.getBounds());
}
// Als het tabblad bij het laden nog niet actief/zichtbaar was, heeft de
// kaartcontainer soms nog geen (goede) grootte, en dan blijft de kaart op
// wereldzoom hangen. Dit blijft daarom gewoon elke animatieframe opnieuw
// proberen tot het wel lukt.
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
  div.innerHTML = html;
  return div;
};
legenda.addTo(map);
</script>
</body>
</html>
"""


def bouw_html(features, breaks, kleuren, titel, naam_property, output_file, gebruik_click_popup):
    html = HTML_TEMPLATE
    html = html.replace("__GEOJSON__", json.dumps({"type": "FeatureCollection", "features": features}, ensure_ascii=False, separators=(",", ":")))
    html = html.replace("__BREAKS__", json.dumps(breaks))
    html = html.replace("__KLEUREN__", json.dumps(kleuren))
    html = html.replace("__TITEL__", titel)
    html = html.replace("__NAAM_PROPERTY__", json.dumps(naam_property))
    html = html.replace("__GEBRUIK_CLICK_POPUP__", json.dumps(gebruik_click_popup))
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(html)
    import os
    grootte_mb = os.path.getsize(output_file) / (1024 * 1024)
    print(f"Klaar! {output_file} ({grootte_mb:.1f} MB) - open in je browser.")


# ----------------------------------------------------------------------------
# Postcodekaart met twee lagen: PC4-vlakken (uitgezoomd) en puntjes per
# exacte postcode (ingezoomd, vanaf ZOOM_DREMPEL_PUNTEN).
# ----------------------------------------------------------------------------

HTML_TEMPLATE_POSTCODE = """<!DOCTYPE html>
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
const GEOJSON_VLAKKEN = __GEOJSON_VLAKKEN__;
const BREAKS_VLAKKEN = __BREAKS_VLAKKEN__;
const KLEUREN_VLAKKEN = __KLEUREN_VLAKKEN__;
const PUNTEN = __PUNTEN__;  // [[lat, lon, postcode6, aantal], ...]
const BREAKS_PUNTEN = __BREAKS_PUNTEN__;
const KLEUREN_PUNTEN = __KLEUREN_PUNTEN__;
const ZOOM_DREMPEL = __ZOOM_DREMPEL__;

const map = L.map('map', { preferCanvas: true }).setView([52.1, 5.3], 7);
// Ook hier bewust geen achtergrondkaart, zelfde reden als bij de
// gemeentekaart (zie de toelichting bovenin het script).

function kleurVoor(aantal, breaks, kleuren) {
  for (let i = 0; i < breaks.length; i++) {
    if (aantal <= breaks[i]) return kleuren[i];
  }
  return kleuren[kleuren.length - 1];
}

// --- Laag 1: PC4-vlakken, zichtbaar bij uitgezoomd ----------------------
const vlakkenLaag = L.geoJSON(GEOJSON_VLAKKEN, {
  style: (feature) => ({
    fillColor: kleurVoor(feature.properties.aantal || 0, BREAKS_VLAKKEN, KLEUREN_VLAKKEN),
    fillOpacity: 0.8,
    color: '#999',
    weight: 1,
  }),
  onEachFeature: (feature, layer) => {
    layer.bindTooltip(`<strong>${feature.properties.postcode}</strong> (postcodegebied)<br>${feature.properties.aantal || 0} initiatieven`, { sticky: true });
    layer.on('mouseover', () => layer.setStyle({ weight: 3, color: '#000' }));
    layer.on('mouseout', () => layer.setStyle({ weight: 1, color: '#999' }));
  },
});

// --- Laag 2: puntjes per exacte postcode, zichtbaar bij inzoomen --------
function straalVoor(aantal) {
  return 4 + Math.sqrt(Math.max(aantal, 0)) * 2.2;
}
const puntenLaag = L.layerGroup(
  PUNTEN.map(([lat, lon, postcode, aantal]) => {
    const marker = L.circleMarker([lat, lon], {
      radius: straalVoor(aantal),
      fillColor: kleurVoor(aantal, BREAKS_PUNTEN, KLEUREN_PUNTEN),
      color: '#555',
      weight: 1,
      fillOpacity: 0.9,
    });
    marker.bindTooltip(`<strong>${postcode}</strong><br>${aantal} initiatieven`, { sticky: true });
    return marker;
  })
);

// --- Wisselen tussen de twee lagen op basis van zoomniveau ---------------
const legenda = L.control({ position: 'bottomright' });
legenda.onAdd = () => {
  const div = L.DomUtil.create('div', 'legenda');
  legenda._div = div;
  return div;
};
legenda.addTo(map);

function legendaHtml(titel, breaks, kleuren, subtekst) {
  let html = `<strong>${titel}</strong><br>`;
  let vorige = 0;
  breaks.forEach((b, i) => {
    html += `<span style="background:${kleuren[i]}"></span>${vorige}${vorige === b ? '' : ' - ' + b}<br>`;
    vorige = b + 1;
  });
  html += `<span style="background:${kleuren[kleuren.length - 1]}"></span>${vorige}+`;
  if (subtekst) html += `<br><small>${subtekst}</small>`;
  return html;
}

function bijwerkenLagen() {
  const ingezoomd = map.getZoom() >= ZOOM_DREMPEL;
  if (ingezoomd) {
    if (!map.hasLayer(puntenLaag)) map.addLayer(puntenLaag);
    if (map.hasLayer(vlakkenLaag)) map.removeLayer(vlakkenLaag);
    legenda._div.innerHTML = legendaHtml('Aantal per exacte postcode', BREAKS_PUNTEN, KLEUREN_PUNTEN, 'Zoom uit voor de vlakken per postcodegebied (PC4).');
  } else {
    if (!map.hasLayer(vlakkenLaag)) map.addLayer(vlakkenLaag);
    if (map.hasLayer(puntenLaag)) map.removeLayer(puntenLaag);
    legenda._div.innerHTML = legendaHtml('Aantal per postcodegebied (PC4)', BREAKS_VLAKKEN, KLEUREN_VLAKKEN, `Zoom in (vanaf niveau ${ZOOM_DREMPEL}) voor losse postcodes.`);
  }
}
map.on('zoomend', bijwerkenLagen);

function pasKaartAanOpInhoud() {
  map.invalidateSize();
  map.fitBounds(vlakkenLaag.getBounds());
  bijwerkenLagen();
}
function wachtTotZichtbaarEnFit(pogingenOver) {
  // Zelfde trucje als bij de gemeentekaart: als het tabblad nog niet
  // zichtbaar was bij het laden, blijft dit gewoon proberen tot de kaart
  // wel een fatsoenlijke grootte heeft.
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
</script>
</body>
</html>
"""


def bouw_postcode_html(vlakken_features, breaks_vlakken, kleuren_vlakken,
                        punten, breaks_punten, kleuren_punten,
                        titel, output_file):
    html = HTML_TEMPLATE_POSTCODE
    html = html.replace("__GEOJSON_VLAKKEN__", json.dumps({"type": "FeatureCollection", "features": vlakken_features}, ensure_ascii=False, separators=(",", ":")))
    html = html.replace("__BREAKS_VLAKKEN__", json.dumps(breaks_vlakken))
    html = html.replace("__KLEUREN_VLAKKEN__", json.dumps(kleuren_vlakken))
    html = html.replace("__PUNTEN__", json.dumps(punten, ensure_ascii=False, separators=(",", ":")))
    html = html.replace("__BREAKS_PUNTEN__", json.dumps(breaks_punten))
    html = html.replace("__KLEUREN_PUNTEN__", json.dumps(kleuren_punten))
    html = html.replace("__ZOOM_DREMPEL__", json.dumps(ZOOM_DREMPEL_PUNTEN))
    html = html.replace("__TITEL__", titel)
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(html)
    import os
    grootte_mb = os.path.getsize(output_file) / (1024 * 1024)
    print(f"Klaar! {output_file} ({grootte_mb:.1f} MB) - open in je browser.")


# ----------------------------------------------------------------------------
# GEMEENTEKAART
# ----------------------------------------------------------------------------

def normaliseer_plaatsnaam(naam: str) -> str:
    """Zet plaatsnamen met wisselende hoofdletters/accenten (bijv.
    'mariËnheem' tegenover 'mariënheem') gelijk, anders telt het
    script ze per ongeluk als twee losse plaatsen."""
    return unicodedata.normalize("NFC", str(naam).strip().lower())


def zonder_diakrieten(naam: str) -> str:
    ontleed = unicodedata.normalize("NFD", naam)
    return "".join(c for c in ontleed if unicodedata.category(c) != "Mn")


def losjes(naam: str) -> str:
    """Voor als je twee namen wilt vergelijken zonder dat koppeltekens,
    dubbele spaties of een weggevallen apostrof voor het voorvoegsel
    ertussen komen (dus "'s gravenzande" == "'s-gravenzande" ==
    "s-gravenzande")."""
    zonder_apostrof = re.sub(r"^['’]\s*", "", zonder_diakrieten(naam))
    return re.sub(r"[-\s]+", " ", zonder_apostrof).strip()


def laad_json_cache(pad_str: str) -> dict:
    pad = Path(pad_str)
    if pad.exists():
        with pad.open(encoding="utf-8") as f:
            return json.load(f)
    return {}


def bewaar_json_cache(pad_str: str, data):
    with open(pad_str, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)


# Deze spreektaal-namen bestaan niet (of niet zo) als officiële BAG-
# woonplaatsnaam - de BAG (Basisregistratie Adressen en Gebouwen, zie
# https://www.kadaster.nl/zakelijk/registraties/basisregistraties/bag) is
# waar PDOK zijn woonplaatsnamen vandaan haalt. Zonder deze lijst zou
# bijvoorbeeld "Den Haag" niet gevonden worden, want de officiële naam is
# 's-Gravenhage. Sleutel is de losjes-genormaliseerde vorm.
PLAATSNAAM_ALIASSEN = {
    "den haag": "'s-Gravenhage",
    "denhaag": "'s-Gravenhage",
    "den bosch": "'s-Hertogenbosch",
    "denbosch": "'s-Hertogenbosch",
}


# Plaatsnamen krijgen soms een provincie-afkorting achterop (bijv.
# "Etten Gld", "Rijswijk ZH") om ze te onderscheiden van een gelijknamige
# plaats elders in het land. PDOK geeft in de zoekresultaten zelf ook de
# provincie van elke kandidaat mee (het veld provincieafkorting), dus dit
# is geen slag in de lucht: alleen als de afkorting in het bestand
# overeenkomt met de provincie van een gevonden kandidaat, wordt het als
# match geaccepteerd.
PROVINCIE_AFKORTINGEN = {
    "dr": "DR", "fl": "FL", "fr": "FR", "fry": "FR", "gld": "GD", "gd": "GD",
    "gr": "GR", "gn": "GR", "lb": "LB", "li": "LB", "ov": "OV", "ut": "UT",
    "zl": "ZL", "ze": "ZL", "nb": "NB", "nh": "NH", "zh": "ZH",
}


def splits_provincie_suffix(plaatsnaam_losjes: str):
    """Kijkt of het laatste woord van de plaatsnaam een provincie-
    afkorting is die we herkennen; zo ja, dan komt (basisnaam,
    PDOK-provinciecode) terug, anders None."""
    delen = plaatsnaam_losjes.rsplit(" ", 1)
    if len(delen) == 2 and delen[1] in PROVINCIE_AFKORTINGEN:
        return delen[0], PROVINCIE_AFKORTINGEN[delen[1]]
    return None


def zoek_gemeente_voor_plaatsnaam(plaatsnaam: str):
    """Zoekt bij de PDOK Locatieserver de gemeente bij een plaatsnaam op.
    Geeft (gemeentecode, gemeentenaam) terug, of None als er geen
    exacte match gevonden werd. Ik laat het liever op None uitkomen dan
    dat ik een dubieuze gok als match aanmerk - een paar plaatsen zelf
    opzoeken is minder erg dan initiatieven bij de verkeerde gemeente
    tellen."""
    doel = normaliseer_plaatsnaam(plaatsnaam)
    doel_losjes = losjes(doel)
    alias = PLAATSNAAM_ALIASSEN.get(doel_losjes)
    provincie_split = splits_provincie_suffix(doel_losjes)

    zoektermen = [plaatsnaam, zonder_diakrieten(plaatsnaam)]
    if alias:
        zoektermen.insert(0, alias)
    if provincie_split:
        zoektermen.insert(0, provincie_split[0])

    for zoekterm in zoektermen:
        try:
            resp = requests.get(
                PDOK_LOCATIESERVER_URL,
                params={"q": zoekterm, "fq": "type:woonplaats", "rows": 5},
                timeout=15,
            )
            resp.raise_for_status()
            docs = resp.json().get("response", {}).get("docs", [])
        except (requests.RequestException, KeyError, ValueError):
            docs = []
        doel_losjes_effectief = losjes(normaliseer_plaatsnaam(alias)) if alias else doel_losjes
        for doc in docs:
            kandidaat = normaliseer_plaatsnaam(doc.get("woonplaatsnaam", ""))
            kandidaat_losjes = losjes(kandidaat)
            if kandidaat == doel or kandidaat_losjes == doel_losjes or kandidaat_losjes == doel_losjes_effectief:
                return (doc.get("gemeentecode"), doc.get("gemeentenaam"))
            if provincie_split and kandidaat_losjes == provincie_split[0] and doc.get("provincieafkorting") == provincie_split[1]:
                return (doc.get("gemeentecode"), doc.get("gemeentenaam"))
    return None


def geocodeer_plaatsnamen(plaatsnamen: list, cache: dict) -> dict:
    nog_te_doen = [p for p in plaatsnamen if p not in cache]
    print(f"{len(plaatsnamen)} unieke plaatsnamen, {len(nog_te_doen)} nog niet in cache - opzoeken bij PDOK...")
    if nog_te_doen:
        t0 = time.time()
        klaar = 0
        with ThreadPoolExecutor(max_workers=GEOCODE_WORKERS) as executor:
            toekomsten = {executor.submit(zoek_gemeente_voor_plaatsnaam, p): p for p in nog_te_doen}
            for toekomst in as_completed(toekomsten):
                p = toekomsten[toekomst]
                resultaat = toekomst.result()
                cache[p] = list(resultaat) if resultaat else None
                klaar += 1
                if klaar % 250 == 0 or klaar == len(nog_te_doen):
                    verstreken = time.time() - t0
                    print(f"  {klaar}/{len(nog_te_doen)} gedaan ({verstreken:.0f}s, ~{klaar / max(verstreken,0.01):.1f}/s)")
        bewaar_json_cache(PLAATSNAAM_CACHE_FILE, cache)
        print(f"Opzoeken klaar in {time.time() - t0:.0f}s.")
    return {p: tuple(cache[p]) for p in plaatsnamen if cache.get(p) is not None}


def haal_bewoond_gebied_op():
    """Bouwt een masker van 'bewoond land' door alle PC4-postcodegebieden
    samen te voegen tot één vlak (postcodes bestaan immers niet op open
    water). Wordt lokaal weggeschreven, want alle ~4000 postcodevlakken
    ophalen bij PDOK duurt 30+ pagina's en dat hoeft niet elke run
    opnieuw."""
    cache_pad = Path(BEWOOND_GEBIED_CACHE_FILE)
    if cache_pad.exists():
        print(f"Bewoond-gebied-masker uit cache geladen ({BEWOOND_GEBIED_CACHE_FILE}).")
        with cache_pad.open(encoding="utf-8") as f:
            return shape(json.load(f))

    print("Bewoond-gebied-masker opbouwen (unie van alle PC4-postcodegebieden, zodat de gemeentekaart geen water "
          "meekleurt) - dit haalt eenmalig alle postcodevlakken op bij PDOK, kan een minuutje duren...")
    ruwe_features = haal_ogc_features(POSTCODE4_ITEMS_URL, "postcodegrenzen (voor watermasker)")
    if not ruwe_features:
        print("Let op: kon geen postcodevlakken ophalen, dus de gemeentekaart kleurt het water dan gewoon mee.")
        return Polygon()  # lege geometrie; intersection() geeft dan het ongeknipte vlak terug, zie de aanroep verderop

    # Alleen de nieuwste jaargang per postcode gebruiken - PDOK heeft ook
    # hier meerdere jaargangen per postcode, en die laten overlappen zou de
    # unie alleen maar onnodig trager maken.
    nieuwste = {}
    for feature in ruwe_features:
        pc = str(feature["properties"].get("postcode", "")).zfill(4)
        jaar = feature["properties"].get("jaarcode") or 0
        huidige = nieuwste.get(pc)
        if huidige is None or jaar > (huidige["properties"].get("jaarcode") or 0):
            nieuwste[pc] = feature

    t0 = time.time()
    vereenvoudigd = [
        shape(f["geometry"]).simplify(BEWOOND_GEBIED_SIMPLIFY_TOLERANTIE, preserve_topology=True)
        for f in nieuwste.values()
    ]
    bewoond = unary_union(vereenvoudigd)
    print(f"  {len(vereenvoudigd)} postcodevlakken samengevoegd in {time.time() - t0:.1f}s.")

    with cache_pad.open("w", encoding="utf-8") as f:
        json.dump(mapping(bewoond), f)
    return bewoond


def haal_gemeentegrenzen_op():
    features = haal_ogc_features(GEMEENTE_ITEMS_URL, "gemeentegrenzen")
    if not features:
        raise SystemExit(
            "Kon geen gemeentegrenzen ophalen bij PDOK. Check je internetverbinding, of "
            "download handmatig een GeoJSON via "
            "https://www.pdok.nl/introductie/-/article/bestuurlijke-gebieden "
            "en pas dit script aan om dat lokale bestand in te lezen."
        )
    print(f"Gemeentegrenzen opgehaald: {len(features)} gemeenten.")
    return features


def laad_aantallen_per_plaatsnaam() -> dict:
    """Leest INPUT_FILE_GEMEENTE in (plaatsnaam + aantal) en telt per
    genormaliseerde plaatsnaam op. Dat laatste is nodig omdat dezelfde
    plaats soms met net andere hoofdletters/accenten twee keer in het
    bestand voorkomt (bijv. 'mariËnheem' en 'mariënheem') - zonder
    normaliseren zou dat als twee aparte plaatsen geteld worden."""
    df = pd.read_excel(INPUT_FILE_GEMEENTE)
    df.columns = [str(c).strip() for c in df.columns]
    ontbrekend = [c for c in COLUMN_MAP_GEMEENTE.values() if c not in df.columns]
    if ontbrekend:
        raise SystemExit(
            f"Kolommen niet gevonden in {INPUT_FILE_GEMEENTE}: {ontbrekend}\n"
            f"Beschikbaar: {list(df.columns)}\n"
            f"Pas COLUMN_MAP_GEMEENTE bovenin het script aan."
        )
    print(f"Ingelezen: {len(df)} rijen uit {INPUT_FILE_GEMEENTE}")

    leeg = df[COLUMN_MAP_GEMEENTE["plaatsnaam"]].isna().sum()
    if leeg:
        weggevallen = int(df.loc[df[COLUMN_MAP_GEMEENTE["plaatsnaam"]].isna(), COLUMN_MAP_GEMEENTE["aantal"]].sum())
        print(f"Let op: {leeg} rij(en) zonder plaatsnaam ({weggevallen} initiatieven) worden overgeslagen "
              f"- die zijn niet aan een plaats te koppelen.")
    df = df.dropna(subset=[COLUMN_MAP_GEMEENTE["plaatsnaam"]]).copy()

    df["_plaats_genorm"] = df[COLUMN_MAP_GEMEENTE["plaatsnaam"]].apply(normaliseer_plaatsnaam)
    # Ook de originele schrijfwijze (met accenten) bewaren per genormaliseerde
    # naam, want die gebruik ik om bij PDOK te zoeken - dat zoekt met accenten
    # er nog in vaak net iets preciezer.
    origineel_per_norm = df.drop_duplicates("_plaats_genorm").set_index("_plaats_genorm")[COLUMN_MAP_GEMEENTE["plaatsnaam"]].to_dict()

    if AANTAL_IS_AL_EEN_TELLING_GEMEENTE:
        per_plaats = df.groupby("_plaats_genorm")[COLUMN_MAP_GEMEENTE["aantal"]].sum()
    else:
        per_plaats = df.groupby("_plaats_genorm")[COLUMN_MAP_GEMEENTE["aantal"]].nunique()

    return {origineel_per_norm[k]: int(v) for k, v in per_plaats.to_dict().items()}


def maak_gemeentekaart():
    print("\n=== Gemeentekaart ===")
    aantallen_per_plaats = laad_aantallen_per_plaatsnaam()

    cache = laad_json_cache(PLAATSNAAM_CACHE_FILE)
    gemeente_per_plaats = geocodeer_plaatsnamen(list(aantallen_per_plaats.keys()), cache)

    aantallen = {}  # gemeentecode -> totaal aantal
    plaatsen_niet_gevonden = []
    for plaats, aantal in aantallen_per_plaats.items():
        gevonden = gemeente_per_plaats.get(plaats)
        if gevonden is None:
            plaatsen_niet_gevonden.append({"plaatsnaam": plaats, "aantal": aantal})
            continue
        gm_code, gm_naam = gevonden
        aantallen[normaliseer(gm_code)] = aantallen.get(normaliseer(gm_code), 0) + aantal

    if plaatsen_niet_gevonden:
        pd.DataFrame(plaatsen_niet_gevonden).sort_values("aantal", ascending=False).to_excel(
            "plaatsnamen_niet_gevonden.xlsx", index=False
        )
        totaal_gemist = sum(p["aantal"] for p in plaatsen_niet_gevonden)
        print(f"Let op: {len(plaatsen_niet_gevonden)} plaatsnamen ({totaal_gemist} initiatieven) konden niet aan "
              f"een Nederlandse gemeente gekoppeld worden (bijv. Aruba/Curaçao/Sint Maarten/Caribisch Nederland - "
              f"die vallen buiten dit PDOK-bestand, of een niet-herkende schrijfwijze). Zie plaatsnamen_niet_gevonden.xlsx.")

    ruwe_features = haal_gemeentegrenzen_op()
    code_kandidaten = ["identificatie", "gemeentecode", "statcode", "code", "GM_CODE", "gemeentecodenaam"]
    naam_kandidaten = ["naam", "gemeentenaam", "statnaam", "GM_NAAM"]
    koppel_property = bepaal_property(ruwe_features, code_kandidaten, "koppel")
    if koppel_property is None:
        raise SystemExit(
            "Kon geen koppel-veld vinden voor gemeentecode. Pas de "
            "kandidatenlijst (code_kandidaten) in het script aan naar het "
            "juiste veld uit de bovenstaande lijst met beschikbare velden."
        )
    naam_property = bepaal_property(ruwe_features, naam_kandidaten, "naam (voor popup)")

    # Het veld 'identificatie' bij Kadaster-data begint soms met "GM" (net
    # als de gm_code die ik zelf gebruik), maar soms is het gewoon numeriek
    # ("0363" i.p.v. "GM0363"). Vandaar dat ik hieronder allebei probeer.
    def kandidaat_codes(ruwe_code: str) -> set:
        c = normaliseer(ruwe_code)
        varianten = {c}
        if c.startswith("GM"):
            varianten.add(c[2:])
        else:
            varianten.add("GM" + c)
        return varianten

    bewoond_gebied = haal_bewoond_gebied_op()

    print(f"Geometrie vereenvoudigen ({len(ruwe_features)} gemeenten)...")
    t0 = time.time()
    gebruikte_codes = set()
    features = []
    for feature in ruwe_features:
        ruwe_code = feature["properties"].get(koppel_property, "")
        gevonden_aantal = None
        for variant in kandidaat_codes(ruwe_code):
            if variant in aantallen:
                gevonden_aantal = aantallen[variant]
                gebruikte_codes.add(variant)
                break

        # Alleen de velden bewaren die we echt gebruiken (koppel + naam +
        # aantal). De rest van wat PDOK meestuurt gaat overboord.
        nieuwe_properties = {koppel_property: ruwe_code}
        if naam_property:
            nieuwe_properties[naam_property] = feature["properties"].get(naam_property)
        nieuwe_properties["aantal"] = int(gevonden_aantal) if gevonden_aantal is not None else 0

        # Water eruit knippen (IJsselmeer, Waddenzee, ...) voordat de
        # geometrie vereenvoudigd wordt - anders zit je tijd te steken in het
        # vereenvoudigen van water dat er toch weer af gaat.
        #
        # Bij gemeenten met een grote/ingewikkelde waterjurisdictie (kwam ik
        # o.a. bij Lelystad en Súdwest-Fryslân tegen) raakt de rand van het
        # gemeentevlak de rand van het bewoond-gebied-masker op meerdere
        # plekken tegelijk. shapely's intersection() geeft dan niet netjes
        # een Polygon/MultiPolygon terug, maar een GeometryCollection: een
        # mengeling van het eigenlijke overlappende vlak met hier en daar
        # een los punt of lijnstukje van randen die elkaar precies raken.
        # Vandaar dat hieronder expliciet alleen de vlak-onderdelen eruit
        # gehaald worden - zonder die stap blijft zo'n gemeente stilletjes
        # ongeknipt (en dus met water erin).
        geometrie = feature["geometry"]
        if not bewoond_gebied.is_empty:
            try:
                geknipt = shape(geometrie).intersection(bewoond_gebied)
                if geknipt.geom_type == "GeometryCollection":
                    vlakdelen = [g for g in geknipt.geoms if g.geom_type in ("Polygon", "MultiPolygon")]
                    geknipt = unary_union(vlakdelen) if vlakdelen else geknipt
                if not geknipt.is_empty and geknipt.geom_type in ("Polygon", "MultiPolygon"):
                    geometrie = mapping(geknipt)
                else:
                    print(f"  Let op: knippen van '{nieuwe_properties.get(naam_property)}' gaf geen bruikbaar vlak "
                          f"(type: {geknipt.geom_type}) - ongeknipt gebruikt.")
            except Exception as e:
                print(f"  Let op: knippen van '{nieuwe_properties.get(naam_property)}' mislukte ({e}) - ongeknipt gebruikt.")

        features.append({
            "type": "Feature",
            "properties": nieuwe_properties,
            "geometry": vereenvoudig_geometrie(geometrie, SIMPLIFY_TOLERANTIE_GEMEENTE),
        })
    print(f"  klaar in {time.time() - t0:.1f}s")

    niet_gevonden = [
        {"gm_code_in_bestand": code, "aantal": aantal}
        for code, aantal in aantallen.items()
        if code not in gebruikte_codes
    ]
    if niet_gevonden:
        pd.DataFrame(niet_gevonden).to_excel(NIET_GEVONDEN_GEMEENTE, index=False)
        print(f"Let op: {len(niet_gevonden)} gm_codes konden niet aan een Nederlandse gemeente "
              f"gekoppeld worden (bijv. buiten het Koninkrijk). Zie {NIET_GEVONDEN_GEMEENTE}.")

    pd.DataFrame([
        {"gm_code": f["properties"][koppel_property], "naam": f["properties"].get(naam_property), "aantal": f["properties"]["aantal"]}
        for f in features
    ]).sort_values("aantal", ascending=False).to_excel(AANTALLEN_GEMEENTE, index=False)
    print(f"Volledige telling weggeschreven naar {AANTALLEN_GEMEENTE} - handig om te vergelijken met je eigen Excel.")

    breaks = bereken_breaks([f["properties"]["aantal"] for f in features], len(KLEUREN_GEMEENTE))
    bouw_html(
        features, breaks, KLEUREN_GEMEENTE,
        "Aantal initiatieven per gemeente",
        naam_property, OUTPUT_GEMEENTE,
        gebruik_click_popup=False,
    )


# ----------------------------------------------------------------------------
# POSTCODEKAART
# ----------------------------------------------------------------------------

def haal_postcodegrenzen_op():
    features = haal_ogc_features(POSTCODE4_ITEMS_URL, "postcodegrenzen")
    if not features:
        raise SystemExit(
            "Kon geen postcodegrenzen ophalen bij PDOK. Check je internetverbinding, of "
            "download handmatig een GeoJSON via "
            "https://www.pdok.nl/introductie/-/article/cbs-postcode4 "
            "en pas dit script aan om dat lokale bestand in te lezen."
        )
    print(f"Postcodegrenzen opgehaald: {len(features)} vlakken (incl. oudere jaargangen).")
    return features


def laad_aantallen_per_exacte_postcode() -> dict:
    df = pd.read_excel(INPUT_FILE_POSTCODE)
    df.columns = [str(c).strip() for c in df.columns]
    ontbrekend = [c for c in COLUMN_MAP_POSTCODE.values() if c not in df.columns]
    if ontbrekend:
        raise SystemExit(
            f"Kolommen niet gevonden in {INPUT_FILE_POSTCODE}: {ontbrekend}\n"
            f"Beschikbaar: {list(df.columns)}\n"
            f"Pas COLUMN_MAP_POSTCODE bovenin het script aan."
        )
    print(f"Ingelezen: {len(df)} rijen uit {INPUT_FILE_POSTCODE}")

    df["_postcode"] = df[COLUMN_MAP_POSTCODE["postcode"]].apply(naar_geldige_postcode6)
    ongeldig = (df["_postcode"] == "").sum()
    if ongeldig:
        print(f"Let op: {ongeldig} rijen hebben geen herkenbare 6-tekens postcode en worden overgeslagen.")
    df = df[df["_postcode"] != ""]

    if AANTAL_IS_AL_EEN_TELLING_POSTCODE:
        per_postcode = df.groupby("_postcode")[COLUMN_MAP_POSTCODE["aantal"]].sum()
    else:
        per_postcode = df.groupby("_postcode")[COLUMN_MAP_POSTCODE["aantal"]].nunique()
    return per_postcode.astype(int).to_dict()


def maak_postcodekaart():
    print("\n=== Postcodekaart (vlakken + puntjes per exacte postcode) ===")
    aantallen_exact = laad_aantallen_per_exacte_postcode()

    cache = laad_json_cache(POSTCODE_COORDS_CACHE_FILE)
    coords = geocodeer_postcodes6(list(aantallen_exact.keys()), cache)

    niet_gegeocodeerd = [pc for pc in aantallen_exact if pc not in coords]
    if niet_gegeocodeerd:
        totaal_gemist = sum(aantallen_exact[pc] for pc in niet_gegeocodeerd)
        pd.DataFrame([
            {"postcode": pc, "aantal": aantallen_exact[pc]} for pc in niet_gegeocodeerd
        ]).sort_values("aantal", ascending=False).to_excel(NIET_GEVONDEN_POSTCODE, index=False)
        print(f"Let op: {len(niet_gegeocodeerd)} exacte postcodes ({totaal_gemist} initiatieven) kon PDOK niet "
              f"vinden (verouderd/niet-bestaand). Zie {NIET_GEVONDEN_POSTCODE}.")

    # Puntenlaag - alleen de postcodes waarvoor ook echt coördinaten
    # gevonden zijn.
    punten = [
        [coords[pc][0], coords[pc][1], pc, aantallen_exact[pc]]
        for pc in aantallen_exact if pc in coords
    ]
    breaks_punten = bereken_breaks([p[3] for p in punten], len(KLEUREN_POSTCODE))
    print(f"{len(punten)} exacte postcodes op de kaart (puntenlaag).")

    # Vlakkenlaag - dezelfde brontelling, maar dan opgeteld tot PC4, zodat de
    # puntjes binnen een vlak altijd precies optellen tot dat vlak.
    aantallen = {}
    for pc, aantal in aantallen_exact.items():
        pc4 = pc[:4]
        aantallen[pc4] = aantallen.get(pc4, 0) + aantal

    ruwe_features = haal_postcodegrenzen_op()
    pc_kandidaten = ["postcode4", "pc4", "postcode", "statcode"]
    # Het jaargang-veld heet bij PDOK 'jaarcode'. Zonder hierop te filteren
    # blijven alle ~10 jaargangen per postcode staan, i.p.v. alleen de
    # nieuwste - vandaar dit filter.
    jaar_kandidaten = ["jaarcode", "jaar", "year", "statjaar", "einddatum"]
    pc_property = bepaal_property(ruwe_features, pc_kandidaten, "postcode")
    if pc_property is None:
        raise SystemExit(
            "Kon geen postcode-veld vinden. Pas de kandidatenlijst "
            "(pc_kandidaten) in het script aan."
        )
    jaar_property = bepaal_property(ruwe_features, jaar_kandidaten, "jaar (voor meest recente editie)")

    # Eerst per postcode alleen de nieuwste editie bewaren, nog vóór het
    # vereenvoudigen van de geometrie - anders steek je rekentijd in oude
    # jaargangen die toch weggegooid worden.
    if jaar_property:
        nieuwste = {}
        for feature in ruwe_features:
            pc = str(feature["properties"].get(pc_property, "")).zfill(4)
            jaar = feature["properties"].get(jaar_property) or 0
            huidige = nieuwste.get(pc)
            if huidige is None or jaar > (huidige["properties"].get(jaar_property) or 0):
                nieuwste[pc] = feature
        gekozen_features = list(nieuwste.values())
        print(f"Na filteren op meest recente editie: {len(gekozen_features)} postcodegebieden "
              f"(was {len(ruwe_features)} incl. oude jaargangen).")
    else:
        # Als er geen jaar-veld te vinden is, val ik terug op "laatst geziene
        # editie wint" - niet perfect, maar dan tellen we in elk geval geen
        # jaargangen dubbel mee.
        print("Let op: geen jaargang-veld gevonden, val terug op 'laatst geziene editie per postcode'.")
        gekozen_features = list({
            str(f["properties"].get(pc_property, "")).zfill(4): f for f in ruwe_features
        }.values())

    # Nu de properties opschonen (PDOK stuurt zo'n 140 CBS-statistieken mee
    # per postcodevlak waar deze kaart niks mee doet - die gaan eraf) en de
    # geometrie vereenvoudigen.
    print(f"Geometrie vereenvoudigen ({len(gekozen_features)} postcodegebieden)...")
    t0 = time.time()
    gebruikte_pcs = set()
    features = []
    for feature in gekozen_features:
        pc = str(feature["properties"].get(pc_property, "")).zfill(4)
        aantal = aantallen.get(pc)
        if aantal is not None:
            gebruikte_pcs.add(pc)
        features.append({
            "type": "Feature",
            "properties": {
                pc_property: pc,
                "aantal": int(aantal) if aantal is not None else 0,
            },
            "geometry": vereenvoudig_geometrie(feature["geometry"], SIMPLIFY_TOLERANTIE_POSTCODE),
        })
    print(f"  klaar in {time.time() - t0:.1f}s")

    pc4_niet_gevonden = [
        {"pc4_in_bestand": pc, "aantal": aantal}
        for pc, aantal in aantallen.items()
        if pc not in gebruikte_pcs
    ]
    if pc4_niet_gevonden:
        pd.DataFrame(pc4_niet_gevonden).to_excel("postcode4_niet_gevonden.xlsx", index=False)
        print(f"Let op: {len(pc4_niet_gevonden)} PC4-gebieden konden niet aan een vlak gekoppeld worden. "
              f"Zie postcode4_niet_gevonden.xlsx.")

    pd.DataFrame([
        {"pc4": f["properties"][pc_property], "aantal": f["properties"]["aantal"]}
        for f in features
    ]).sort_values("aantal", ascending=False).to_excel(AANTALLEN_POSTCODE, index=False)
    print(f"Volledige telling weggeschreven naar {AANTALLEN_POSTCODE} - handig om te vergelijken met je eigen Excel "
          f"(let op: dit is per PC4, dus de som van alle volledige postcodes die met diezelfde 4 cijfers beginnen).")

    breaks_vlakken = bereken_breaks([f["properties"]["aantal"] for f in features], len(KLEUREN_POSTCODE))
    bouw_postcode_html(
        features, breaks_vlakken, KLEUREN_POSTCODE,
        punten, breaks_punten, KLEUREN_POSTCODE,
        "Aantal initiatieven per postcode", OUTPUT_POSTCODE,
    )


def main():
    maak_gemeentekaart()   # leest zelf INPUT_FILE_GEMEENTE (Locaties_per_gemeente.xlsx)
    maak_postcodekaart()   # leest zelf INPUT_FILE_POSTCODE (Locaties_per_postcode.xlsx)


if __name__ == "__main__":
    main()
