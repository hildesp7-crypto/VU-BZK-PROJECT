"""
03_kaart_bouwen.py
===================
Bouwt één standalone HTML-bestand (geen server nodig — gewoon dubbelklikken
en openen in de browser) met een interactieve kaart van alle burgerinitiatieven,
inclusief filters en een schakelaar tussen Europees Nederland en het Caribisch
gebied (Bonaire, Sint Eustatius, Saba, en evt. Aruba/Curaçao/Sint Maarten als
die in je data voorkomen — de kaart zoomt automatisch in op wat er is).

VOOR GEBRUIK:
  1. Zorg dat 02_geocoderen.py al is gedraaid (er moeten latitude/longitude
     kolommen zijn), of pas INPUT_FILE / COLUMN_MAP aan.
  2. Vul COLUMN_MAP hieronder in met de echte kolomnamen.
  3. python 03_kaart_bouwen.py
  4. Open burgerinitiatieven_kaart.html in je browser.
"""

import json
import re

import pandas as pd

# ----------------------------------------------------------------------------
INPUT_FILE = "Locaties_organisaties.xlsx"
OUTPUT_FILE = "burgerinitiatieven_kaart.html"

COLUMN_MAP = {
    "naam": "Naam organisatie",
    "adres": "Adres",
    "plaats": "Plaats",
    "lat": "latitude",
    "lon": "longitude",
    "kvk": "KvK-nummer",
    "soort_organisatie": "Soort organisatie",
    "organisatietype": "Organisatietype",
    "opgericht": "Opgericht in",
    "werkgebied": "Werkgebied",
    "doelgroepen": "Doelgroepen",  # cel mag meerdere waarden bevatten, gescheiden door , of ;
}

# Vaste optielijsten (zoals opgegeven) — gebruikt voor de filter-checkboxes.
SOORT_ORGANISATIE_OPTIES = [
    "Accommodatie", "Accommodatie - in professioneel beheer",
    "Accommodatie - in zelfbeheer", "Belangen- en Zelforganisatie",
    "Belangenorganisatie", "Bewonersorganisatie", "Brede welzijnsorganisatie",
    "Informele groep", "Kennis instituut", "Koepelorganisatie",
    "Onderwijs- en educatieorganisatie", "Onderwijsinstelling",
    "Sociale onderneming", "Vrijwilligersorganisatie", "Zelforganisatie",
    "Zorginstelling",
]

ORGANISATIETYPE_OPTIES = [
    "BV/NV", "Commanditaire Vennootschap (CV)", "Coöperatie",
    "Eenmanszaak/ZZP", "Informeel/werkgroep", "Kerkgenootschap",
    "Onderneming met rechtspersoonlijkheid", "Onderneming zonder rechtspersoonlijkheid",
    "Organisatie in oprichting", "Overheid", "Publiekrechtelijk", "Stichting",
    "Vereniging", "VOF", "Zorg groep",
]

WERKGEBIED_OPTIES = ["Caribisch", "Internationaal", "Landelijk", "Plaatselijk", "Regionaal"]

DOELGROEPEN_OPTIES = [
    "(Ex-)gedetineerden", "Algemeen", "Buurtbewoners", "Dak- en thuislozen",
    "Gezinnen", "Jongens/mannen", "Jongeren (vanaf 12 jaar)",
    "Jongeren (vanaf 18 jaar)", "Kinderen (tot 12 jaar)", "LHBTI+",
    "Mantelzorgers", "Meisjes/vrouwen", "Mensen die eenzaamheid ervaren",
    "Mensen in armoede", "Mensen met een beperking/patiënt",
    "Mensen met een lichamelijke beperking", "Mensen met een psychische kwetsbaarheid",
    "Mensen met een verstandelijke beperking", "Mensen met schuldenproblematiek",
    "Mensen met verslavingsproblematiek", "Nieuwe Nederlanders", "Ouderen",
    "Ouders/opvoeders", "Vrijwilligers",
]
# ----------------------------------------------------------------------------

GEEN_DATA_WAARDEN = {"", "geen", "nvt", "n.v.t.", "onbekend", "-", "n/a"}


def is_leeg(waarde) -> bool:
    """True als een cel leeg is of een 'geen data'-tekst bevat (bijv. 'Geen', 'nvt')."""
    if pd.isna(waarde):
        return True
    return str(waarde).strip().lower() in GEEN_DATA_WAARDEN


def parse_jaar(waarde):
    """Haalt een jaartal (19xx/20xx) uit de cel, ook als er ruis omheen staat
    (bijv. '2012 (2017 centrum de steen)' -> 2012). Geeft None als er geen
    herkenbaar jaartal in staat."""
    if pd.isna(waarde):
        return None
    match = re.search(r"(19|20)\d{2}", str(waarde))
    return int(match.group()) if match else None


def split_multivalue(cell) -> list:
    if pd.isna(cell):
        return []
    return [v.strip() for v in re.split(r"[;,]", str(cell)) if v.strip()]


def build_records(df: pd.DataFrame) -> list:
    records = []
    for _, row in df.iterrows():
        lat, lon = row.get(COLUMN_MAP["lat"]), row.get(COLUMN_MAP["lon"])
        if pd.isna(lat) or pd.isna(lon):
            continue  # kan niet op de kaart zonder coördinaten
        kvk_val = row.get(COLUMN_MAP["kvk"])
        records.append({
            "naam": str(row.get(COLUMN_MAP["naam"], "") or ""),
            "adres": str(row.get(COLUMN_MAP["adres"], "") or ""),
            "plaats": str(row.get(COLUMN_MAP["plaats"], "") or ""),
            "lat": float(lat),
            "lon": float(lon),
            "heeft_kvk": not is_leeg(kvk_val),
            "kvk": "" if is_leeg(kvk_val) else str(kvk_val),
            "soort": str(row.get(COLUMN_MAP["soort_organisatie"], "") or ""),
            "type": str(row.get(COLUMN_MAP["organisatietype"], "") or ""),
            "opgericht": parse_jaar(row.get(COLUMN_MAP["opgericht"])),
            "werkgebied": str(row.get(COLUMN_MAP["werkgebied"], "") or ""),
            "doelgroepen": split_multivalue(row.get(COLUMN_MAP["doelgroepen"])),
        })
    return records


def checkbox_html(name: str, options: list) -> str:
    items = "\n".join(
        f'<label><input type="checkbox" class="filter-{name}" value="{opt}">{opt}</label>'
        for opt in options
    )
    return items


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="UTF-8">
<title>Kaart burgerinitiatieven</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
<link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css" />
<link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css" />
<style>
  html, body { margin:0; padding:0; height:100%; font-family: system-ui, Arial, sans-serif; }
  #app { display:flex; height:100vh; }
  #sidebar { width: 320px; flex-shrink:0; overflow-y:auto; background:#f7f7f7; border-right:1px solid #ddd; padding:16px; box-sizing:border-box; }
  #sidebar h2 { font-size:16px; margin:0 0 12px; }
  #kaart-wrap { position:relative; flex-grow:1; }
  #map { position:absolute; top:0; left:0; right:0; bottom:0; background:#fff; }
  .filter-group { margin-bottom:16px; }
  .filter-group > strong { display:block; margin-bottom:6px; font-size:13px; color:#333; }
  .filter-group.scroll { max-height:150px; overflow-y:auto; border:1px solid #e0e0e0; padding:6px; background:#fff; border-radius:4px; }
  label { display:block; font-size:13px; margin-bottom:4px; cursor:pointer; }
  input[type=checkbox] { margin-right:6px; }
  #teller { font-size:13px; font-weight:bold; margin-bottom:12px; padding:8px; background:#eef; border-radius:4px; }
  #reset { width:100%; padding:8px; margin-bottom:16px; cursor:pointer; }
  .jaar-row { display:flex; gap:8px; align-items:center; }
  .jaar-row input { width:70px; }
  #view-switch { position:absolute; top:10px; right:10px; z-index:1000; background:#fff; border-radius:6px; box-shadow:0 1px 4px rgba(0,0,0,.3); overflow:hidden; }
  #view-switch button { border:none; background:#fff; padding:8px 14px; cursor:pointer; font-size:13px; }
  #view-switch button.actief { background:#2563eb; color:#fff; }
</style>
</head>
<body>
<div id="app">
  <div id="sidebar">
    <h2>Filters</h2>
    <div id="teller">__ van __ initiatieven zichtbaar</div>
    <button id="reset">Filters resetten</button>

    <div class="filter-group">
      <strong>KvK-nummer</strong>
      <label><input type="checkbox" id="filter-kvk"> Alleen initiatieven met KvK-nummer</label>
    </div>

    <div class="filter-group">
      <strong>Opgericht in (jaar)</strong>
      <div class="jaar-row">
        <input type="number" id="jaar-van" placeholder="van">
        <span>-</span>
        <input type="number" id="jaar-tot" placeholder="tot">
      </div>
    </div>

    <div class="filter-group">
      <strong>Werkgebied</strong>
      <div class="scroll">
        __WERKGEBIED_CHECKBOXES__
      </div>
    </div>

    <div class="filter-group">
      <strong>Soort organisatie</strong>
      <div class="filter-group scroll">
        __SOORT_CHECKBOXES__
      </div>
    </div>

    <div class="filter-group">
      <strong>Organisatietype</strong>
      <div class="filter-group scroll">
        __TYPE_CHECKBOXES__
      </div>
    </div>

    <div class="filter-group">
      <strong>Doelgroepen</strong>
      <div class="filter-group scroll">
        __DOELGROEPEN_CHECKBOXES__
      </div>
    </div>
  </div>
  <div id="kaart-wrap">
    <div id="view-switch">
      <button id="btn-nl" class="actief">Europees Nederland</button>
      <button id="btn-car">Caribisch gebied</button>
    </div>
    <div id="map"></div>
  </div>
</div>

<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>
<script>
const DATA = __DATA_JSON__;

const map = L.map('map', { zoomControl: true });
// Bewust GEEN achtergrondkaart (tegelserver) — na drie providers die
// allemaal om een andere reden niet betrouwbaar bleken (osm.org blokkeert
// dit soort verspreide bestanden, CARTO vraagt alsnog een API-key, en
// PDOK's dekking liep niet ver genoeg door rond Nederland), is een effen
// achtergrond de simpelste, altijd-werkende oplossing: geen enkele
// afhankelijkheid van een externe dienst meer.

// Alles met een negatieve lengtegraad onder -30 ligt in het Caribisch gebied
// (Bonaire, Sint Eustatius, Saba, Aruba, Curaçao, Sint Maarten), de rest is
// Europees Nederland. Bonaire (ver zuidelijk) en Sint Eustatius/Saba (ver
// noordelijker) liggen zelf ruim 600 km uit elkaar, dus we zoomen dynamisch
// in op de punten die er daadwerkelijk zijn i.p.v. een vaste positie.
const VIEW_NL = { center: [52.2, 5.3], zoom: 8 };
const isCaribisch = (item) => item.lon < -30;

function fitOpPunten(punten, fallback) {
  if (!punten.length) {
    if (fallback) map.setView(fallback.center, fallback.zoom);
    return false;
  }
  const bounds = L.latLngBounds(punten.map(p => [p.lat, p.lon]));
  map.fitBounds(bounds, { padding: [50, 50], maxZoom: 12 });
  return true;
}

fitOpPunten(DATA.filter(d => !isCaribisch(d)), VIEW_NL);
// De kaart-container krijgt zijn hoogte pas ná de eerste render (CSS/lettertype
// laden e.d.); Leaflet meet zijn eigen grootte soms te vroeg. Dit forceert een
// hermeting zodat de tegels altijd goed gevuld worden.
setTimeout(() => map.invalidateSize(), 200);

document.getElementById('btn-nl').addEventListener('click', () => {
  fitOpPunten(DATA.filter(d => !isCaribisch(d)), VIEW_NL);
  setActief('btn-nl');
});
document.getElementById('btn-car').addEventListener('click', () => {
  const gevonden = fitOpPunten(DATA.filter(isCaribisch));
  if (!gevonden) alert('Geen initiatieven in het Caribisch gebied gevonden in deze dataset.');
  setActief('btn-car');
});
function setActief(id) {
  document.getElementById('btn-nl').classList.remove('actief');
  document.getElementById('btn-car').classList.remove('actief');
  document.getElementById(id).classList.add('actief');
}

let clusterGroup = L.markerClusterGroup();
map.addLayer(clusterGroup);

function maakPopup(item) {
  const doelgroepen = item.doelgroepen.length ? item.doelgroepen.join(', ') : '-';
  return `<strong>${item.naam}</strong><br>
    ${item.adres ? item.adres + ', ' : ''}${item.plaats}<br>
    <em>${item.soort || '-'}</em> · ${item.type || '-'}<br>
    Werkgebied: ${item.werkgebied || '-'}<br>
    Opgericht: ${item.opgericht || '-'}<br>
    KvK: ${item.heeft_kvk ? item.kvk : 'geen'}<br>
    Doelgroepen: ${doelgroepen}`;
}

function getChecked(selector) {
  return Array.from(document.querySelectorAll(selector + ':checked')).map(el => el.value);
}

function passtFilter(item) {
  if (document.getElementById('filter-kvk').checked && !item.heeft_kvk) return false;

  const jaarVan = document.getElementById('jaar-van').value;
  const jaarTot = document.getElementById('jaar-tot').value;
  if (jaarVan && (!item.opgericht || item.opgericht < parseInt(jaarVan))) return false;
  if (jaarTot && (!item.opgericht || item.opgericht > parseInt(jaarTot))) return false;

  const werkgebied = getChecked('.filter-werkgebied');
  if (werkgebied.length && !werkgebied.includes(item.werkgebied)) return false;

  const soort = getChecked('.filter-soort');
  if (soort.length && !soort.includes(item.soort)) return false;

  const type = getChecked('.filter-type');
  if (type.length && !type.includes(item.type)) return false;

  const doelgroepen = getChecked('.filter-doelgroepen');
  if (doelgroepen.length && !doelgroepen.some(d => item.doelgroepen.includes(d))) return false;

  return true;
}

function herteken() {
  clusterGroup.clearLayers();
  let zichtbaar = 0;
  DATA.forEach(item => {
    if (passtFilter(item)) {
      zichtbaar++;
      const marker = L.marker([item.lat, item.lon]);
      marker.bindPopup(maakPopup(item));
      clusterGroup.addLayer(marker);
    }
  });
  document.getElementById('teller').textContent = `${zichtbaar} van ${DATA.length} initiatieven zichtbaar`;
}

document.querySelectorAll('#sidebar input').forEach(el => el.addEventListener('change', herteken));
document.getElementById('jaar-van').addEventListener('input', herteken);
document.getElementById('jaar-tot').addEventListener('input', herteken);

document.getElementById('reset').addEventListener('click', () => {
  document.querySelectorAll('#sidebar input[type=checkbox]').forEach(el => el.checked = false);
  document.getElementById('jaar-van').value = '';
  document.getElementById('jaar-tot').value = '';
  herteken();
});

herteken();
</script>
</body>
</html>
"""


def main():
    df = pd.read_excel(INPUT_FILE)
    records = build_records(df)
    print(f"{len(records)} van {len(df)} rijen hebben geldige coördinaten en komen op de kaart.")

    html = HTML_TEMPLATE
    html = html.replace("__DATA_JSON__", json.dumps(records, ensure_ascii=False))
    html = html.replace("__WERKGEBIED_CHECKBOXES__", checkbox_html("werkgebied", WERKGEBIED_OPTIES))
    html = html.replace("__SOORT_CHECKBOXES__", checkbox_html("soort", SOORT_ORGANISATIE_OPTIES))
    html = html.replace("__TYPE_CHECKBOXES__", checkbox_html("type", ORGANISATIETYPE_OPTIES))
    html = html.replace("__DOELGROEPEN_CHECKBOXES__", checkbox_html("doelgroepen", DOELGROEPEN_OPTIES))

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Klaar! Open {OUTPUT_FILE} in je browser.")


if __name__ == "__main__":
    main()

