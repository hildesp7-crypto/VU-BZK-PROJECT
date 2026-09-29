"""
01_dedupliceren.py
====================
Ontdubbelt het burgerinitiatieven-bestand in drie stappen:
 
  1. EXACT   - identieke rijen / identiek KvK-nummer -> altijd duplicaat
  2. FUZZY   - lijkt sterk op elkaar (naam + adres) via rapidfuzz -> automatisch duplicaat
  3. AI-CHECK - twijfelgevallen (niet zeker genoeg voor stap 2) worden voorgelegd aan
               de Nebula AI van de VU (OpenAI-compatible API) die JA/NEE antwoordt.
 
Waarom niet gewoon df.drop_duplicates()?
Dat vangt alleen 100% identieke rijen. In de praktijk staan initiatieven vaak
dubbel met kleine verschillen (spatie, hoofdletter, "Stichting X" vs "stichting x",
net iets ander adres). Dit script vangt die gevallen ook op, met een audit-rapport
zodat je kan controleren wat er is samengevoegd.
 
VOOR GEBRUIK:
  1. Vul COLUMN_MAP hieronder in met de echte kolomnamen uit jouw Excel-bestand.
     "adres" en "postcode" mogen allebei op None staan als je die kolommen niet
     hebt (bijv. omdat er om privacyredenen alleen postcode+plaats bewaard is).
  2. Zet je Nebula API-sleutel klaar (zie README.md) of zet USE_AI_CHECK = False
     om zonder AI te draaien (dan doet het script alleen stap 1 + 2).
  3. pip install -r requirements.txt
  4. python 01_dedupliceren.py
"""
 
import os
import re
import sys
 
import pandas as pd
from rapidfuzz import fuzz
 
# ----------------------------------------------------------------------------
# INSTELLINGEN — pas dit aan naar jouw situatie
# ----------------------------------------------------------------------------
INPUT_FILE = "Locaties_organisaties.xlsx"
OUTPUT_FILE = "Locaties_organisaties_ontdubbeld.xlsx"
REPORT_FILE = "ontdubbeling_rapport.xlsx"
 
# Kolomnamen zoals ze in JOUW Excel-bestand heten. Pas de rechterkant aan.
# "adres" en "postcode" mogen allebei None zijn als je die kolommen niet hebt
# (het script dedupliceert dan op naam + postcode + plaats + KvK-nummer).
COLUMN_MAP = {
    "naam": "Relatienaam",
    "adres": None,
    "postcode": "Postcode",
    "plaats": "Plaats",
    "kvk": "KvK",
}
 
# Drempelwaarden voor de fuzzy-vergelijking (0-100, hoger = strenger)
FUZZY_THRESHOLD_AUTO = 92   # score >= dit -> automatisch duplicaat, geen AI nodig
FUZZY_THRESHOLD_AI = 80     # score tussen deze grens en AUTO -> voorleggen aan Nebula AI
                            # score < deze grens -> beschouwd als twee verschillende organisaties
 
USE_AI_CHECK = True         # False = alleen stap 1 (exact) + 2 (fuzzy), geen Nebula nodig
 
# Nebula (VU) instellingen — Open WebUI met OpenAI-compatible API.
# Haal je API-key op via je Nebula-account: Instellingen > Account > API keys.
# Vraag het exacte endpoint na bij het Nebula-team (r.apsan@vu.nl / m.otte@vu.nl)
# als "https://nebula.vu.nl/api" hieronder niet klopt.
NEBULA_BASE_URL = os.environ.get("NEBULA_BASE_URL", "https://nebula.vu.nl/api")
NEBULA_API_KEY = os.environ.get("NEBULA_API_KEY", "VUL-HIER-JE-NEBULA-API-KEY-IN")
NEBULA_MODEL = os.environ.get("NEBULA_MODEL", "llama3.1:70b")  # kies een model dat op Nebula beschikbaar is
 
# ----------------------------------------------------------------------------
 
 
def normalize(text) -> str:
    """Lowercase, trim, en verwijder leestekens/dubbele spaties zodat kleine
    schrijfverschillen niet als 'verschillend' worden gezien."""
    if pd.isna(text):
        return ""
    text = str(text).strip().lower()
    text = re.sub(r"[^\w\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text
 
 
def load_data(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]  # spaties in kolomnamen weghalen (bijv. "Adres ")
 
    # "adres" en "postcode" zijn optioneel — als de kolom niet bestaat (of op
    # None staat), dedupliceren we gewoon zonder dat veld.
    for optioneel in ("adres", "postcode"):
        kolomnaam = COLUMN_MAP.get(optioneel)
        if kolomnaam and kolomnaam not in df.columns:
            print(f"Let op: {optioneel}-kolom '{kolomnaam}' niet gevonden, "
                  f"ontdubbeling gebeurt zonder dat veld.")
            COLUMN_MAP[optioneel] = None
        elif kolomnaam is None:
            COLUMN_MAP[optioneel] = None
 
    verplicht = ["naam", "plaats", "kvk"]
    missing = [COLUMN_MAP[k] for k in verplicht if COLUMN_MAP[k] not in df.columns]
    if missing:
        sys.exit(
            "Kolommen niet gevonden in het Excel-bestand: "
            f"{missing}\nControleer COLUMN_MAP bovenin dit script tegen de "
            f"echte kolomnamen: {list(df.columns)}"
        )
    return df
 
 
def add_helper_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["_naam_norm"] = df[COLUMN_MAP["naam"]].apply(normalize)
 
    delen = []
    if COLUMN_MAP.get("adres"):
        delen.append(df[COLUMN_MAP["adres"]].apply(normalize))
    if COLUMN_MAP.get("postcode"):
        delen.append(df[COLUMN_MAP["postcode"]].apply(normalize))
    delen.append(df[COLUMN_MAP["plaats"]].apply(normalize))
    df["_adres_norm"] = delen[0]
    for deel in delen[1:]:
        df["_adres_norm"] = df["_adres_norm"] + " " + deel
 
    df["_kvk_norm"] = df[COLUMN_MAP["kvk"]].apply(
        lambda x: re.sub(r"\D", "", str(x)) if pd.notna(x) else ""
    )
    df["_compleetheid"] = df.notna().sum(axis=1)  # meer ingevulde velden = "vollediger"
    df["_blok"] = df[COLUMN_MAP["plaats"]].apply(normalize)  # groeperen per plaats
    df["_leeg"] = (df["_naam_norm"] == "") & (df["_adres_norm"].str.strip() == "")
    return df
 
 
def step1_exact_duplicates(df: pd.DataFrame, report_rows: list) -> pd.DataFrame:
    """Verwijdert rijen die identiek zijn qua naam+adres, of hetzelfde (niet-lege)
    KvK-nummer hebben."""
    df = df.sort_values("_compleetheid", ascending=False)
 
    # Exacte naam+adres match
    before = len(df)
    dup_mask = df.duplicated(subset=["_naam_norm", "_adres_norm"], keep="first")
    for idx in df[dup_mask].index:
        report_rows.append(
            {"type": "exact (naam+adres)", "verwijderde_rij": idx, "score": 100}
        )
    df = df[~dup_mask]
 
    # Exacte KvK match (alleen als KvK is ingevuld)
    heeft_kvk = df["_kvk_norm"] != ""
    dup_kvk_mask = heeft_kvk & df.duplicated(subset=["_kvk_norm"], keep="first")
    for idx in df[dup_kvk_mask].index:
        report_rows.append(
            {"type": "exact (KvK-nummer)", "verwijderde_rij": idx, "score": 100}
        )
    df = df[~dup_kvk_mask]
 
    print(f"Stap 1 (exact): {before - len(df)} duplicaten verwijderd.")
    return df
 
 
def ask_nebula_ai(row_a: pd.Series, row_b: pd.Series) -> bool:
    """Vraagt Nebula AI of twee rijen dezelfde organisatie beschrijven.
    Retourneert True als de AI denkt dat het dezelfde organisatie is."""
    from openai import OpenAI
 
    client = OpenAI(base_url=NEBULA_BASE_URL, api_key=NEBULA_API_KEY)
 
    def fmt(row):
        adres = row[COLUMN_MAP["adres"]] if COLUMN_MAP.get("adres") else ""
        postcode = row[COLUMN_MAP["postcode"]] if COLUMN_MAP.get("postcode") else ""
        return (
            f"Naam: {row[COLUMN_MAP['naam']]}\n"
            f"Adres: {adres}, {postcode} {row[COLUMN_MAP['plaats']]}\n"
            f"KvK-nummer: {row[COLUMN_MAP['kvk']]}"
        )
 
    prompt = (
        "Je krijgt twee registraties uit een lijst met burgerinitiatieven. "
        "Bepaal of dit twee registraties van DEZELFDE organisatie zijn "
        "(bijv. door een tikfout, afkorting, of net iets ander adres), "
        "of dat het om twee VERSCHILLENDE organisaties gaat.\n\n"
        f"Organisatie A:\n{fmt(row_a)}\n\nOrganisatie B:\n{fmt(row_b)}\n\n"
        "Antwoord met precies één woord: JA (zelfde organisatie) of NEE (verschillend)."
    )
 
    response = client.chat.completions.create(
        model=NEBULA_MODEL,
        messages=[{"role": "user", "content": prompt}],
        temperature=0,
        max_tokens=5,
    )
    answer = response.choices[0].message.content.strip().upper()
    return answer.startswith("JA")
 
 
def step2_and_3_fuzzy_and_ai(df: pd.DataFrame, report_rows: list) -> pd.DataFrame:
    """Vergelijkt rijen binnen dezelfde plaats (om het aantal vergelijkingen
    beheersbaar te houden) met fuzzy matching, en legt twijfelgevallen voor
    aan Nebula AI."""
    to_drop = set()
 
    for plaats, groep in df.groupby("_blok"):
        indices = list(groep.index)
        for i in range(len(indices)):
            idx_a = indices[i]
            if idx_a in to_drop:
                continue
            for j in range(i + 1, len(indices)):
                idx_b = indices[j]
                if idx_b in to_drop:
                    continue
 
                row_a, row_b = df.loc[idx_a], df.loc[idx_b]
                naam_score = fuzz.token_sort_ratio(row_a["_naam_norm"], row_b["_naam_norm"])
                adres_score = fuzz.token_sort_ratio(row_a["_adres_norm"], row_b["_adres_norm"])
                score = 0.6 * naam_score + 0.4 * adres_score
 
                is_duplicate = None
                methode = None
 
                if score >= FUZZY_THRESHOLD_AUTO:
                    is_duplicate = True
                    methode = "fuzzy"
                elif score >= FUZZY_THRESHOLD_AI and USE_AI_CHECK:
                    try:
                        is_duplicate = ask_nebula_ai(row_a, row_b)
                        methode = "nebula-ai"
                    except Exception as e:
                        print(f"  Nebula AI-check mislukt ({e}); rij overgeslagen, handmatig checken.")
                        is_duplicate = None
                        methode = "nebula-ai (fout)"
 
                if is_duplicate:
                    # de minst complete rij van het paar verwijderen
                    verwijder_idx = idx_b if row_a["_compleetheid"] >= row_b["_compleetheid"] else idx_a
                    to_drop.add(verwijder_idx)
                    report_rows.append(
                        {
                            "type": methode,
                            "verwijderde_rij": verwijder_idx,
                            "vergeleken_met": idx_a if verwijder_idx == idx_b else idx_b,
                            "score": round(score, 1),
                            "naam_a": row_a[COLUMN_MAP["naam"]],
                            "naam_b": row_b[COLUMN_MAP["naam"]],
                        }
                    )
                elif methode == "nebula-ai (fout)":
                    report_rows.append(
                        {
                            "type": methode,
                            "verwijderde_rij": None,
                            "score": round(score, 1),
                            "naam_a": row_a[COLUMN_MAP["naam"]],
                            "naam_b": row_b[COLUMN_MAP["naam"]],
                            "let_op": "Kon niet automatisch worden gecontroleerd, handmatig nakijken.",
                        }
                    )
 
    print(f"Stap 2+3 (fuzzy/AI): {len(to_drop)} duplicaten verwijderd.")
    return df.drop(index=list(to_drop))
 
 
def main():
    df = load_data(INPUT_FILE)
    origineel_aantal = len(df)
    print(f"Origineel aantal rijen: {origineel_aantal}")
 
    df = add_helper_columns(df)
    report_rows = []
 
    # Rijen zonder naam én zonder adres/postcode bevatten geen bruikbare
    # informatie en kunnen ook niet betrouwbaar vergeleken worden (anders
    # worden ze onterecht als elkaars duplicaat gezien). Die halen we eruit
    # en melden we apart (i.p.v. ze als lege rij te bewaren — Excel/pandas
    # verliest volledig lege rijen sowieso bij het opslaan+inlezen).
    leeg_mask = df["_leeg"]
    if leeg_mask.any():
        print(f"Let op: {leeg_mask.sum()} volledig lege rijen (geen naam, geen adres/postcode) "
              "gevonden — deze zijn verwijderd uit de output, zie rapport voor de rijnummers.")
        for idx in df[leeg_mask].index:
            report_rows.append({"type": "leeg (verwijderd)", "verwijderde_rij": idx})
    df = df[~leeg_mask]
 
    df = step1_exact_duplicates(df, report_rows)
    df = step2_and_3_fuzzy_and_ai(df, report_rows)
 
    df = df.sort_index()
    df_clean = df.drop(columns=[c for c in df.columns if c.startswith("_")])
 
    print(f"Aantal rijen na ontdubbelen: {len(df_clean)} "
          f"({origineel_aantal - len(df_clean)} verwijderd totaal)")
 
    df_clean.to_excel(OUTPUT_FILE, index=False)
    print(f"Klaar! Opgeslagen als: {OUTPUT_FILE}")
 
    if report_rows:
        pd.DataFrame(report_rows).to_excel(REPORT_FILE, index=False)
        print(f"Audit-rapport (wat is verwijderd en waarom) opgeslagen als: {REPORT_FILE}")
        print("Controleer vooral de 'nebula-ai' en 'nebula-ai (fout)' rijen even met de hand.")
 
 
if __name__ == "__main__":
    main()
