"""STAGE 1 - EXTRACT
Downloads real Philippine data from official sources and saves it, uncleaned,
to raw/ph_food_economy_raw.xlsx (one sheet per source) for Power Query.

Sources
  PSA OpenSTAT (openstat.psa.gov.ph): rice retail prices, CPI inflation,
      purchasing power of the peso, quarterly GDP growth
  ECB Data Portal (data-api.ecb.europa.eu): USD/PHP via EUR reference rates
  Published releases (SWS, DOLE/RTWPB, DA): hunger, minimum wage, rice policy
      events - typed in below with their source links, since they have no API

Run:  python3 extract_data.py      (needs: pandas, openpyxl)
"""
import io, json, pathlib, urllib.request
from datetime import datetime, timezone
import pandas as pd

ROOT = pathlib.Path(__file__).parent
OUT = ROOT / "raw" / "ph_food_economy_raw.xlsx"
PSA = "https://openstat.psa.gov.ph/PXWeb/api/v1/en/DB/"
HEAD = {"User-Agent": "Mozilla/5.0 (portfolio data project)", "Content-Type": "application/json"}
YEARS = ("2024", "2025", "2026")


def http(url, body=None):
    req = urllib.request.Request(url, data=json.dumps(body).encode() if body else None, headers=HEAD)
    with urllib.request.urlopen(req, timeout=120) as r:
        return r.read()


def psa_table(path, pick):
    """Query a PSA PXWeb table. pick = {variable code: function(text) -> keep?}. Returns the raw wide CSV."""
    meta = json.loads(http(PSA + path))
    query = []
    for var in meta["variables"]:
        if var["code"] in pick:
            keep = [v for v, t in zip(var["values"], var["valueTexts"]) if pick[var["code"]](t)]
            assert keep, f"{path}: nothing selected for {var['code']}"
            query.append({"code": var["code"], "selection": {"filter": "item", "values": keep}})
    raw = http(PSA + path, {"query": query, "response": {"format": "csv"}})
    return pd.read_csv(io.BytesIO(raw), encoding="utf-8-sig", dtype=str)


region_or_country = lambda t: not t.startswith("....")  # drop provinces/cities, keep PH + regions

sheets = {}
print("PSA rice retail prices ...")
sheets["psa_rice_prices"] = psa_table("2M/2018NEW/0042M4ARN01.px", {
    "Geolocation": region_or_country,
    "Commodity": lambda t: t.startswith("RICE, REGULAR-MILLED"),
    "Year": lambda t: t in YEARS})

print("PSA CPI inflation ...")
sheets["psa_inflation"] = psa_table("2M/PI/CPI/2018NEW/0012M4ACP23.px", {
    "Geolocation": region_or_country,
    "Commodity Description": lambda t: t in ("0 - ALL ITEMS", "01.1 - FOOD", "01.1.1.12 - Rice"),
    "Year": lambda t: t in YEARS})

print("PSA purchasing power of the peso ...")
sheets["psa_peso_power"] = psa_table("2M/PI/CPI/2018NEW/0012M4ACP24.px", {
    "Geolocation": lambda t: t == "PHILIPPINES",
    "Year": lambda t: t in YEARS})

print("PSA GDP growth ...")
sheets["psa_gdp_growth"] = psa_table("2B/NA/QT/1SUM/0022B5BEXQ2.px", {
    "Type of Expenditure": lambda t: t.strip(".") == "Gross Domestic Product",
    "Type of Valuation": lambda t: t == "At Constant 2018 Prices",
    "Year": lambda t: t in ("2022-2023", "2023-2024", "2024-2025", "2025-2026")})

print("ECB exchange rates ...")
ecb = http("https://data-api.ecb.europa.eu/service/data/EXR/M.PHP+USD.EUR.SP00.A?startPeriod=2024-01&format=csvdata")
sheets["ecb_fx"] = pd.read_csv(io.BytesIO(ecb), dtype=str)[["CURRENCY", "CURRENCY_DENOM", "TIME_PERIOD", "OBS_VALUE", "TITLE"]]

# No API for these: values typed from the published releases linked in `source`.
sheets["sws_hunger"] = pd.DataFrame([
    ("Mar 2025", "2025-03-15", 27.2, "https://www.gmanetwork.com/news/topstories/nation/941035/hunger-survey-sws-march-2025/story/"),
    ("Jun 2025", "2025-06-25", 16.1, "https://www.philstar.com/headlines/2025/08/16/2465924/hunger-among-filipino-families-drops-161-june-sws"),
    ("Sep 2025", "2025-09-24", 22.0, "https://www.bworldonline.com/the-nation/2025/11/05/710365/hunger-rises-to-22-in-q3-sws/"),
    ("Nov 2025", "2025-11-24", 20.1, "https://mb.com.ph/2026/01/14/hunger-incidence-among-filipino-families-eases-in-q4-2025-sws"),
    ("Mar 2026", "2026-03-24", 23.2, "https://newsinfo.inquirer.net/2227801/sws-hunger-among-filipino-families-rises-to-23-2-in-march-2026"),
    ("Jun 2026", "2026-06-20", 20.3, "https://www.philstar.com/headlines/2026/08/26/2551864/hunger-slightly-down-203-sws"),
    ("Sep 2026", "2026-09-21", 22.6, "https://www.gmanetwork.com/news/topstories/nation/1005905/more-filipino-families-experienced-hunger-in-3q-of-2026-sws-report/story/"),
], columns=["survey", "fieldwork_start", "hunger_pct", "source"])

sheets["ncr_min_wage"] = pd.DataFrame([
    ("Wage Order NCR-24", "2023-07-16", 610, "https://nwpc.dole.gov.ph/regionandwages/current-real-minimum-wage-rates/"),
    ("Wage Order NCR-25", "2024-07-17", 645, "https://www.pna.gov.ph/articles/1227990"),
    ("Wage Order NCR-26", "2025-07-18", 695, "https://www.sunstar.com.ph/manila/ncr-wage-hike-starts-july-25"),
    ("Wage Order NCR-28", "2026-09-26", 755, "https://www.gmanetwork.com/news/money/economy/1003769/dole-p60-wage-hike-in-ncr-takes-effect-sept-26/story/"),
], columns=["wage_order", "effective_date", "daily_wage_php", "source"])

sheets["rice_policy_events"] = pd.DataFrame([
    ("2024-07-05", "Rice import tariff cut from 35% to 15% (EO 62)", "https://portcalls.com/marcos-issues-eo-62-modifying-import-tariffs-including-on-rice"),
    ("2025-09-01", "Imports of regular and well-milled rice suspended (EO 93)", "https://portcalls.com/ph-rice-import-ban-extended-to-end-2025/"),
    ("2026-01-01", "Price-indexed rice tariff (15-35%) starts; imports resume (EO 105)", "https://www.fas.usda.gov/data/gain/2026/02/philippines-philippines-implements-quarterly-rice-import-tariff-adjustment-based-international"),
], columns=["date", "event", "source"])

# PSA labels regions differently in each table; this lookup lets Power Query join them.
sheets["region_lookup"] = pd.DataFrame([
    ("PH", "Philippines", "National", "Philippines", "PHILIPPINES"),
    ("NCR", "NCR", "Luzon", "NCR - National Capital Region", "National Capital Region (NCR)"),
    ("CAR", "CAR", "Luzon", "CAR - Cordillera Administrative Region", "Cordillera Administrative Region (CAR)"),
    ("R01", "Ilocos", "Luzon", "I - Ilocos Region", "Region I (Ilocos Region)"),
    ("R02", "Cagayan Valley", "Luzon", "II - Cagayan Valley", "Region II (Cagayan Valley)"),
    ("R03", "Central Luzon", "Luzon", "III - Central Luzon", "Region III (Central Luzon)"),
    ("R4A", "CALABARZON", "Luzon", "IV-A - CALABARZON", "Region IV-A (CALABARZON)"),
    ("R4B", "MIMAROPA", "Luzon", "MIMAROPA Region", "MIMAROPA Region"),
    ("R05", "Bicol", "Luzon", "V - Bicol Region", "Region V (Bicol Region)"),
    ("R06", "Western Visayas", "Visayas", "VI - Western Visayas", "Region VI (Western Visayas)"),
    ("NIR", "Negros Island", "Visayas", "Negros Island Region (NIR)", "Negros Island Region (NIR)"),
    ("R07", "Central Visayas", "Visayas", "VII - Central Visayas", "Region VII (Central Visayas)"),
    ("R08", "Eastern Visayas", "Visayas", "VIII - Eastern Visayas", "Region VIII (Eastern Visayas)"),
    ("R09", "Zamboanga Peninsula", "Mindanao", "IX - Zamboanga Peninsula", "Region IX (Zamboanga Peninsula)"),
    ("R10", "Northern Mindanao", "Mindanao", "X - Northern Mindanao", "Region X (Northern Mindanao)"),
    ("R11", "Davao", "Mindanao", "XI - Davao Region", "Region XI (Davao Region)"),
    ("R12", "SOCCSKSARGEN", "Mindanao", "XII - SOCCSKSARGEN", "Region XII (SOCCSKSARGEN)"),
    ("R13", "Caraga", "Mindanao", "XIII - Caraga", "Region XIII (Caraga)"),
    ("BARMM", "BARMM", "Mindanao", "BARMM - Bangsamoro Autonomous Region in Muslim Mindanao",
     "Bangsamoro Autonomous Region in Muslim Mindanao (BARMM)"),
], columns=["region_key", "region", "island_group", "rice_label", "cpi_label"])

# Every PSA region label must be in the lookup, or Power Query would silently drop it.
strip = lambda s: s.lstrip(".")
assert set(map(strip, sheets["psa_rice_prices"]["Geolocation"])) <= set(sheets["region_lookup"]["rice_label"])
cpi_regions = set(map(strip, sheets["psa_inflation"]["Geolocation"])) - {"Areas Outside National Capital Region (AONCR)"}
assert cpi_regions <= set(sheets["region_lookup"]["cpi_label"]), cpi_regions - set(sheets["region_lookup"]["cpi_label"])

sheets["about"] = pd.DataFrame({
    "item": ["extracted_at_utc", "psa", "ecb", "note"],
    "value": [datetime.now(timezone.utc).isoformat(timespec="seconds"),
              "https://openstat.psa.gov.ph (tables 0042M4ARN01, 0012M4ACP23, 0012M4ACP24, 0022B5BEXQ2)",
              "https://data-api.ecb.europa.eu EXR M.PHP/USD.EUR (USD/PHP = PHP per EUR / USD per EUR)",
              "Raw, uncleaned. '.' means not yet released. Clean in Power Query (see power_query.m)."]})

for name, df in sheets.items():
    assert len(df) > 0, f"{name} is empty"
OUT.parent.mkdir(exist_ok=True)
with pd.ExcelWriter(OUT) as xl:
    for name, df in sheets.items():
        df.to_excel(xl, sheet_name=name, index=False)
print("Saved", OUT.relative_to(ROOT), {k: len(v) for k, v in sheets.items()})
