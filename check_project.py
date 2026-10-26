"""STAGE 6 - VALIDATE before sharing.  Run: python3 check_project.py
Checks that the three layers agree:
  raw Excel sheets  <->  power_query.m (sheet names, output columns)
  power_query.m     <->  measures.dax  (every Table[column] and [Measure] exists, brackets balance)
  clean data        ->   no blanks, values in sensible ranges"""
import re, sys, pathlib
import pandas as pd

root = pathlib.Path(__file__).parent
errors = []

# Columns each Power Query query outputs (stage 2), plus the DAX Calendar table (stage 3)
schema = {
    "Region": {"region_key", "region", "island_group", "rice_label", "cpi_label"},
    "RicePrice": {"region_key", "month", "price_php_kg"},
    "Inflation": {"region_key", "month", "item", "yoy_pct"},
    "PesoPower": {"month", "peso_power"},
    "GDP": {"quarter", "quarter_start", "gdp_growth_pct"},
    "FX": {"month", "usd_php"},
    "Hunger": {"survey", "month", "hunger_pct", "source"},
    "Wage": {"wage_order", "effective_date", "daily_wage_php", "source"},
    "Events": {"date", "event", "source"},
    "Calendar": {"Date", "Year", "Month Start", "Month", "Month Number", "Quarter"},
}

# 1) power_query.m: every sheet it reads exists; each query's final columns match the schema
m = (root / "power_query.m").read_text()
sheets = set(pd.ExcelFile(root / "raw" / "ph_food_economy_raw.xlsx").sheet_names)
for s in re.findall(r'Item = "([^"]+)"', m):
    if s not in sheets: errors.append(f"power_query.m reads missing sheet '{s}'")
blocks = re.split(r"// ===== Query \d+: (\w+)", m)
queries = dict(zip(blocks[1::2], blocks[2::2]))
for q in schema:
    if q == "Calendar": continue
    if q not in queries: errors.append(f"power_query.m has no query {q}"); continue
    sel = re.findall(r"Table\.SelectColumns\(\w+, \{([^}]*)\}\)\s*in", queries[q])
    if sel:
        cols = set(re.findall(r'"([^"]+)"', sel[-1]))
        if cols != schema[q]: errors.append(f"query {q} outputs {sorted(cols)}, expected {sorted(schema[q])}")
    body = queries[q]
    if body.count("(") != body.count(")"): errors.append(f"query {q}: unbalanced ( )")
    if body.count("{") != body.count("}"): errors.append(f"query {q}: unbalanced {{ }}")

# 2) measures.dax: references resolve, brackets balance, no duplicates
text = (root / "measures.dax").read_text()
code = "\n".join(l for l in text.splitlines() if not l.lstrip().startswith("//"))
parts = re.split(r'^(?!VAR\b|RETURN\b)([A-Za-z][^=\n\["]*?)\s*=(?!=)', code, flags=re.M)
measures = {}
for name, body in zip(parts[1::2], parts[2::2]):
    name = name.strip()
    if name in measures: errors.append(f"duplicate measure: {name}")
    measures[name] = body
for name, body in measures.items():
    if body.count("(") != body.count(")"): errors.append(f"{name}: unbalanced ( )")
    if body.count("[") != body.count("]"): errors.append(f"{name}: unbalanced [ ]")
    for t, c in re.findall(r"'?\b([A-Z][A-Za-z]+)'?\[([^\]]+)\]", body):
        if t not in schema: errors.append(f"{name}: unknown table {t}")
        elif c not in schema[t]: errors.append(f"{name}: unknown column {t}[{c}]")
    for ref in re.findall(r"(?<![\w\]'])\[([^\]]+)\]", body):
        if ref not in measures: errors.append(f"{name}: unknown measure [{ref}]")
assert len(measures) >= 40, f"only parsed {len(measures)} measures - parser broken?"

# 3) clean data (from clean_reference.py): complete and plausible
for f in ("national_monthly.csv", "fx_monthly.csv", "regions_latest.csv", "gdp_quarterly.csv", "hunger_sws.csv", "rice_policy_events.csv"):
    d = pd.read_csv(root / "data" / f)
    if d.isna().any().any(): errors.append(f"data/{f} has blanks")
nat = pd.read_csv(root / "data" / "national_monthly.csv")
if not nat["usd_php"].between(40, 90).all(): errors.append("usd_php out of range")
if not nat["headline_inflation_pct"].between(-10, 30).all(): errors.append("inflation out of range")
if nat["month"].duplicated().any(): errors.append("duplicate months in national_monthly.csv")
# FX: monthly series and the provisional month agree with the national table and the raw ECB rows
fx = pd.read_csv(root / "data" / "fx_monthly.csv")
both = nat.merge(fx, on="month", suffixes=("", "_fx"))
if len(both) != len(nat) or (both["usd_php"] - both["usd_php_fx"]).abs().max() > 0.005: errors.append("fx_monthly.csv disagrees with national_monthly.csv")
if fx["provisional"].iloc[:-1].any(): errors.append("a provisional FX month is not the newest")
if not fx["usd_php"].between(40, 90).all() or fx["usd_php"].diff().abs().max() > 3: errors.append("fx_monthly.csv has an implausible value or jump")
raw_fx = pd.read_excel(root / "raw" / "ph_food_economy_raw.xlsx", sheet_name="ecb_fx", dtype=str)
if raw_fx["TITLE"].str.startswith("PROVISIONAL").sum() != 2 * fx["provisional"].sum(): errors.append("provisional FX rows in raw Excel don't match fx_monthly.csv")
# Daily rates: same month as the provisional one, and their average matches it
fd = pd.read_csv(root / "data" / "fx_daily.csv")
if len(fd):
    pm = fx.loc[fx["provisional"] == 1, "month"]
    if pm.empty or not fd["date"].str[:7].eq(pm.iloc[0][:7]).all(): errors.append("fx_daily.csv is not the provisional month")
    elif abs(fd["usd_php"].mean() - fx["usd_php"].iloc[-1]) > 0.05: errors.append("fx_daily.csv average disagrees with the provisional month")
# Forecast tile/chart data: the page must read the fx table
page = (root / "web-dashboard" / "index.html").read_text()
if "inputs: ['fx']" not in page: errors.append("dollar forecast is not fed by fx_monthly.csv")

print(f"Checked {len(measures)} DAX measures, {len(queries)} Power Query queries, {len(sheets)} raw sheets.")
if errors:
    print("PROBLEMS:", *errors, sep="\n  ")
    sys.exit(1)
print("OK - no problems found.")
