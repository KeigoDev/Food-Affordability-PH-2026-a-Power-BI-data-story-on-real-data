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
for f in ("national_monthly.csv", "regions_latest.csv", "gdp_quarterly.csv", "hunger_sws.csv", "rice_policy_events.csv"):
    d = pd.read_csv(root / "data" / f)
    if d.isna().any().any(): errors.append(f"data/{f} has blanks")
nat = pd.read_csv(root / "data" / "national_monthly.csv")
if not nat["usd_php"].between(40, 90).all(): errors.append("usd_php out of range")
if not nat["headline_inflation_pct"].between(-10, 30).all(): errors.append("inflation out of range")
if nat["month"].duplicated().any(): errors.append("duplicate months in national_monthly.csv")

print(f"Checked {len(measures)} DAX measures, {len(queries)} Power Query queries, {len(sheets)} raw sheets.")
if errors:
    print("PROBLEMS:", *errors, sep="\n  ")
    sys.exit(1)
print("OK - no problems found.")
