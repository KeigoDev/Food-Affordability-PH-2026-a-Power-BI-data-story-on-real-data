"""Builds the Power BI Project (powerbi/FoodAffordabilityPH.pbip) from power_query.m and measures.dax,
so the report always matches the documented Power Query and DAX.  Run: python3 build_pbip.py
Open the .pbip in Power BI Desktop (Windows), set the RawFile parameter, then Refresh."""
import json, re, uuid, shutil, pathlib, hashlib

ROOT = pathlib.Path(__file__).parent
OUT = ROOT / "powerbi"
NAME = "FoodAffordabilityPH"
DEFAULT_RAW = r"C:\PH POWER BI\food-affordability-powerbi\raw\ph_food_economy_raw.xlsx"

# ---------- 1. Power Query: one block per query in power_query.m ----------
m_text = (ROOT / "power_query.m").read_text()
blocks = re.split(r"// ===== Query \d+: (\w+)[^\n]*\n", m_text)
queries = {}
for name, body in zip(blocks[1::2], blocks[2::2]):
    lines = [l for l in body.strip("\n").splitlines()]
    while lines and (lines[0].lstrip().startswith("//") or not lines[0].strip()):
        lines.pop(0)
    queries[name] = "\n".join(lines).rstrip()
assert queries.pop("RawFile") == "", "RawFile block should only hold instructions"

# ---------- 2. DAX measures: one block per measure in measures.dax ----------
dax_text = "\n".join(l for l in (ROOT / "measures.dax").read_text().splitlines() if not l.lstrip().startswith("//"))
parts = re.split(r'^(?!VAR\b|RETURN\b)([A-Za-z][^=\n\["]*?)\s*=(?!=)', dax_text, flags=re.M)
measures = {n.strip(): b.strip() for n, b in zip(parts[1::2], parts[2::2])}


def fmt(name):
    if name.endswith("%"): return "0.0%"
    if name in ("Latest Month", "Year Ago Month", "Latest GDP Quarter"): return "mmm yyyy"
    if name.startswith(("Rice Price", "USD PHP", "Peso Purchasing", "NCR Rice Price")): return "0.00"
    if name.startswith(("Months Above", "NCR Min Wage", "Inflation Target")): return "0"
    if name in ("Selected Region Key", "Inflation Status", "Rice Pressure Flag", "Latest GDP Quarter Label",
                "Headline", "Economy Headline", "Inflation Colour", "Rice Pressure Colour"): return None
    return "0.0"


# ---------- 3. Tables: columns each query outputs (same names the checker validates) ----------
S, D, N, I = "string", "dateTime", "double", "int64"
tables = {
    "Region": [("region_key", S), ("region", S), ("island_group", S), ("rice_label", S), ("cpi_label", S)],
    "RicePrice": [("region_key", S), ("month", D), ("price_php_kg", N)],
    "Inflation": [("region_key", S), ("month", D), ("item", S), ("yoy_pct", N)],
    "PesoPower": [("month", D), ("peso_power", N)],
    "GDP": [("quarter", S), ("quarter_start", D), ("gdp_growth_pct", N)],
    "FX": [("month", D), ("usd_php", N)],
    "Hunger": [("survey", S), ("month", D), ("hunger_pct", N), ("source", S)],
    "Wage": [("wage_order", S), ("effective_date", D), ("daily_wage_php", N), ("source", S)],
    "Events": [("date", D), ("event", S), ("source", S)],
}
hidden = {("Region", "rice_label"), ("Region", "cpi_label"), ("RicePrice", "region_key"), ("Inflation", "region_key")}


def column(table, name, dtype):
    c = {"name": name, "dataType": dtype, "sourceColumn": name, "summarizeBy": "none"}
    if dtype == D: c["formatString"] = "yyyy-mm-dd"
    if (table, name) in hidden: c["isHidden"] = True
    return c


model_tables = []
for t, cols in tables.items():
    model_tables.append({
        "name": t,
        "columns": [column(t, n, d) for n, d in cols],
        "partitions": [{"name": t, "mode": "import", "source": {"type": "m", "expression": queries[t].splitlines()}}],
    })

calendar_dax = ('ADDCOLUMNS ( CALENDAR ( DATE ( 2023, 1, 1 ), DATE ( 2026, 12, 31 ) ), '
                '"Year", YEAR ( [Date] ), "Month Start", DATE ( YEAR ( [Date] ), MONTH ( [Date] ), 1 ), '
                '"Month", FORMAT ( [Date], "mmm yyyy" ), "Month Number", YEAR ( [Date] ) * 100 + MONTH ( [Date] ), '
                '"Quarter", YEAR ( [Date] ) & " Q" & QUARTER ( [Date] ) )')
cal_cols = [("Date", D, True), ("Year", I, False), ("Month Start", D, False), ("Month", S, False), ("Month Number", I, False), ("Quarter", S, False)]
calendar = {"name": "Calendar", "dataCategory": "Time", "columns": [], "partitions": [
    {"name": "Calendar", "mode": "import", "source": {"type": "calculated", "expression": calendar_dax}}]}
for n, d, key in cal_cols:
    c = {"type": "calculatedTableColumn", "name": n, "dataType": d, "isNameInferred": True,
         "isDataTypeInferred": True, "sourceColumn": f"[{n}]", "summarizeBy": "none"}
    if d == D: c["formatString"] = "mmm yyyy" if n == "Month Start" else "yyyy-mm-dd"
    if key: c["isKey"] = True
    if n == "Month": c["sortByColumn"] = "Month Number"
    calendar["columns"].append(c)
model_tables.append(calendar)

measure_table = {"name": "_Measures", "columns": [
    {"type": "calculatedTableColumn", "name": "Hidden", "dataType": "string", "isNameInferred": True,
     "isDataTypeInferred": True, "sourceColumn": "[Hidden]", "isHidden": True}],
    "partitions": [{"name": "_Measures", "mode": "import", "source": {"type": "calculated", "expression": 'ROW ( "Hidden", "" )'}}],
    "measures": []}
for n, expr in measures.items():
    entry = {"name": n, "expression": expr.splitlines()}
    if fmt(n): entry["formatString"] = fmt(n)
    measure_table["measures"].append(entry)
model_tables.append(measure_table)

rels = [("RicePrice", "region_key", "Region", "region_key"), ("Inflation", "region_key", "Region", "region_key"),
        ("RicePrice", "month", "Calendar", "Date"), ("Inflation", "month", "Calendar", "Date"),
        ("FX", "month", "Calendar", "Date"), ("PesoPower", "month", "Calendar", "Date"),
        ("Hunger", "month", "Calendar", "Date"), ("GDP", "quarter_start", "Calendar", "Date")]
relationships = [{"name": hashlib.md5(f"{a}{b}{c}{d}".encode()).hexdigest(), "fromTable": a, "fromColumn": b,
                  "toTable": c, "toColumn": d} for a, b, c, d in rels]

model = {"compatibilityLevel": 1567, "model": {
    "culture": "en-US", "defaultPowerBIDataSourceVersion": "powerBI_V3", "sourceQueryCulture": "en-US",
    "dataAccessOptions": {"legacyRedirects": True, "returnErrorValuesAsNull": True},
    "expressions": [
        {"name": "RawFile", "kind": "m",
         "expression": json.dumps(DEFAULT_RAW) + ' meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]'},
        {"name": "fnToNumber", "kind": "m", "expression": queries.pop("fnToNumber").splitlines()}],
    "tables": model_tables, "relationships": relationships,
    "annotations": [{"name": "__PBI_TimeIntelligenceEnabled", "value": "0"}]}}
assert set(queries) == set(tables), f"queries without a table: {set(queries) ^ set(tables)}"

# ---------- 4. Report: 4 story pages ----------
fields = {t: {c for c, _ in cols} for t, cols in tables.items()}
fields["Calendar"] = {n for n, _, _ in cal_cols}
fields["_Measures"] = set(measures)


def ref(table, field):
    assert field in fields[table], f"report uses unknown field {table}[{field}]"
    return table, field


def visual(vtype, x, y, w, h, roles, title=None, z=0):
    """roles: {projection role: [(table, field), ...]}"""
    froms, selects, projections = {}, [], {}
    for role, items in roles.items():
        projections[role] = []
        for table, f in (ref(*i) for i in items):
            alias = froms.setdefault(table, "t%d" % len(froms))
            kind = "Measure" if table == "_Measures" else "Column"
            q = f"{table}.{f}"
            if q not in [s["Name"] for s in selects]:
                selects.append({kind: {"Expression": {"SourceRef": {"Source": alias}}, "Property": f}, "Name": q})
            projections[role].append({"queryRef": q, **({"active": True} if role == "Category" else {})})
    sv = {"visualType": vtype, "projections": projections,
          "prototypeQuery": {"Version": 2, "From": [{"Name": a, "Entity": t, "Type": 0} for t, a in froms.items()], "Select": selects},
          "drillFilterOtherVisuals": True}
    if title:
        sv["vcObjects"] = {"title": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}},
                                                     "text": {"expr": {"Literal": {"Value": "'" + title.replace("'", "''") + "'"}}}}}]}
    return container(x, y, w, h, sv, z)


def textbox(x, y, w, h, text, size="20pt", bold=True, z=0):
    style = {"fontSize": size, **({"fontWeight": "bold"} if bold else {})}
    sv = {"visualType": "textbox", "drillFilterOtherVisuals": True,
          "objects": {"general": [{"properties": {"paragraphs": [{"textRuns": [{"value": text, "textStyle": style}]}]}}]}}
    return container(x, y, w, h, sv, z)


def container(x, y, w, h, sv, z):
    name = uuid.uuid5(uuid.NAMESPACE_URL, json.dumps(sv, sort_keys=True) + f"{x},{y}").hex[:20]
    cfg = {"name": name, "layouts": [{"id": 0, "position": {"x": x, "y": y, "z": z, "width": w, "height": h}}], "singleVisual": sv}
    return {"x": x, "y": y, "z": z, "width": w, "height": h, "config": json.dumps(cfg), "filters": "[]"}


M = "_Measures"
month = ("Calendar", "Month Start")
# Same date slicer on pages 1-3: the cards show its last month (see measures.dax)
slicer_dates = lambda: visual("slicer", 920, 10, 340, 60, {"Values": [month]}, "Date range")
cards = lambda items, y=80: [visual("card", 20 + i * 250, y, 240, 110, {"Values": [(M, f)]}, t) for i, (f, t) in enumerate(items)]

pages = [
    ("1 · Bigger picture", [
        textbox(20, 10, 880, 60, "Growth is slowing, the dollar is strong, and inflation is back"), slicer_dates(),
        *cards([("GDP Growth (Latest)", "GDP growth, latest quarter (%)"), ("USD PHP (Latest)", "Pesos per US dollar"),
                ("Peso Weakening YoY %", "Peso weaker vs a year ago"), ("Headline Inflation (Latest)", "Headline inflation (%)"),
                ("Peso Purchasing Power (Latest)", "Purchasing power of ₱1 (2018 = 1)")]),
        visual("clusteredColumnChart", 20, 210, 400, 490, {"Category": [("GDP", "quarter")], "Y": [(M, "GDP Growth")]}, "GDP growth by quarter (%)"),
        visual("lineChart", 440, 210, 400, 490, {"Category": [month], "Y": [(M, "USD PHP")]}, "Pesos per US dollar (monthly average)"),
        visual("lineChart", 860, 210, 400, 490, {"Category": [month], "Y": [(M, "Headline Inflation"), (M, "Food Inflation")]}, "Inflation: headline vs food (%), BSP target 2–4%"),
    ]),
    ("2 · The rice shock", [
        textbox(20, 10, 880, 60, "Import policy swung rice prices down, then sharply up"), slicer_dates(),
        *cards([("Rice Price (Latest)", "Rice, regular-milled (₱/kg)"), ("Rice Price Change YoY %", "Rice price vs a year ago"),
                ("Rice Inflation (Latest)", "Rice inflation (%)"), ("Kg Rice per NCR Wage (Latest)", "Kg of rice per NCR wage day"),
                ("Kg Rice per NCR Wage (Peak)", "Best point (kg)")]),
        visual("lineChart", 20, 210, 610, 330, {"Category": [month], "Y": [(M, "Rice Price")]}, "Rice retail price, Philippines (₱/kg)"),
        visual("lineChart", 650, 210, 610, 330, {"Category": [month], "Y": [(M, "Kg Rice per NCR Wage")]}, "Kg of rice one NCR minimum-wage day buys"),
        visual("tableEx", 20, 555, 1240, 150, {"Values": [("Events", "date"), ("Events", "event")]}, "What changed: rice import policy"),
    ]),
    ("3 · Hunger", [
        textbox(20, 10, 880, 60, "Hunger stays stubbornly high"), slicer_dates(),
        *cards([("Hunger Rate (Latest)", "Families going hungry, latest SWS (%)"), ("Hunger Rate (Low)", "Lowest in the selected range (%)")]),
        visual("clusteredColumnChart", 20, 210, 1240, 490, {"Category": [("Hunger", "month")], "Y": [(M, "Hunger Rate")]}, "Families experiencing involuntary hunger, per SWS survey (%)"),
    ]),
    ("4 · Where it hurts most", [
        textbox(20, 10, 1240, 60, "Rice inflation by region: where relief should go first"),
        visual("slicer", 20, 80, 300, 110, {"Values": [("Region", "island_group")]}, "Island group"),
        *[visual("card", 340 + i * 310, 80, 300, 110, {"Values": [(M, f)]}, t) for i, (f, t) in enumerate(
            [("Rice Inflation (National Latest)", "National rice inflation (%)"), ("Headline Inflation (Latest)", "Headline inflation (%)"),
             ("Rice Price (Latest)", "Rice price (₱/kg)")])],
        visual("clusteredBarChart", 20, 210, 520, 490, {"Category": [("Region", "region")], "Y": [(M, "Regional Rice Inflation")]}, "Rice inflation by region, latest month (%)"),
        visual("tableEx", 560, 210, 700, 490, {"Values": [("Region", "region"), (M, "Rice Price (Latest)"), (M, "Rice Price Change YoY %"),
                                                         (M, "Rice Inflation (Latest)"), (M, "Food Inflation (Latest)"),
                                                         (M, "Headline Inflation (Latest)"), (M, "Rice Pressure Flag")]}, "Regional detail (latest month)"),
    ]),
]
report = {
    "config": json.dumps({"version": "5.43", "themeCollection": {}, "activeSectionIndex": 0,
                          "defaultDrillFilterOtherVisuals": True, "settings": {"useStylableVisualContainerHeader": True}}),
    "layoutOptimization": 0,
    "sections": [{"name": "ReportSection%d" % i, "displayName": name, "displayOption": 1, "width": 1280, "height": 720,
                  "config": "{}", "filters": "[]", "ordinal": i, "visualContainers": vis} for i, (name, vis) in enumerate(pages)]}

# ---------- 5. Write the project ----------
if OUT.exists(): shutil.rmtree(OUT)
rep, sem = OUT / f"{NAME}.Report", OUT / f"{NAME}.SemanticModel"
rep.mkdir(parents=True); sem.mkdir()
platform = lambda kind: {"$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
                         "metadata": {"type": kind, "displayName": NAME},
                         "config": {"version": "2.0", "logicalId": str(uuid.uuid5(uuid.NAMESPACE_URL, NAME + kind))}}
files = {
    OUT / f"{NAME}.pbip": {"version": "1.0", "artifacts": [{"report": {"path": f"{NAME}.Report"}}], "settings": {"enableAutoRecovery": True}},
    rep / "definition.pbir": {"version": "4.0", "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}},
    rep / "report.json": report,
    rep / ".platform": platform("Report"),
    sem / "definition.pbism": {"version": "4.0", "settings": {}},
    sem / "model.bim": model,
    sem / ".platform": platform("SemanticModel"),
}
for path, obj in files.items():
    path.write_text(json.dumps(obj, indent=2, ensure_ascii=False), encoding="utf-8")
    json.loads(path.read_text(encoding="utf-8"))  # every file must be valid JSON
(OUT / ".gitignore").write_text("**/.pbi/localSettings.json\n**/.pbi/cache.abf\n")

n_vis = sum(len(v) for _, v in pages)
print(f"Built {OUT.relative_to(ROOT)}/{NAME}.pbip: {len(model_tables)} tables, {len(relationships)} relationships, "
      f"{len(measures)} measures, {len(pages)} pages, {n_vis} visuals")
