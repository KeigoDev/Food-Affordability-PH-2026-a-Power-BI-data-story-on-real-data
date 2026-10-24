"""Draws a layout preview of the Power BI report pages (preview/report_layout.html) from
powerbi/.../report.json, filled with the expected values from data/*.csv.
It is NOT a Power BI screenshot: it shows where each visual sits and what it should show."""
import json, html, pathlib
import pandas as pd

ROOT = pathlib.Path(__file__).parent
rep = json.loads((ROOT / "powerbi/FoodAffordabilityPH.Report/report.json").read_text())
nat = pd.read_csv(ROOT / "data/national_monthly.csv")
gdp = pd.read_csv(ROOT / "data/gdp_quarterly.csv")
hun = pd.read_csv(ROOT / "data/hunger_sws.csv")
reg = pd.read_csv(ROOT / "data/regions_latest.csv")
ev = pd.read_csv(ROOT / "data/rice_policy_events.csv")
l, ya = nat.iloc[-1], nat[nat.month == f"{int(nat.iloc[-1].month[:4]) - 1}{nat.iloc[-1].month[4:]}"].iloc[0]
pct = lambda a, b: f"{(b / a - 1) * 100:.1f}%"

cards = {
    "GDP Growth (Latest)": f"{gdp.iloc[-1].gdp_growth_pct:.1f}", "USD PHP (Latest)": f"{l.usd_php:.2f}",
    "Peso Weakening YoY %": pct(ya.usd_php, l.usd_php), "Headline Inflation (Latest)": f"{l.headline_inflation_pct:.1f}",
    "Peso Purchasing Power (Latest)": f"{l.peso_power:.2f}", "Rice Price (Latest)": f"{l.rice_price_php_kg:.2f}",
    "Rice Price Change YoY %": pct(ya.rice_price_php_kg, l.rice_price_php_kg), "Rice Inflation (Latest)": f"{l.rice_inflation_pct:.1f}",
    "Kg Rice per NCR Wage (Latest)": f"{l.kg_rice_per_ncr_wage:.1f}", "Kg Rice per NCR Wage (Peak)": f"{nat.kg_rice_per_ncr_wage.max():.1f}",
    "Hunger Rate (Latest)": f"{hun.iloc[-1].hunger_pct:.1f}", "Hunger Rate (Low)": f"{hun.hunger_pct.min():.1f}",
    "Rice Inflation (National Latest)": f"{l.rice_inflation_pct:.1f}",
}
series = {  # measure -> (labels, values)
    "GDP Growth": (gdp.quarter.tolist(), gdp.gdp_growth_pct.tolist()),
    "USD PHP": (nat.month.tolist(), nat.usd_php.tolist()),
    "Headline Inflation": (nat.month.tolist(), nat.headline_inflation_pct.tolist()),
    "Food Inflation": (nat.month.tolist(), nat.food_inflation_pct.tolist()),
    "Rice Price": (nat.month.tolist(), nat.rice_price_php_kg.tolist()),
    "Kg Rice per NCR Wage": (nat.month.tolist(), nat.kg_rice_per_ncr_wage.tolist()),
    "Hunger Rate": (hun.survey.tolist(), hun.hunger_pct.tolist()),
}
colors = ["#C2410C", "#64748B"]


def svg_chart(kind, names, w, h):
    vals = [series[n][1] for n in names]
    lo, hi = min(min(v) for v in vals + [[0]] if kind != "line" or True), max(max(v) for v in vals)
    if kind == "line": lo = min(min(v) for v in vals)
    lo, hi = (lo - (hi - lo) * .1, hi + (hi - lo) * .1) if kind == "line" else (min(0, lo), hi * 1.1)
    y = lambda v: h - 20 - (v - lo) / (hi - lo) * (h - 30)
    out = []
    for i, v in enumerate(vals):
        n = len(v)
        if kind == "line":
            pts = " ".join(f"{30 + j * (w - 40) / (n - 1):.1f},{y(x):.1f}" for j, x in enumerate(v))
            out.append(f'<polyline points="{pts}" fill="none" stroke="{colors[i]}" stroke-width="2"/>')
        else:
            bw = (w - 40) / n
            out += [f'<rect x="{30 + j * bw + bw * .15:.1f}" y="{y(x):.1f}" width="{bw * .7:.1f}" height="{y(0) - y(x):.1f}" fill="{colors[0] if j == n - 1 else "#B9B5AA"}"/>'
                    f'<text x="{30 + j * bw + bw / 2:.1f}" y="{y(x) - 3:.1f}" font-size="10" text-anchor="middle">{x:.1f}</text>' for j, x in enumerate(v)]
    labels = series[names[0]][0]
    out.append(f'<text x="30" y="{h - 4}" font-size="10" fill="#6b6a66">{labels[0]}</text><text x="{w - 10}" y="{h - 4}" font-size="10" fill="#6b6a66" text-anchor="end">{labels[-1]}</text>')
    return f'<svg width="{w}" height="{h}">{"".join(out)}</svg>'


def body(sv, w, h):
    t, proj = sv["visualType"], sv.get("projections", {})
    meas = [p["queryRef"].split(".", 1)[1] for role in ("Y", "Values") for p in proj.get(role, []) if p["queryRef"].startswith("_Measures.")]
    if t == "textbox":
        return f'<div class="tb">{html.escape(sv["objects"]["general"][0]["properties"]["paragraphs"][0]["textRuns"][0]["value"])}</div>'
    if t == "card":
        return f'<div class="card">{cards[meas[0]]}</div>'
    if t == "slicer":
        return '<div class="sl">' + "".join(f"<span>☐ {g}</span>" for g in ["Luzon", "Mindanao", "National", "Visayas"]) + "</div>"
    if t in ("lineChart", "clusteredColumnChart"):
        return svg_chart("line" if t == "lineChart" else "bar", meas, w - 16, h - 50)
    if t == "clusteredBarChart":
        r = reg.sort_values("rice_inflation_pct", ascending=False)
        mx = r.rice_inflation_pct.max()
        return "".join(f'<div class="hb"><span>{x.region}</span><i style="width:{x.rice_inflation_pct / mx * 55:.0f}%"></i><b>{x.rice_inflation_pct:.1f}</b></div>' for x in r.itertuples())
    if t == "tableEx":
        if "Events.event" in [p["queryRef"] for p in proj["Values"]]:
            return "<table>" + "".join(f"<tr><td>{e.date}</td><td>{html.escape(e.event)}</td></tr>" for e in ev.itertuples()) + "</table>"
        r = reg.sort_values("rice_inflation_pct", ascending=False)
        head = "<tr><th>Region</th><th>Rice ₱/kg</th><th>Change</th><th>Rice infl.</th><th>Food</th><th>All</th><th>Flag</th></tr>"
        rows = "".join(f"<tr><td>{x.region}</td><td>{x.rice_price_php_kg:.2f}</td><td>{x.rice_price_change_pct:.1f}%</td><td>{x.rice_inflation_pct:.1f}</td>"
                       f"<td>{x.food_inflation_pct:.1f}</td><td>{x.headline_inflation_pct:.1f}</td>"
                       f"<td>{'Above national' if x.rice_inflation_pct > l.rice_inflation_pct else 'At or below'}</td></tr>" for x in r.itertuples())
        return f"<table>{head}{rows}</table>"
    return t


pages = []
for sec in rep["sections"]:
    items = []
    for vc in sec["visualContainers"]:
        sv = json.loads(vc["config"])["singleVisual"]
        title = sv.get("vcObjects", {}).get("title", [{}])[0].get("properties", {}).get("text", {}).get("expr", {}).get("Literal", {}).get("Value", "")[1:-1]
        cls = "v tbv" if sv["visualType"] == "textbox" else "v"
        items.append(f'<div class="{cls}" style="left:{vc["x"]}px;top:{vc["y"]}px;width:{vc["width"]}px;height:{vc["height"]}px">'
                     f'{f"<h4>{html.escape(title)}</h4>" if title else ""}{body(sv, vc["width"], vc["height"])}</div>')
    pages.append(f'<h2>{html.escape(sec["displayName"])}</h2><div class="page">{"".join(items)}</div>')

doc = f"""<!doctype html><html><head><meta charset="utf-8"><title>Report layout preview</title><style>
body{{font-family:Segoe UI,system-ui,sans-serif;background:#e9e7e1;margin:20px;color:#1f2937}}
.note{{background:#fff7ed;border:1px solid #fdba74;padding:8px 12px;border-radius:6px;display:inline-block;font-size:13px}}
h2{{font-size:15px;margin:22px 0 6px}}
.page{{position:relative;width:1280px;height:720px;background:#FAFAF9;box-shadow:0 2px 8px rgba(0,0,0,.15)}}
.v{{position:absolute;background:#fff;border:1px solid #E5E7EB;border-radius:8px;padding:6px 8px;box-sizing:border-box;overflow:hidden}}
.tbv{{background:none;border:0}} .tb{{font-size:24px;font-weight:700}}
h4{{margin:0 0 4px;font-size:12px;font-weight:600}} .card{{font-family:Georgia,serif;font-size:32px;margin-top:10px}}
.sl span{{display:block;font-size:12px;margin:2px 0}}
.hb{{display:flex;align-items:center;font-size:11px;margin:3px 0}} .hb span{{width:130px;text-align:right;padding-right:6px}}
.hb i{{display:block;height:14px;background:#C2410C;border-radius:3px}} .hb b{{font-weight:400;padding-left:4px}}
table{{border-collapse:collapse;font-size:11px;width:100%}} td,th{{border-bottom:1px solid #eee;padding:3px 5px;text-align:left}}
</style></head><body><div class="note">Layout preview generated from the Power BI Project's report file, filled with expected values from the real data. Not a Power BI screenshot; open the .pbip in Power BI Desktop for the real report.</div>
{"".join(pages)}</body></html>"""
(ROOT / "preview").mkdir(exist_ok=True)
(ROOT / "preview/report_layout.html").write_text(doc, encoding="utf-8")
print("wrote preview/report_layout.html with", len(pages), "pages")
