"""Builds dashboard.html: the interactive dashboard as one standalone file (double-click to open).
It combines the page in web-dashboard/index.html with the clean data from data/*.csv and a small
data layer. Needs internet once to load d3 from cdnjs.  Run after clean_reference.py:
    python3 build_web_dashboard.py"""
import base64, io, json, pathlib, zipfile

ROOT = pathlib.Path(__file__).parent
page = (ROOT / "web-dashboard" / "index.html").read_text(encoding="utf-8")
datasets = {"nat": "national_monthly.csv", "regions": "regions_latest.csv", "gdp": "gdp_quarterly.csv",
            "hunger": "hunger_sws.csv", "events": "rice_policy_events.csv"}
csv = {k: (ROOT / "data" / f).read_text(encoding="utf-8") for k, f in datasets.items()}
assert all(len(v.splitlines()) > 1 for v in csv.values()), "a data file is empty - run clean_reference.py"

TOKENS = """
:root { color-scheme: light dark;
  --font-sans: "Segoe UI", system-ui, -apple-system, sans-serif; --font-serif: Georgia, "Times New Roman", serif;
  --ui-font-size-caption: 12px; --ui-font-size-body: 15px; --ui-font-size-heading: 18px; --ui-font-size-title: 30px; --ui-font-weight-medium: 550;
  --ui-gap-xs: 4px; --ui-gap-sm: 8px; --ui-gap-md: 14px; --ui-gap-lg: 22px; --ui-gap-xl: 36px;
  --ui-pad-xs: 4px; --ui-pad-sm: 8px; --ui-pad-lg: 18px; --ui-radius: 10px; --ui-dur-slow: .3s;
  --color-fg: #1f1f1e; --color-fg-muted: #6b6a66; --color-bg: #faf9f5; --color-panel: #ffffff; --color-border-line: #e6e3da;
  --color-ok: #2f7d32; --color-warn: #a15c07; --color-bad: #b3261e;
  --ui-chart-grid: #efece4; --ui-chart-axis: #cfcbc0; --ui-chart-reference: #8a877e; --ui-chart-muted: #b9b5aa; --ui-chart-status-critical: #c0392b;
  --ui-surface-popover: #ffffff; --ui-border: #e6e3da; --ui-shadow-popover: 0 4px 14px rgba(0,0,0,.12); --ui-text-secondary: #6b6a66; }
@media (prefers-color-scheme: dark) { :root {
  --color-fg: #ecebe6; --color-fg-muted: #a3a19a; --color-bg: #1c1c1b; --color-panel: #262624; --color-border-line: #3a3936;
  --color-ok: #6fbf73; --color-bad: #f0857c; --ui-chart-grid: #33322f; --ui-chart-axis: #55534e; --ui-chart-reference: #8a877e;
  --ui-chart-muted: #6d6a62; --ui-chart-status-critical: #e5695e; --ui-surface-popover: #2e2e2b; --ui-border: #44433f; --ui-text-secondary: #a3a19a; } }
body { margin: 0; background: var(--color-bg); color: var(--color-fg); }
#dash-root { padding: 24px 16px; }
.dash-skeleton { color: transparent !important; background: var(--ui-chart-grid); border-radius: 6px; }
"""

# Data layer: serves the embedded CSV rows to the page and runs its calculated tables.
DATA_LAYER = """
const RAW = %s;
const calcs = {}, listeners = [], cache = {};
window.dash = {
  colors: ['#C2410C', '#2563EB', '#15803D', '#7C3AED', '#A16207', '#0E7490', '#BE185D', '#4D7C0F'],
  calc(id, o) { calcs[id] = o; },
  loader() {},
  data(id) {
    if (cache[id]) return cache[id];
    if (RAW[id] != null) return (cache[id] = { status: 'ok', data: d3.csvParse(RAW[id]), refreshing: false, meta: [] });
    const c = calcs[id];
    if (!c) return { status: 'missing', data: [], meta: [] };
    try { return (cache[id] = { status: 'ok', data: c.fn(...c.inputs.map(i => this.data(i).data)), refreshing: false, meta: [] }); }
    catch (e) { console.error(id, e); return { status: 'error', data: [], message: String(e), meta: [] }; }
  },
  onData(fn) { listeners.push(fn); },
  setLink(name) { history.replaceState(null, '', '#' + name); },
  params() { return {}; }, setParams() {}, resetParams() {}, refresh() {}
};
let t;
const run = () => listeners.forEach(fn => fn());
addEventListener('load', run);
addEventListener('resize', () => { clearTimeout(t); t = setTimeout(run, 150); });
"""

# ---- Downloads: the Power BI project (zip) and the raw Excel, embedded so the buttons work offline ----
ZIP_NAME = "FoodAffordabilityPH_PowerBI.zip"
bundle = ["powerbi", "raw/ph_food_economy_raw.xlsx", "theme.json", "logo.svg", "measures.dax", "power_query.m", "README.md"]
buf = io.BytesIO()
with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
    for item in bundle:
        src = ROOT / item
        assert src.exists(), f"missing {item} - run build_pbip.py / extract_data.py first"
        for f in sorted([src] if src.is_file() else (p for p in src.rglob("*") if p.is_file())):
            info = zipfile.ZipInfo(f"food-affordability-powerbi/{f.relative_to(ROOT).as_posix()}", (2026, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, f.read_bytes())
zip_bytes = buf.getvalue()
assert zipfile.ZipFile(io.BytesIO(zip_bytes)).testzip() is None
assert any(n.endswith(".pbip") for n in zipfile.ZipFile(io.BytesIO(zip_bytes)).namelist())
(ROOT / "downloads").mkdir(exist_ok=True)
(ROOT / "downloads" / ZIP_NAME).write_bytes(zip_bytes)
xlsx_bytes = (ROOT / "raw" / "ph_food_economy_raw.xlsx").read_bytes()
files = {"pbip": (ZIP_NAME, "application/zip", zip_bytes),
         "xlsx": ("ph_food_economy_raw.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", xlsx_bytes)}

DOWNLOAD_BAR = """
<div class="dl-bar" role="group" aria-label="Downloads">
  <button class="dl-btn dl-primary" data-file="pbip">
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4v11m0 0-4-4m4 4 4-4M5 19h14"/></svg>
    Download Power BI project (.pbip) <span class="dl-size">%s</span></button>
  <button class="dl-btn" data-file="xlsx">
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4v11m0 0-4-4m4 4 4-4M5 19h14"/></svg>
    Download raw data (Excel) <span class="dl-size">%s</span></button>
  <span class="dl-note">Unzip, then open <b>powerbi/FoodAffordabilityPH.pbip</b> in Power BI Desktop (Windows). Steps are in README.md.</span>
</div>"""
kb = lambda b: f"{max(1, round(len(b) / 1024))} KB"
DOWNLOAD_CSS = """
.dl-bar { display: flex; flex-wrap: wrap; align-items: center; gap: 8px 10px; margin: 0 0 var(--ui-gap-md); font-family: var(--font-sans); }
.dl-btn { display: inline-flex; align-items: center; gap: 6px; font: inherit; font-size: 14px; color: var(--color-fg); background: var(--color-panel);
  border: 1px solid var(--color-border-line); border-radius: 8px; padding: 7px 12px; cursor: pointer; }
.dl-btn:hover { border-color: var(--color-fg-muted); }
.dl-btn:focus-visible { outline: 2px solid #C2410C; outline-offset: 2px; }
.dl-primary { background: #C2410C; border-color: #C2410C; color: #fff; }
.dl-primary:hover { background: #9A3412; border-color: #9A3412; }
.dl-btn svg { width: 16px; height: 16px; fill: none; stroke: currentColor; stroke-width: 2; stroke-linecap: round; stroke-linejoin: round; }
.dl-size { opacity: .75; font-size: 12px; }
.dl-note { font-size: 12px; color: var(--color-fg-muted); }
"""
DOWNLOAD_JS = """
const FILES = %s;
document.addEventListener('click', e => {
  const b = e.target.closest('.dl-btn');
  if (!b) return;
  const [name, type, data] = FILES[b.dataset.file];
  const bytes = Uint8Array.from(atob(data), c => c.charCodeAt(0));
  const url = URL.createObjectURL(new Blob([bytes], { type }));
  const a = Object.assign(document.createElement('a'), { href: url, download: name });
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
});
""" % json.dumps({k: [n, t, base64.b64encode(b).decode()] for k, (n, t, b) in files.items()})

anchor = "same pages, measures and slicers as the Power BI report</p>"
assert page.count(anchor) == 1, "download bar anchor not found in web-dashboard/index.html"
page = page.replace(anchor, anchor + DOWNLOAD_BAR % (kb(zip_bytes), kb(xlsx_bytes)))

html = f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Food Affordability PH 2026</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/d3/7.9.0/d3.min.js"></script>
<style>{TOKENS}{DOWNLOAD_CSS}</style>
<script>{DATA_LAYER % json.dumps(csv)}</script>
</head><body><div id="dash-root">
{page}
</div>
<script>{DOWNLOAD_JS}</script>
</body></html>
"""
(ROOT / "dashboard.html").write_text(html, encoding="utf-8")
print(f"wrote dashboard.html ({len(html) // 1024} KB) with", ", ".join(f"{k}: {len(v.splitlines()) - 1} rows" for k, v in csv.items()))
