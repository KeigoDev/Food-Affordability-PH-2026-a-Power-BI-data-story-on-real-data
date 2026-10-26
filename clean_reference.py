"""Reference copy of the Power Query cleaning (power_query.m) in Python.
Uses: (1) the clean CSVs that feed the web dashboard, and (2) expected values to
compare with Power BI after STAGE 2 (see README, stage 6).
Run after extract_data.py:  python3 clean_reference.py"""
import pathlib
import pandas as pd

ROOT = pathlib.Path(__file__).parent
raw = pd.read_excel(ROOT / "raw" / "ph_food_economy_raw.xlsx", sheet_name=None, dtype=str)
out = ROOT / "data"
out.mkdir(exist_ok=True)

num = lambda s: pd.to_numeric(s.str.strip().replace({".": None, "..": None, "": None}), errors="raise")
region = raw["region_lookup"]


def long_monthly(df, id_cols, drop_suffix, label_col, value_name):
    d = df.melt(id_vars=id_cols, var_name="period", value_name="raw")
    d = d[~d["period"].str.endswith(drop_suffix)].copy()
    d[value_name] = num(d["raw"])
    d = d.dropna(subset=[value_name])
    d["month"] = pd.to_datetime("1 " + d["period"].str.split(" ").str[1] + " " + d["period"].str.split(" ").str[0], format="%d %B %Y" if drop_suffix == "Annual" else "%d %b %Y")
    if label_col:
        d[label_col] = d[id_cols[0]].str.lstrip(".")
        d = d.merge(region[["region_key", label_col]], on=label_col, how="inner")
    return d


rice = long_monthly(raw["psa_rice_prices"], ["Geolocation", "Commodity"], "Annual", "rice_label", "price_php_kg")
infl = long_monthly(raw["psa_inflation"], ["Geolocation", "Commodity Description"], "Ave", "cpi_label", "yoy_pct")
infl["item"] = infl["Commodity Description"].map({"0 - ALL ITEMS": "All items", "01.1 - FOOD": "Food", "01.1.1.12 - Rice": "Rice"})
power = long_monthly(raw["psa_peso_power"], ["Geolocation", "Description"], "Ave", None, "peso_power")

g = raw["psa_gdp_growth"].melt(id_vars=["Type of Expenditure"], var_name="header", value_name="raw")
g["gdp_growth_pct"] = num(g["raw"])
g = g.dropna(subset=["gdp_growth_pct"])
g["year"] = g["header"].str[-7:-3].astype(int)
g["qtr"] = g["header"].str[-1].astype(int)
g["quarter"] = g["year"].astype(str) + " Q" + g["qtr"].astype(str)
g["quarter_start"] = pd.to_datetime(dict(year=g["year"], month=(g["qtr"] - 1) * 3 + 1, day=1))

fx = raw["ecb_fx"].copy()
fx["v"] = num(fx["OBS_VALUE"])
provisional = set(fx.loc[fx["TITLE"].str.startswith("PROVISIONAL"), "TIME_PERIOD"])
fx_note = fx.loc[fx["TITLE"].str.startswith("PROVISIONAL"), "TITLE"].head(1).tolist()
fx = fx.pivot(index="TIME_PERIOD", columns="CURRENCY", values="v").dropna()
fx = pd.DataFrame({"month": pd.to_datetime(fx.index + "-01"), "usd_php": (fx["PHP"] / fx["USD"]).round(2).values,
                   "provisional": [p in provisional for p in fx.index]})

hunger = raw["sws_hunger"].copy()
hunger["hunger_pct"] = num(hunger["hunger_pct"])
hunger["month"] = pd.to_datetime(hunger["fieldwork_start"]).dt.to_period("M").dt.to_timestamp()
wage = raw["ncr_min_wage"].copy()
wage["effective_date"] = pd.to_datetime(wage["effective_date"])
wage["daily_wage_php"] = num(wage["daily_wage_php"])
events = raw["rice_policy_events"]

# ---- web dashboard tables (flattened) ----
ym = lambda s: s.dt.strftime("%Y-%m-%d")
pick = lambda key, item: infl[(infl.region_key == key) & (infl["item"] == item)].set_index("month")["yoy_pct"]
ph = rice[rice.region_key == "PH"].set_index("month")["price_php_kg"]
ncr = rice[rice.region_key == "NCR"].set_index("month")["price_php_kg"]
nat = pd.DataFrame({"rice_price_php_kg": ph, "ncr_rice_price_php_kg": ncr,
                    "headline_inflation_pct": pick("PH", "All items"), "food_inflation_pct": pick("PH", "Food"),
                    "rice_inflation_pct": pick("PH", "Rice"),
                    "peso_power": power.set_index("month")["peso_power"],
                    "usd_php": fx.set_index("month")["usd_php"]}).sort_index()
nat = nat.dropna(subset=["rice_price_php_kg", "headline_inflation_pct"])  # months PSA has released
# wage in force on the 15th of each month (wage orders take effect mid-month)
nat["ncr_min_wage_php"] = [wage[wage.effective_date <= m + pd.Timedelta(days=14)]["daily_wage_php"].iloc[-1] for m in nat.index]
nat["kg_rice_per_ncr_wage"] = (nat["ncr_min_wage_php"] / nat["ncr_rice_price_php_kg"]).round(1)
nat = nat.reset_index().rename(columns={"index": "month"})
nat["month"] = ym(nat["month"])

latest = rice[rice.region_key != "PH"]["month"].max()
year_ago = latest - pd.DateOffset(years=1)
r_now = rice[rice.month == latest].set_index("region_key")["price_php_kg"]
r_ago = rice[rice.month == year_ago].set_index("region_key")["price_php_kg"]
i_now = infl[infl.month == latest]
regions = region[region.region_key != "PH"][["region_key", "region", "island_group"]].set_index("region_key")
regions["rice_price_php_kg"] = r_now
regions["rice_price_year_ago"] = r_ago
regions["rice_price_change_pct"] = ((r_now / r_ago - 1) * 100).round(1)
for item, col in (("Rice", "rice_inflation_pct"), ("Food", "food_inflation_pct"), ("All items", "headline_inflation_pct")):
    regions[col] = i_now[i_now["item"] == item].set_index("region_key")["yoy_pct"]
regions = regions.reset_index()

gdp = g[["quarter", "quarter_start", "gdp_growth_pct"]].sort_values("quarter_start")
gdp["quarter_start"] = ym(gdp["quarter_start"])
hunger_out = hunger[["survey", "month", "hunger_pct"]].assign(month=ym(hunger["month"]))

# ---- self-checks: catch a broken source or join before anything is shared ----
assert len(nat) >= 24 and nat.isna().sum().sum() == 0, nat.isna().sum()
assert regions.notna().all().all() and len(regions) == 18, regions[regions.isna().any(axis=1)]
assert nat["usd_php"].between(40, 90).all() and nat["rice_price_php_kg"].between(20, 100).all()
assert gdp["gdp_growth_pct"].between(-20, 20).all() and len(gdp) >= 12

# FX alone runs ahead of PSA: it can include the current month (provisional, from ECB daily rates).
fx_out = fx.sort_values("month").assign(month=lambda d: ym(d["month"]), provisional=lambda d: d["provisional"].astype(int))
assert fx_out["usd_php"].between(40, 90).all() and not fx_out["month"].duplicated().any()
assert fx_out["provisional"].iloc[:-1].eq(0).all(), "only the newest FX month may be provisional"
assert set(nat["month"]) <= set(fx_out["month"])

nat.to_csv(out / "national_monthly.csv", index=False)
fx_out.to_csv(out / "fx_monthly.csv", index=False)
if fx_note: print("FX", fx_out["month"].iloc[-1], fx_out["usd_php"].iloc[-1], "-", fx_note[0])
regions.to_csv(out / "regions_latest.csv", index=False)
gdp.to_csv(out / "gdp_quarterly.csv", index=False)
hunger_out.to_csv(out / "hunger_sws.csv", index=False)
events.to_csv(out / "rice_policy_events.csv", index=False)
print("latest month:", latest.date(), "| rows:", len(nat), len(regions), len(gdp), len(hunger_out), len(events))
