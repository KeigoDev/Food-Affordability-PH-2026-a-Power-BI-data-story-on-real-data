// STAGE 2 - CLEAN & TRANSFORM (Power Query)
// How to use: Power BI Desktop > Transform data > Home > New Source > Blank Query >
// Advanced Editor > paste ONE block below > Done > rename the query to the name in its header.
// Create them in this order. Every query reads the raw workbook made by extract_data.py.


// ===== Query 1: RawFile  (a parameter: Home > Manage Parameters > New, Type = Text) =====
// Current value = full path to raw\ph_food_economy_raw.xlsx, e.g.
// C:\Users\you\Documents\food-affordability-powerbi\raw\ph_food_economy_raw.xlsx


// ===== Query 2: fnToNumber  (helper function) =====
// PSA writes "." or ".." for figures not yet released; those become null, not errors.
(value as any) as nullable number =>
let
    t = if value = null then "" else Text.Trim(Text.From(value, "en-US")),
    result = if t = "" or t = "." or t = ".." then null else Number.FromText(t, "en-US")
in
    result


// ===== Query 3: Region =====
let
    Source = Excel.Workbook(File.Contents(RawFile), null, true),
    Sheet = Source{[Item = "region_lookup", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"region_key", type text}, {"region", type text}, {"island_group", type text},
        {"rice_label", type text}, {"cpi_label", type text}})
in
    Typed


// ===== Query 4: RicePrice  (PSA retail price, regular-milled rice, PHP per kg) =====
let
    Source = Excel.Workbook(File.Contents(RawFile), null, true),
    Sheet = Source{[Item = "psa_rice_prices", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    // wide (one column per month) -> long (one row per region and month)
    Unpivoted = Table.UnpivotOtherColumns(Promoted, {"Geolocation", "Commodity"}, "Period", "RawValue"),
    NoAnnual = Table.SelectRows(Unpivoted, each not Text.EndsWith([Period], "Annual")),
    Numbers = Table.AddColumn(NoAnnual, "price_php_kg", each fnToNumber([RawValue]), type nullable number),
    NoMissing = Table.SelectRows(Numbers, each [price_php_kg] <> null),
    // "2025 January" -> 2025-01-01
    Month = Table.AddColumn(NoMissing, "month", each Date.FromText(
        "1 " & Text.AfterDelimiter([Period], " ") & " " & Text.BeforeDelimiter([Period], " "), [Culture = "en-US"]), type date),
    // "..NCR - National Capital Region" -> "NCR - National Capital Region"
    Label = Table.AddColumn(Month, "rice_label", each Text.TrimStart([Geolocation], "."), type text),
    Joined = Table.NestedJoin(Label, {"rice_label"}, Region, {"rice_label"}, "R", JoinKind.Inner),
    Expanded = Table.ExpandTableColumn(Joined, "R", {"region_key"}),
    Result = Table.SelectColumns(Expanded, {"region_key", "month", "price_php_kg"})
in
    Result


// ===== Query 5: Inflation  (PSA CPI, % change vs same month a year earlier) =====
let
    Source = Excel.Workbook(File.Contents(RawFile), null, true),
    Sheet = Source{[Item = "psa_inflation", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Unpivoted = Table.UnpivotOtherColumns(Promoted, {"Geolocation", "Commodity Description"}, "Period", "RawValue"),
    NoAverage = Table.SelectRows(Unpivoted, each not Text.EndsWith([Period], "Ave")),
    Numbers = Table.AddColumn(NoAverage, "yoy_pct", each fnToNumber([RawValue]), type nullable number),
    NoMissing = Table.SelectRows(Numbers, each [yoy_pct] <> null),
    Month = Table.AddColumn(NoMissing, "month", each Date.FromText(
        "1 " & Text.AfterDelimiter([Period], " ") & " " & Text.BeforeDelimiter([Period], " "), [Culture = "en-US"]), type date),
    Item = Table.AddColumn(Month, "item", each
        if [Commodity Description] = "0 - ALL ITEMS" then "All items"
        else if [Commodity Description] = "01.1 - FOOD" then "Food"
        else if [Commodity Description] = "01.1.1.12 - Rice" then "Rice"
        else null, type text),
    KnownItems = Table.SelectRows(Item, each [item] <> null),
    Label = Table.AddColumn(KnownItems, "cpi_label", each Text.TrimStart([Geolocation], "."), type text),
    // inner join also drops "Areas Outside NCR", which is not a region
    Joined = Table.NestedJoin(Label, {"cpi_label"}, Region, {"cpi_label"}, "R", JoinKind.Inner),
    Expanded = Table.ExpandTableColumn(Joined, "R", {"region_key"}),
    Result = Table.SelectColumns(Expanded, {"region_key", "month", "item", "yoy_pct"})
in
    Result


// ===== Query 6: PesoPower  (PSA purchasing power of the peso, 2018 = 1.00) =====
let
    Source = Excel.Workbook(File.Contents(RawFile), null, true),
    Sheet = Source{[Item = "psa_peso_power", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Unpivoted = Table.UnpivotOtherColumns(Promoted, {"Geolocation", "Description"}, "Period", "RawValue"),
    NoAverage = Table.SelectRows(Unpivoted, each not Text.EndsWith([Period], "Ave")),
    Numbers = Table.AddColumn(NoAverage, "peso_power", each fnToNumber([RawValue]), type nullable number),
    NoMissing = Table.SelectRows(Numbers, each [peso_power] <> null),
    Month = Table.AddColumn(NoMissing, "month", each Date.FromText(
        "1 " & Text.AfterDelimiter([Period], " ") & " " & Text.BeforeDelimiter([Period], " "), [Culture = "en-US"]), type date),
    Result = Table.SelectColumns(Month, {"month", "peso_power"})
in
    Result


// ===== Query 7: GDP  (PSA real GDP growth, % vs same quarter a year earlier) =====
let
    Source = Excel.Workbook(File.Contents(RawFile), null, true),
    Sheet = Source{[Item = "psa_gdp_growth", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Unpivoted = Table.UnpivotOtherColumns(Promoted, {"Type of Expenditure"}, "Header", "RawValue"),
    Numbers = Table.AddColumn(Unpivoted, "gdp_growth_pct", each fnToNumber([RawValue]), type nullable number),
    NoMissing = Table.SelectRows(Numbers, each [gdp_growth_pct] <> null),
    // "At Constant 2018 Prices 2025-2026 Q2" -> year 2026, quarter 2
    Year = Table.AddColumn(NoMissing, "year", each Number.FromText(Text.Middle([Header], Text.Length([Header]) - 7, 4)), Int64.Type),
    Qtr = Table.AddColumn(Year, "qtr", each Number.FromText(Text.End([Header], 1)), Int64.Type),
    QuarterStart = Table.AddColumn(Qtr, "quarter_start", each #date([year], ([qtr] - 1) * 3 + 1, 1), type date),
    Quarter = Table.AddColumn(QuarterStart, "quarter", each Text.From([year]) & " Q" & Text.From([qtr]), type text),
    Result = Table.SelectColumns(Quarter, {"quarter", "quarter_start", "gdp_growth_pct"})
in
    Result


// ===== Query 8: FX  (ECB monthly average; USD/PHP = PHP per EUR / USD per EUR) =====
let
    Source = Excel.Workbook(File.Contents(RawFile), null, true),
    Sheet = Source{[Item = "ecb_fx", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Kept = Table.SelectColumns(Promoted, {"CURRENCY", "TIME_PERIOD", "OBS_VALUE"}),
    Numbers = Table.TransformColumns(Kept, {{"OBS_VALUE", fnToNumber, type nullable number}}),
    Pivoted = Table.Pivot(Numbers, {"PHP", "USD"}, "CURRENCY", "OBS_VALUE"),
    Complete = Table.SelectRows(Pivoted, each [PHP] <> null and [USD] <> null),
    UsdPhp = Table.AddColumn(Complete, "usd_php", each Number.Round([PHP] / [USD], 2), type number),
    Month = Table.AddColumn(UsdPhp, "month", each Date.FromText([TIME_PERIOD] & "-01"), type date),
    Result = Table.SelectColumns(Month, {"month", "usd_php"})
in
    Result


// ===== Query 9: Hunger  (SWS involuntary hunger, % of families) =====
let
    Source = Excel.Workbook(File.Contents(RawFile), null, true),
    Sheet = Source{[Item = "sws_hunger", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Typed = Table.TransformColumns(Promoted, {
        {"fieldwork_start", each Date.From(_), type date},
        {"hunger_pct", fnToNumber, type number}}),
    Month = Table.AddColumn(Typed, "month", each Date.StartOfMonth([fieldwork_start]), type date),
    Result = Table.SelectColumns(Month, {"survey", "month", "hunger_pct", "source"})
in
    Result


// ===== Query 10: Wage  (NCR daily minimum wage, non-agriculture) =====
let
    Source = Excel.Workbook(File.Contents(RawFile), null, true),
    Sheet = Source{[Item = "ncr_min_wage", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Typed = Table.TransformColumns(Promoted, {
        {"effective_date", each Date.From(_), type date},
        {"daily_wage_php", fnToNumber, type number}}),
    Result = Table.SelectColumns(Typed, {"wage_order", "effective_date", "daily_wage_php", "source"})
in
    Result


// ===== Query 11: Events  (rice policy events, for chart annotations) =====
let
    Source = Excel.Workbook(File.Contents(RawFile), null, true),
    Sheet = Source{[Item = "rice_policy_events", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Typed = Table.TransformColumns(Promoted, {{"date", each Date.From(_), type date}, {"event", Text.From, type text}})
in
    Typed
