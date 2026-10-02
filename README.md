# Merchandising & Product Profitability — Contoso Retail Group

**Power BI capstone project** — a full business intelligence build for a consumer-electronics retailer trading across 8 countries, 73 stores, and online/click-and-collect channels. 121,160 sales transactions, 9,919 returns, cleaned, modeled, and turned into a two-page decision-ready report.

![Product Profitability](Product%20Profitability.jpeg)
![Discounting & Returns](Discounting%20Returns.jpeg)

---

## Table of Contents

1. [Business Context](#business-context)
2. [Project Structure](#project-structure)
3. [Task 1 — Data Cleaning (Power Query)](#task-1--data-cleaning-power-query)
4. [Task 2 — Data Model](#task-2--data-model)
5. [Task 3 — DAX Measures](#task-3--dax-measures)
6. [Task 4 — Report](#task-4--report)
7. [What I actually learned](#what-i-actually-learned)
8. [Tools](#tools)
9. [Repository Contents](#repository-contents)
10. [Author](#author)

---

## Business Context

Contoso Retail Group is a consumer-electronics retailer trading in **8 countries** (United States, Canada, United Kingdom, Germany, France, Italy, Netherlands, Australia) through **73 physical stores**, an online shop, and a Click & Collect service. The data covers **1 January 2015 to 31 March 2024**, extracted on 30 April 2024.

Three business rules shaped almost every decision in this project:

1. **Money is in local currency.** Every amount in `sales` and `returns` is recorded in the currency the store trades in — a German store records euros, an Australian store records Australian dollars. To compare or add amounts across countries, every figure has to be converted: `USD amount = local amount ÷ ExchangeRate`. Prices in `product` are already in USD.
2. **Two fact tables, no join between them.** `sales` and `returns` sit at different grains and must never be directly related — they only meet through shared dimensions (Date, store, product).
3. **A return is not a cancelled sale.** A returned line still exists in `sales` — the goods were bought and paid for. `returns` records what came back afterward, sometimes only part of a line. Revenue and returned value are reported separately, never subtracted from each other at the row level — only `Net Revenue` combines them, deliberately, as a measure.

---

## Project Structure

The source workbook contains 8 tables matching this structure:

| Table | Type | Rows | Key columns |
|---|---|---|---|
| `sales` | Fact | 121,160 | OrderKey, LineNumber, ProductKey, StoreKey, PromotionKey |
| `returns` | Fact | 9,919 | ReturnKey, OrderKey, LineNumber, ProductKey, StoreKey |
| `product` | Dimension | 2,517 | ProductKey, ManufacturerKey |
| `manufacturer` | Dimension | 11 | ManufacturerKey |
| `store` | Dimension | 73 | StoreKey, GeoAreaKey, CurrencyCode |
| `geography` | Dimension | 595 | GeoAreaKey |
| `promotion` | Dimension | 65 | PromotionKey |
| `currency` | Dimension | 5 | CurrencyCode |

A Date table is not provided in the source and had to be built from scratch (see Task 2).

---

## Task 1 — Data Cleaning (Power Query)

The rule I followed throughout: **never trust a table because it looks fine at a glance.** Every table was checked against its documented row count, its expected distinct-value count, and the business rules in the data dictionary — on the *entire* dataset, not the default 1,000-row preview Power Query shows by default.

### `geography`

- **Problem:** The `Continent` column had **12 inconsistent spellings for only 3 real continents** — case variants (`Oceania` / `OCEANIA` / `oceania`, `NORTH AMERICA` / `North America` / `north america`) and an abbreviation (`N. America`) that casing fixes alone couldn't catch.
- **Fix:** Trim → Clean (strips invisible/hidden characters) → Capitalize Each Word → Replace Values (`"N. America"` → `"North America"`).
- **Verification:** Grouped by `Continent` + `CountryCode` and confirmed exactly **8 rows** — one continent per country, with no country split across two spellings. Final column profile: 3 distinct values, 595 rows intact.

### `currency`

- **Problem:** The column headers (`CurrencyCode`, `CurrencyName`, `Symbol`, `PrimaryMarket`) had been loaded as a data row instead of being promoted to headers — the table's real first row was data labeled `Column1`–`Column4`.
- **Fix:** Removed the auto-generated "Changed Type" step that had locked onto the wrong headers, applied "Use First Row as Headers", then set all 4 columns to Text.
- **Verification:** 5 distinct/unique `CurrencyCode` values (USD, CAD, AUD, GBP, EUR), matching the 5 currencies used across `store`.

### `manufacturer`

- **Problem:** A single `Supplier` column held both the manufacturer name and its country combined in one text field (e.g. `"Fabrikam, Inc. (Germany)"`), with a trailing parenthesis that needed cleanup.
- **Fix:** Split the column into `ManufacturerName` and `SupplierCountry`, cleaned the trailing parenthesis left behind by the split, renamed both resulting columns, and set their data types.
- **Verification:** 11 distinct/unique manufacturers, `BrandTier` (Value / Mainstream / Premium / Own Brand) and `IsOwnBrand` (TRUE only for Contoso, Ltd) both intact and matching the data dictionary.

### `promotion`

- **Checked, no fix needed.** `PromotionKey = 0` ("No Promotion") has blank `StartDate`/`EndDate` — this is correct by design, not missing data: it's the valid key assigned to every undiscounted sales line, not a blank row to be deleted. Verified 65 distinct/unique `PromotionKey`, 4 `PromotionType` values, and confirmed `CategoryScope` text (e.g. `"Cell Phones"`) matches `product[CategoryName]` casing exactly so the two tables can be compared reliably downstream.

### `store`

- **Checked, no fix needed.** `CloseDate` is blank on 78% of rows — confirmed correct, since blank means the store is still trading. 73 distinct/unique `StoreKey`, matching the documented store count exactly.

### `product`

- **Problem 1 — uninformative columns:** Two leftover columns, `LegacyRef` and `SourceFile`, carried no analytical value (internal export artifacts) and were removed.
- **Problem 2 — exact duplicate rows:** **5 `ProductKey` values (106, 151, 256, 611, 1170) each appeared twice**, as full-row duplicates — every column (name, manufacturer, category, color, weight, cost, price) identical between the pair. Confirmed by filtering to "Keep Duplicates" on `ProductKey`, sorting, and comparing every column side by side before touching anything, specifically to rule out the alternative explanation (two *different* products accidentally sharing a key, which would need to be flagged, not deleted).
- **Fix:** Removed with `Remove Duplicates` applied to the **whole table** (every column selected via Ctrl+A), not just the key column — the only safe way to remove duplicates, since deduplicating on a single column discards rows blindly without checking whether the rest of the row actually matches.
- **Verification:** `ProductKey` went from 2,522 rows with duplicates to exactly **2,517 distinct/unique rows**, matching the data dictionary and — critically — making the column usable as the "one" side of a relationship, which it could not be while duplicated.

### `sales`

- **Problem:** `sales[Discount]` is a **rate**, not a percentage out of 100 — `0.15` means 15% off, `0` means no discount. This kind of column is a classic silent-failure trap: a value accidentally stored as `15` instead of `0.15` would inflate every discount-based measure by **100×**, and the error wouldn't be obvious on a chart — it would just look like a very aggressive (but plausible-looking) discount strategy.
- **Fix:** No data change was needed — the column was already correctly scaled (0–1) in this table. The fix here was procedural: verifying the value range *before* building `Discount Given`, `Discount Depth %`, and `% Revenue on Promotion` on top of it, rather than trusting the column by name alone.
- **Verification:** 121,160 rows, 0% Error across every column (OrderKey, LineNumber, dates, keys, prices, currency, discount, promotion link).

### `returns`

- **Problem 1 — junk header rows:** The sheet had **2 non-data rows before the real header** (a title row and an "Exported by" / period-reference row), which Power Query had initially tried to read as column headers.
- **Problem 2 — inconsistent return-reason text:** `ReturnReason` had entries that looked identical on screen but weren't — extra whitespace and hidden (non-printing) characters that silently broke grouping and filtering — plus one outright spelling inconsistency, `"Change Of Mind"`, that needed to be merged into the standard `"Changed Mind"`.
- **Problem 3 — uninformative columns:** `Notes` and `ExportedBy` carried no analytical value and were removed.
- **Fix, in order:** Removed the 2 junk header rows → promoted the real first row to headers → set correct data types → trimmed whitespace from `ReturnReason` → removed hidden characters from `ReturnReason` → standardized `ReturnReason` casing → merged `"Change Of Mind"` into `"Changed Mind"` → removed `Notes` and `ExportedBy`.
- **Verification:** `ReturnReason` reduced to exactly 5 clean, distinct values (Changed Mind, Faulty, Damaged, Wrong Item, Late Delivery), matching the data dictionary's documented 5 return reasons.
- **No relationship to `sales`, by design.** Confirmed no accidental relationship was ever created between the two fact tables — they are only connected through shared dimensions (Date, store, product), per business rule 2.

### General cleaning discipline applied across all 8 tables

- Every column profile was checked on the **entire dataset** ("Column profiling based on entire data set"), not the default 1,000-row preview — a column that looks 0% error on the preview can still hide problems further down.
- Every Applied Step was renamed from Power Query's defaults (`Source`, `Navigation`, `Promoted Headers`, `Changed Type`) to a name describing what it actually does — both because the task marks this explicitly, and because six months from now, "Changed Type" tells you nothing and "Set correct data types" does.
- No row was ever deleted without first proving it was a genuine duplicate (every column matching), per the task's explicit rule against unjustified row deletion.

---

## Task 2 — Data Model

Built a proper **star schema** from scratch — nothing in the source file is a working relational model on its own.

### Date table (DAX)

```dax
Date =
ADDCOLUMNS(
    CALENDAR(DATE(2015,1,1), DATE(2024,12,31)),
    "Year", YEAR([Date]),
    "Quarter", "Q" & QUARTER([Date]),
    "Month Number", MONTH([Date]),
    "Month Name", FORMAT([Date], "MMMM")
)
```

- Range extends through the end of 2024 (not just 31 March, the last sale date) so that any returns processed after the final sale still fall inside the calendar.
- `Month Name` sorted by `Month Number` (Sort by Column) so months display in calendar order, not alphabetically.
- Marked as the official Date Table in Power BI, on the `Date` column.
- `Year` and `Month Number` set to **Don't Summarize** — without this, dropping either into a visual sums the raw numbers (e.g. "4,074" instead of "2016"), which is meaningless.

### Relationships (10 total, all One-to-Many, Single direction, Active)

| From | To |
|---|---|
| `sales[OrderDate]` | `Date[Date]` |
| `returns[ReturnDate]` | `Date[Date]` |
| `sales[StoreKey]` | `store[StoreKey]` |
| `returns[StoreKey]` | `store[StoreKey]` |
| `sales[ProductKey]` | `product[ProductKey]` |
| `returns[ProductKey]` | `product[ProductKey]` |
| `product[ManufacturerKey]` | `manufacturer[ManufacturerKey]` |
| `store[GeoAreaKey]` | `geography[GeoAreaKey]` |
| `store[CurrencyCode]` | `currency[CurrencyCode]` |
| `sales[PromotionKey]` | `promotion[PromotionKey]` |

The deliberate omission: **no relationship between `sales` and `returns`.** Forcing one — even indirectly through `OrderKey` + `LineNumber` — would let a return "cancel out" a sale at the row level, contradicting business rule 3 and silently understating both gross revenue and the true number of transactions.

**Why `product` had to be deduplicated before this step:** Power BI refused to create the `product` → `returns` relationship until the duplicate `ProductKey` values were removed, since a one-to-many relationship requires the "one" side to be genuinely unique — this is the mechanism, not just a style rule, that caught the Task 1 duplicate issue.

---

## Task 3 — DAX Measures

All 15 measures live in a dedicated `_Measures` table (created with `_Measures = {BLANK()}`, with its placeholder `Value` column hidden), as required — never as calculated columns.

### The revenue ladder

Three of the 15 measures are the same money viewed at three points, not three unrelated numbers:

```
Gross Sales        (everything that left the shelf, at full ticket price)
  less Discount Given
= Total Revenue     (what the customer was actually charged)
  less Returned Value
= Net Revenue       (what Contoso kept)
```

### Full measure list

```dax
Total Revenue = SUMX(sales, sales[Quantity] * sales[NetPrice] / sales[ExchangeRate])

Total Cost = SUMX(sales, sales[Quantity] * sales[UnitCost] / sales[ExchangeRate])

Gross Profit = [Total Revenue] - [Total Cost]

Gross Margin % = DIVIDE([Gross Profit], [Total Revenue])

Total Units = SUM(sales[Quantity])

Average Selling Price = DIVIDE([Total Revenue], [Total Units])

Gross Sales = SUMX(sales, sales[Quantity] * sales[UnitPrice] / sales[ExchangeRate])

Discount Given = [Gross Sales] - [Total Revenue]

Discount Depth % = DIVIDE([Discount Given], [Gross Sales])

% Revenue on Promotion =
DIVIDE(CALCULATE([Total Revenue], sales[Discount] > 0), [Total Revenue])

Returned Units = SUM(returns[ReturnedQuantity])

Returned Value =
SUMX(returns, returns[ReturnedQuantity] * returns[NetPrice] / returns[ExchangeRate])

Return Rate % = DIVIDE([Returned Units], [Total Units])

Net Revenue = [Total Revenue] - [Returned Value]

Revenue YoY % =
VAR Curr = [Total Revenue]
VAR Prev = CALCULATE([Total Revenue], SAMEPERIODLASTYEAR('Date'[Date]))
RETURN IF(NOT ISBLANK(Curr), DIVIDE(Curr - Prev, Prev))
```

**Why every monetary measure uses `SUMX` with a per-row division, not `SUM` with a single conversion:** `ExchangeRate` varies row by row (a sale in euros has a different rate than one in Australian dollars on the same day), so the USD conversion has to happen at the row level, inside the iterator, before the sum — converting an already-summed total with one rate would silently produce a wrong number for every country except whichever currency happened to dominate the filter context.

**Why `Returned Value` reads `NetPrice`/`ExchangeRate` from `returns`, not `sales`:** There is no relationship between the two fact tables (see Task 2), so a `Returned Value` measure cannot reach into `sales` for pricing data — the `returns` table carries its own copy of these columns precisely so this measure can stand on its own.

**Validation performed against each measure, not just visually:**
- `Discount Given` was independently cross-checked against `SUMX(sales, sales[Quantity] * sales[UnitPrice] * sales[Discount] / sales[ExchangeRate])` — the two matched, confirming `NetPrice` really is "price after discount" and not something else.
- `Gross Sales > Total Revenue > Net Revenue` holds at every filter level, as the ladder requires.
- `Gross Margin %` lands between roughly 20–60% (consumer electronics range), not a nonsense value like a negative number or >100%.
- `Revenue YoY %` was checked year-by-year in a table visual: 2015 correctly blank (no prior year to compare), and 2024 intentionally reads as a large negative number — because the data only runs through March 2024, so the measure is comparing one quarter against a full prior year. This isn't a bug; it's flagged directly in the report wherever 2024 appears.

### Calculated columns (the 2 exceptions — these must be columns, not measures)

```dax
product[Initial Margin %] = DIVIDE(product[Price] - product[Cost], product[Price])

product[Weight in Grams] =
SWITCH(
    product[WeightUnit],
    "grams", product[Weight],
    "kilograms", product[Weight] * 1000,
    "ounces", product[Weight] * 28.3495
)
```

`Weight in Grams` exists because `product[Weight]` is **not additive across units** — a table of 500 grams and 2 kilograms cannot be summed as "502" without first converting both to the same unit, which is exactly what a per-row calculated column (not a measure) is for.

---

## Task 4 — Report

Two pages, one consistent non-default color theme, every number formatted in K/M (not raw digits), percentages to 1 decimal (`Discount Depth %` to 2, per spec).

### Page 1 — "Product Profitability"

| Visual | What it shows |
|---|---|
| 5 KPI cards | Total Revenue, Gross Profit, Gross Margin %, Total Units, Average Selling Price |
| Clustered bar | Top 10 subcategories by Gross Profit |
| Scatter chart | Gross Margin % vs. Return Rate % by category, bubble size = revenue — **one point per category** (8 bubbles, not one per product), with data labels instead of a legend so each bubble is readable without hovering |
| Matrix | Brand tier rows (Total Revenue, Gross Margin %, Return Rate %, Net Revenue), conditional formatting (red→green) on the margin column |
| Bar chart | Revenue by manufacturer, sorted highest to lowest |
| 3 slicers | Year, Category, Brand Tier |

### Page 2 — "Discounting & Returns"

| Visual | What it shows |
|---|---|
| 5 KPI cards | Discount Given, Discount Depth %, Returned Value, Return Rate %, Net Revenue |
| Clustered column | Gross Margin % by promotion type — makes the cost of each campaign type visible side by side (No Promotion 58.5% margin down to Flash 40.3%, an 18-point spread) |
| Bar chart | Top 10 subcategories by Returned Value |
| Donut chart | Returned units by return reason — **exactly 5 slices**, matching the cleaned `ReturnReason` values |
| Line + clustered column | Total Revenue (columns) and Gross Margin % (line) plotted against `Discount` rate directly on the X-axis — margin declines in a near-straight line as discount rate increases, from 58.6% at 0% discount to 40.4% at 30% |
| 3 slicers | Year, Manufacturer, Promotion Type |

### Rules specifically checked before calling the report done

- No donut/pie with more than 5 slices.
- No stacked chart with more than 4–5 series.
- The one dual-series chart (Total Revenue vs. Gross Margin %) uses two genuinely different units (currency vs. percentage), so it correctly uses a **secondary axis** — the "force both series onto one axis" rule only applies when both series share the *same* unit (e.g., this year's revenue vs. last year's), where a secondary axis would make two comparable lines cross at points that aren't real.

---

## What I actually learned

The hardest part of this project wasn't the charts — it was the roughly 80% of the work that happens before a single visual gets built:

- **"It looks fine" and "it is fine" are different claims**, and only one of them is checkable. A column can show 0% errors in a 1,000-row preview and still hide a problem in row 50,000.
- **A duplicate is only safe to remove once every column has been compared**, not just the key. Deduplicating on a single column can silently discard a row that was actually a different, valid record.
- **Business rules change what "correct" means.** The instinct to "fix" the Revenue YoY % 2024 number, or to join `sales` and `returns` for convenience, would both have produced technically-working formulas that quietly violated the data's actual structure.
- **Every number on a finished dashboard has to be traceable back to a decision you can explain** — not just a formula that compiles without error.

---

## Tools

Power Query (M) · DAX · Power BI data modeling · Power BI report design

---

## Repository Contents

| File | Description |
|---|---|
| `task 4.pbix` | Full Power BI file — cleaned data, model, measures, and both report pages |
| `Contoso Retail Data-v2.xlsx` | Source data |
| `Data Dictionary (Cheat Sheet)_after_edit.pdf` | Business rules and column definitions the project was built against |
| `Project 3 - Requirements_after_edit.pdf` | Original project brief |
| `Product Profitability.jpeg`, `Discounting Returns.jpeg` | Report page screenshots |

---

## Author

**Eslam Arafa** — Data Analyst
[LinkedIn](https://www.linkedin.com/in/eslam-m-arafa/) · [arafaeslam2003@gmail.com](mailto:arafaeslam2003@gmail.com)
