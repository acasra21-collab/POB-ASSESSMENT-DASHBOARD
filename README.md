# POB Collection & Import Analytics — Port of Batangas (Assessment Division)

An interactive Flask dashboard for the BOC **Port of Batangas** masterlist. The
raw `.xlsx` is parsed **once on upload** (openpyxl, pure Python — no Streamlit /
DuckDB / pandas) and held in memory; the browser then drives a **live** dashboard
through JSON API endpoints — a date-range calendar and oil/port filters that
re-aggregate every tab on demand, plus click-through drill-downs that reveal the
underlying **declaration (entry) numbers**.

## Run

```
py app.py
```
Open http://127.0.0.1:8060, upload the masterlist `.xlsx`, and the dashboard
appears. Parsing the full ~87 MB / ~220k-row file takes ~60–80 s (once per
upload); after that every tab, filter, and drill-down responds live.

Or double-click **`run.bat`** (installs deps, opens the browser).

> The parsed data lives in memory for as long as the server runs. If you restart
> the app you re-upload the file. (The dashboard's **💾 Save snapshot** button
> exports the current on-screen view as a standalone HTML file.)

## Counts are by declaration, not line item

Revenue counts use **distinct `ENTRY_CODE`** (one customs declaration), not raw
spreadsheet rows. The 2026 file has **219,928 line items but only 14,286
declarations** (one declaration can carry hundreds of line items). Everything
labelled "Declarations" / "Shipments" is a distinct entry-number count.

## Inputs (upload form)

| Field | Required | Notes |
|---|---|---|
| Masterlist (current) | ✅ | BOC consumption `.xlsx`, Sheet1 |
| Date basis | — | COLLECTIONDATE (default) / ASSESSMENTDATE / REGISTRYDATE |
| Prior-year masterlist | optional | enables YoY columns |
| DBCC targets CSV | optional | `year,month,monthly_target_php`; defaults to bundled `targets.csv` |

## Metric choices

* **Revenue** = `TOTALASSESSMENT` (PHP duties + taxes)
* **Volume** = `GROSSMASS` (kg; shown as MT for petroleum)
* **Importer** = `CONSIGNEE`, normalized (UPPER + collapsed whitespace)
* **Declaration** = distinct `ENTRY_CODE`

## Filter bar (applies to every tab)

* **Period calendar** — pick any start/end date, or a quick preset (Full span /
  YTD / Latest month / Last 30 days). Click **Apply ⟳** to re-aggregate.
* **Oil** — Both / Oil / Non-oil.
* **Port** — toggle any of P04 / P04A / P04B / P04C.

## Tabs

1. **Overview** — totals (incl. distinct declarations), monthly revenue vs DBCC target, oil pie, daily revenue, top-10 importers/commodities.
2. **Daily Collection** — selected-period & YTD revenue vs prorated DBCC target, cumulative line, revenue components, YoY.
3. **Oil / Non-oil** — revenue & volume split with MoM + YoY.
4. **Importers** — top importers with MoM/YoY; **click a row → its products**; **▦ entry nos** on any row/product shows that selection's declarations.
5. **Commodities** — INDEX2 / GENERAL DESC trends + two-way drill (**commodity → importers** / **importer → products**), each with entry-number views.
6. **🚗 Car Imports** — two sub-tabs:
   - **⚡ Powertrain (cars):** Electric / Hybrid / ICE mix **classified by HS code** (ICE = 8703.21–8703.33 · Hybrid = 8703.40–8703.70 · Electric = 8703.80 + 8704.60), units by port, powertrain filter, top importers → each model → **▦ entry nos**. "Other vehicles" = buses/pickups/trucks not in the car-powertrain codes.
   - **🛻 Pickups & Trucks:** goods vehicles (HS 8704) classified by the official HS-code lists (`PICKUP_HS` / `TRUCK_HS` in `app.py`), with units/declarations/duties by class, by port, and top importers. Electric goods vehicles (8704.60) are flagged.
7. **⛽ Petroleum** — totals in MT, by category → products, by port, top importers → products, all with entry-number views.
8. **📝 Memo** — the formal BOC **"Daily Collection Report and Assessment Import Analysis"** memorandum (sections A–F) computed for the selected period, in the same layout/wording as the official template (header officials, A daily collection vs DBCC, B cumulative, C vs last year, Daily Revenue Contributor, D.1/D.2 volume analytics, E top-30 importers, F top-30 commodities). A **"Generate Word memo (.docx)"** button downloads it as a `.docx` with styled tables and the **original letterhead header + footer** (carried from `memo_template.docx`, a stripped copy of the official memo). The C / D.2 / "vs last year" sections need a prior-year file; DBCC figures come from the loaded targets CSV. Officials' names are set in `memo_builder.py` (`OFFICIALS`); the header/footer come from `memo_template.docx` — replace that file to change the letterhead.

## Drill-downs & entry numbers

Anywhere you see **▦ entry nos**, clicking it opens a modal listing the distinct
declaration numbers for that selection (importer, product, commodity, port, or a
specific item within a parent) with date, port, importer, revenue, and volume.

## Business rules (pure Python in `app.py`)

* **Oil** if `validity` = `OIL DV` or `INDEX2` ∈ {GASOLINE, DIESEL, JET FUEL,
  LUBRICATING OIL, OTHER OIL PRODUCTS, LPG, BUNKER FUEL}.
* **Vehicle** if HS heading ∈ {8702, 8703, 8704} or a motor-vehicle INDEX2 bucket;
  units = `QUANTITY` for CBU rows (recovered from "<n> UNITS" text when 0); CKD
  parts excluded from unit counts.
* **Powertrain** (Hybrid → Electric → ICE) inferred from the goods description +
  brand (incl. EV-only brands: Tesla, BYD, VinFast, …).
