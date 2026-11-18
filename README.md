# McCoy's Fleet Forecast 2027

Company-wide 2027 fleet cost forecast for Regions 01–08, Distribution, Millwork and Headquarters (every active asset except 950 Rio Truss), with a scenario tool to move, terminate or relocate assets.

Open `index.html` (or the GitHub Pages site). Scenario changes are kept in the viewer's browser only; use **Export to Excel** to save a scenario and **Import scenario** to load it back.

## Cost rules
- **Stores (Regions 01–08) and regional loaner pools (761–768):** fuel + flat rate + $0.425 per truck mile.
- **Actual cost (Distribution, Millwork, Headquarters, 355CD, 355 Reload):** fuel + maintenance + depreciation + lease + registration, taxes, Geotab and Lytx.
- Two views, switched from the header: location charges (rules above) and actual cost (real spend for every asset). Millwork 450 and Distribution 353/356/359/489 are on straight P&L, so they are the same in both views.

## Files
- `index.html`: the built dashboard.
- `data/data.json`: asset-level model inputs.
- `src/build_company.py`: builds `data.json` from the source exports (not in this repo).
- `src/head.html`, `src/body.html`: dashboard source; `src/build_html.py` assembles `index.html`.

Method and data flags are on the dashboard's **Assumptions & data** tab.

## Accounting review (Nov 2026)

`accounting-review/Fleet_2027_Budget_Variance_Review.xlsx` revises the 2027 maintenance, registration & inspection, and rental lines that accounting flagged against the FY2026 outlook ($15.67M). Revised 2027 total: **$15,163,631** (−$511K, −3.3% vs 2026), split Mfg/DC (353, 356, 359, 489, 450, 355CD, 355 Reload) $4.32M and flat rate $10.84M.

- Maintenance: 2026 outlook less replacement savings, plus store 122 and the new 359 forklift.
- Registration & inspections: held at 2026; ad valorem tax ($749K) belongs in the property-tax account.
- Rentals: Jan–Oct 2026 rental invoices (`accounting-review/inputs/2026_Rentals_Jan-Oct.xlsx`) × 12/10, less 011 South Austin (turned in), 950 Rio Truss and the 359 Burnet forklift rental (unit purchased).
- Fuel stays on the 2027 price curve.

Rebuild with `python3 accounting-review/variance_xlsx.py`. Not yet applied to the dashboard.

## Depreciation check against 2026 actuals (Nov 18)

Jan–Sep 2026 actual depreciation (BA depreciation detail) ran $1.145M/month in January down to $1.092M in September ($10.07M year to date). Assets on the roster match the forecast within about 1%. Added:

- **758002** (2027 International HV607, 24' TA Flatbed, 057 Tomball) and **755026** (2024 Kenworth T880 tractor, 355CD): on the depreciation books but missing from the roster export. Added as fleet units; registration, Geotab and Lytx use same-type medians.
- **77 book-only depreciation lines** (`inputs/book_depreciation_not_on_roster.csv`, from the Sep 2026 file): mostly Chameleon rolling tarps at 353/356/359, plus wraps/refurbs and units no longer on the roster. Actual expense only (no flat rate, fuel or maintenance); each runs to its estimated end date (start estimated from the BA number, 60/84-month life). 2027: $217,504.

2027 depreciation is now $11.92M (was $11.64M); actual cost $36.26M, location charges $35.96M.
