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
