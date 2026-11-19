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

**Roster update (Nov 18):** switched to `inputs/roster_current_assets_15.xlsx` (column K = delivery date). New units 758002, 758006, 480009, 381071 and 753047 (fees from same-type medians where not in the fee file). 758006 (Tractor TA, 355) was delivered Nov 18, 2026 to replace 747106 (lost an engine), so 747106 and its 2027 replacement are out. 750046 is being sold. 755026 is on risk hold after an accident and carries depreciation only.

2027 depreciation is now $11.91M (was $11.64M); actual cost $36,248,570, location charges $35,951,164; replacements 76 units, $5,624,000.

## 2027 flat rate: fixed + variable (Nov 19)

`inputs/flat_rate_2027.xlsx` replaces the 2026 flat rates and the $0.425/mile charge for stores, loaner pools and HQ departments. Fuel is still charged separately.

- **Fixed:** $/month by asset type (base + ad valorem, registration, GPS, Lytx; trailers, reach trucks and scissor lifts include a flat maintenance amount).
- **Variable:** $/mile for trucks, tractors, pickups and cars (each unit's own 2026 miles, including the mileage-change input); $/hour for forklifts and truck-mounted forklifts, using each unit's average monthly hours through Oct 31, 2026 (`inputs/avg_monthly_hours_oct2026.xlsx`, validated DVIR hours); units without readings use the type average.
- Loaners that sit at a store are billed to their pool (761–768) or 735, per the file's Billed To. The five department units stay at 530, 590 and 620.
- Both rates are editable by type on the Assumptions tab. Saved browser scenarios reset to the new rates on first load.

Location charges $36,954,240 (fixed $12.20M + variable $8.98M, against $17.02M flat + $3.16M mileage before); actual cost $36,248,725.
