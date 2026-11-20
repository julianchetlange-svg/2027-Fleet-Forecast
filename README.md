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

Location charges $36,875,154 (fixed $12.20M + variable $8.90M, +$0.92M vs the old rate; 14 TMFLs and 3 forklifts use service-visit hour readings instead of DVIR (TMFLs 244041, 247012, 245042, 244009, 244026, 246008, 245009, 243012, 243026, 247006, 246009, 245010, 245001, 244045; forklifts 383109, 383043, 382048) use hours at their last service ÷ months in service, against $17.02M flat + $3.16M mileage before); actual cost $36,225,313 (747105 is on a $250/month holding rental, like 747104).

## Agreed accounting revision built in (Nov 19)

The dashboard and the Adaptive export now carry the revised 2027 red-box lines agreed with accounting (**$15,163,631**, −3.3% vs the FY2026 outlook):

| Line | 2027 |
|---|---|
| Truck repair & maintenance | $7,041,518 |
| Auto (LD) repair & maintenance | $117,544 |
| Forklift maintenance | $5,485,894 |
| Trailer repair & tires | $1,780,625 |
| Registrations & inspections | $432,055 |
| Truck rental | $268,815 |
| Forklift rental | $37,180 |

- One **budget adjustment** line per location (`inputs/budget_adjustments_2027.csv`, actual expense only) brings the unit-by-unit baseline to these totals. Maintenance and registration gaps are spread in proportion to each location's baseline; rentals are set by location from the Jan–Oct 2026 rental invoices × 12/10 (less 011, 950 and the 359 forklift rental). Regenerate with `node src/dump_base.js base_by_loc.json && python3 src/budget_adjust.py` after a build without the file.
- **Ad valorem tax** is now its own line ($770,860 in 2027 with replacements and new units) instead of sitting in registration. It shows in the dashboard but is left out of the Adaptive export, since accounting budgets it in the property-tax account.
- Split of the red-box lines: Mfg/DC $4,364,071, flat rate $10,799,560.

Totals: actual cost $38,081,457 ($28,480,390 without fuel); location charges $37,671,638 (the DCs and 355s carry the revision at actual cost).

## Adaptive export fixes after accounting's DC variance check (Nov 20)

- **Forklift maintenance** now goes to `6602_ForkliftRepairs_Tires`, where the 2026 history is. It had gone to Forklift Repair & Maintenance, which left accounting's existing Repairs & Tires amounts in place and counted forklift maintenance twice at the DC levels.
- **Level names** for 356 and 450 now match Adaptive: "356 Midland - DC" and "450 Millwork - NB".
- `accounting-review/Fleet_2027_MfgDC_by_Level_and_Account.xlsx`: 2027 annual by Mfg/DC level and account for 6602, 6610 and 6625.
- Open with accounting: whether 2027 depreciation at the DC and Millwork levels should come from this export or from accounting's own depreciation budget.

## New-unit maintenance backfill (Nov 20)

Units delivered since Oct 1, 2025 have less than a year in the Oct 2025–Sep 2026 repair window, so their history understated a full 2027 year (some were $0). 60 units are now budgeted at least the median maintenance of same-type units bought Jan 2024–Sep 2025, like the 2027 replacements (+$360,563 at unit level). The budget adjustments were regenerated, so the red-box total stays $15,163,631; the backfill changes where it lands (Mfg/DC $4,344,742). Location charges $37,652,309.

## Mfg/DC levels anchored to 2026 (Nov 20, option B)

Maintenance, registration and rentals at each Mfg/DC level (353, 355CD, 355, 356, 359, 450, 489) are now the FY2026 amount from accounting's DC variance check (`inputs/dc_fy2026_by_level.xlsx`), adjusted only for units that moved between locations; the 359 forklift rental is removed (unit purchased). Units that left the fleet are not taken out, the same conservative treatment as the company total. This fixes 356 and 450 (GL costs with no unit attached that the repair invoices miss) and 359 (which had carried more than its share). The flat-rate side carries the rest of each agreed total.

- Red-box lines: $15,163,631 (Mfg/DC $4,227,270; flat rate $10,936,361).
- Mfg/DC 6602/6610/6625 excluding depreciation: $4.27M vs $4.35M in 2026 (−1.9%). See the "2027 vs 2026" tab in `accounting-review/Fleet_2027_MfgDC_by_Level_and_Account.xlsx`.
- Location charges $37,534,838; actual cost $38,081,457.

## Revised 2027 flat rates (Nov 20)

`inputs/flat_rate_2027.xlsx` replaced with the revised rate file: fixed base + fees lowered about 2.5–3.5% on 23 rate types (e.g. 24' TA Flatbed $1,992.89 → $1,943.49, 6K forklift $466.26 → $453.86, TMFL $738.44 → $717.78, Pickup $646.91 → $631.55). Variable rates, flat maintenance and billing locations unchanged.

Stores, pools and HQ: fixed $11,888,457 + variable $8,900,867 = $20,789,324, against $20,180,006 under the old rate (+$609K, +3.0%). Location charges $37,220,320; actual cost unchanged at $38,081,457.
