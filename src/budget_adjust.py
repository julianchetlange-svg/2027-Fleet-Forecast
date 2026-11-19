"""Make inputs/budget_adjustments_2027.csv: per-location adjustments that bring the baseline to the revised 2027 lines agreed with accounting.
Maintenance and registration & inspections: the gap to the revised total is spread over locations in proportion to each location's baseline amount.
Rentals: set to the Jan-Oct 2026 rental invoices x 12/10 by location, less 011 South Austin, 950 Rio Truss and the 359 Burnet forklift rental.
Run after a build without the adjustment file: node dump_base.js base_by_loc.json && python3 budget_adjust.py"""
import json, pandas as pd, sys
TARGET = {'Maintenance - Truck': 7041518, 'Maintenance - LD': 117544, 'Maintenance - Forklift': 5485894, 'Maintenance - Trailer': 1780625,
          'Registration & inspections': 432055}
B = json.load(open('base_by_loc.json'))
rows = []
for c, t in TARGET.items():
    tot = sum(v.get(c, 0) for v in B.values()); gap = t - tot
    for l, v in B.items():
        if v.get(c): rows.append((l, c, gap * v[c] / tot))
    print(f'{c}: baseline {tot:,.0f} -> {t:,.0f} ({gap:+,.0f})')
r = pd.read_excel('inputs/2026_Rentals_Jan-Oct.xlsx'); r = r[~r.Status.str.startswith('Canc')]
r['dt'] = pd.to_datetime(r['Transaction Date'].str[:10]); r = r[r.dt >= '2026-01-01']
r['loc'] = r.Store.astype(str).str.split(' ').str[0].str.split('_').str[0].str.split('-').str[0]
r.loc[r.Store.astype(str).str.contains('Central', case=False), 'loc'] = '355CD'
r['loc'] = r['loc'].replace({'900': '735'})   # 900 Corporate has no fleet units; billed to 735 Fleet
r['t'] = r['Rental Type'].str.strip()
r = r[~r['loc'].isin(['011', '950']) & ~((r['loc'] == '359') & (r.t == 'Forklift Rental'))]
for (l, t), g in r.groupby(['loc', 't']):
    c = 'Truck rental' if t == 'Truck Rental' else 'Forklift rental'
    if l not in B: print('rental location not in forecast', l, round(g.Total.sum())); continue
    rows.append((l, c, g.Total.sum() * 1.2 - B[l].get(c, 0)))
for l, v in B.items():
    for c in ('Truck rental', 'Forklift rental'):
        if v.get(c) and not any(x[0] == l and x[1] == c for x in rows): rows.append((l, c, -v[c]))
D = pd.DataFrame(rows, columns=['loc', 'line', 'amount']).groupby(['loc', 'line'], as_index=False).amount.sum()
D.to_csv('inputs/budget_adjustments_2027.csv', index=False)
print(D.groupby('line').amount.sum().round(0).to_dict(), round(D.amount.sum()))
