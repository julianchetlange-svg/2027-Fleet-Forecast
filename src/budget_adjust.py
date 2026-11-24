"""Make inputs/budget_adjustments_2027.csv: per-location adjustments that bring the baseline to the revised 2027 lines agreed with accounting.

Mfg/DC levels (353, 355CD, 355, 356, 359, 450, 489), step 2 option B (Julian, Nov 20): each level and line is its FY2026 amount in
accounting's DC variance check (inputs/dc_fy2026_by_level.xlsx), adjusted only for units that moved between locations (their Oct 2025-Sep 2026
repair invoices leave the old level and go with them to the new one). Units that left the fleet are not taken out, the same conservative
treatment as the company total. The 359 forklift rental is removed (unit purchased).
Flat-rate side: the rest of each agreed company total, spread over stores, pools and HQ in proportion to their baseline (rentals: in
proportion to their Jan-Oct 2026 rental invoices, less 011 South Austin and 950 Rio Truss).
Run after a build without the adjustment file: node dump_base.js base_by_loc.json && python3 budget_adjust.py"""
import json, pandas as pd
TARGET = {'Maintenance - Truck': 7041518, 'Maintenance - LD': 117544, 'Maintenance - Forklift': 5485894, 'Maintenance - Trailer': 1780625,
          'Registration & inspections': 432055, 'Truck rental': 268815, 'Forklift rental': 37180}
DCS = ['353', '355CD', '355', '356', '359', '450', '489']
B = json.load(open('base_by_loc.json'))
D = json.load(open('data.json')); A = {a['id']: a for a in D['assets']}; MT = D['maintType']
# FY2026 by level and line from accounting's report
r = pd.read_excel('inputs/dc_fy2026_by_level.xlsx', header=None, skiprows=2)
r.columns = ['level', 'acct', 'fy26', 'fy27', 'var', 'pct', 'chk']; r['level'] = r.level.replace('', None).ffill()
LV = {'353 Dayton Distribution': '353', '355 Centralized Delivery': '355CD', '355 Pharr - Reload': '355', '356 Midland - DC': '356',
      '359 Burnet Reload': '359', '450 Millwork - NB': '450', '489 Burnet EWP': '489'}
AC = {'Forklift Repair & Maintenance': 'Maintenance - Forklift', 'Forklift Repairs & Tires': 'Maintenance - Forklift', 'Trailer Repair & Tires': 'Maintenance - Trailer',
      'Truck Repair & Maintenance': 'Maintenance - Truck', 'Auto Repair & Maintenance': 'Maintenance - LD', 'Vehicle Registrations & Inspections': 'Registration & inspections',
      'Truck Rental': 'Truck rental', 'Forklift Rental': 'Forklift rental'}
r['loc'] = r.level.map(LV); r['line'] = r.acct.str.split(':').str[1].map(AC)
FY = r[r['loc'].notna() & r.line.notna()].groupby(['loc', 'line']).fy26.sum()
# unit changes since the repair window
s = pd.read_pickle('../spend_all.pkl'); s['dt'] = pd.to_datetime(s['Subcontractor Invoice Transaction Date'], format='mixed')
s = s[(s.dt >= '2025-10-01') & (s.dt < '2026-10-01')].copy()
st = s['Work Order Store Identifier'].astype(str)
s['L'] = st.str[:3].where(~st.str.contains('Central', case=False), '355CD'); s['nm'] = s['Asset Name'].astype(str)
s['line'] = s['Asset Template Name'].map(lambda x: MT.get(x))
was = s[s.L.isin(DCS) & s.line.notna()]
cur = lambda i: A[i]['loc'] if i in A and A[i]['status'] not in ('Book only', 'Budget adjustment') else None
mv = was[[(cur(i) is not None and cur(i) != L) for i, L in zip(was.nm, was.L)]]
out_ = mv.groupby(['L', 'line'])['Subcontractor Invoice Total'].sum()
in_ = {}
for (L, nm), g in mv.groupby(['L', 'nm']):
    for ln, v in g.groupby('line')['Subcontractor Invoice Total'].sum().items(): in_[(cur(nm), ln)] = in_.get((cur(nm), ln), 0) + v
# 355CD disband (Nov 24): units replaced by 355CD units leave their level (their invoices come out), and 355CD itself carries nothing
_dx = pd.read_excel('inputs/355CD_disband_2027.xlsx', dtype=str); _dx.columns = [c.strip() for c in _dx.columns]
_old = {str(x).strip() for x in _dx['Repalces'] if str(x).strip() not in ('', 'nan')}
_ro = was[was.nm.isin(_old)].groupby(['L', 'line'])['Subcontractor Invoice Total'].sum()
for k, v in _ro.items(): out_[k] = out_.get(k, 0) + v
DISBANDED = {'355CD'}
rows, dc_t = [], {}
for L in DCS:
    for c in TARGET:
        t = float(FY.get((L, c), 0))
        if c.startswith('Maintenance'): t += -float(out_.get((L, c), 0)) + in_.get((L, c), 0)
        if L == '359' and c == 'Forklift rental': t = 0.0     # Burnet rental unit purchased
        if L in DISBANDED: t = 0.0
        t = max(t, 0.0); dc_t[(L, c)] = t
        rows.append((L, c, t - B.get(L, {}).get(c, 0)))
# flat-rate side
rent = pd.read_excel('inputs/2026_Rentals_Jan-Oct.xlsx'); rent = rent[~rent.Status.str.startswith('Canc')]
rent['dt'] = pd.to_datetime(rent['Transaction Date'].str[:10]); rent = rent[rent.dt >= '2026-01-01']
rent['loc'] = rent.Store.astype(str).str.split(' ').str[0]; rent.loc[rent.Store.astype(str).str.contains('Central', case=False), 'loc'] = '355CD'
rent['loc'] = rent['loc'].replace({'900': '735'}); rent = rent[~rent['loc'].isin(['011', '950'])]
rent['line'] = rent['Rental Type'].str.strip().map({'Truck Rental': 'Truck rental', 'Forklift Rental': 'Forklift rental'})
RB = rent[~rent['loc'].isin(DCS)].groupby(['loc', 'line']).Total.sum() * 1.2
for c, T in TARGET.items():
    rest = T - sum(dc_t[(L, c)] for L in DCS)
    if c.endswith('rental'):
        share = {l: v for (l, cc), v in RB.items() if cc == c and l in B}
    else:
        share = {l: v.get(c, 0) for l, v in B.items() if l not in DCS and v.get(c, 0)}
    tot = sum(share.values())
    for l, v in B.items():
        if l in DCS: continue
        want = rest * share.get(l, 0) / tot if tot else 0
        if want or v.get(c, 0): rows.append((l, c, want - v.get(c, 0)))
    print(f'{c}: Mfg/DC {sum(dc_t[(L, c)] for L in DCS):,.0f}  flat-rate {rest:,.0f}  total {T:,.0f}')
X = pd.DataFrame(rows, columns=['loc', 'line', 'amount']).groupby(['loc', 'line'], as_index=False).amount.sum()
X = X[X.amount.abs() > 0.005]; X.to_csv('inputs/budget_adjustments_2027.csv', index=False)
print('Mfg/DC by level', {L: round(sum(dc_t[(L, c)] for c in TARGET)) for L in DCS}, round(sum(dc_t.values())))
