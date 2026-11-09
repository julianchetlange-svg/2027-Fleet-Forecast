"""Company-wide 2027 fleet forecast data model (all locations except 950 Rio Truss)."""
import pandas as pd, numpy as np, json, os, re
A = '/mnt/user-data/attachments/beabc8da-8639-59c1-85f4-c5bad2fd86c1/'
S = '/tmp/claude-0/-home-claude/beabc8da-8639-59c1-85f4-c5bad2fd86c1/scratchpad/'
OUT = S + 'company/'
WIN = [f'2025-{m:02d}' for m in (10, 11, 12)] + [f'2026-{m:02d}' for m in range(1, 10)]
INPUTS = OUT + 'inputs/'   # user-supplied fill-ins (depreciation, leases, fees) land here
os.makedirs(INPUTS, exist_ok=True)

# ---------- roster ----------
a = pd.read_excel(A + '76d1ece1-Current_Assets_7.xlsx', header=1, dtype=str)
a = a[a['Store Number'] != '950'].copy()          # Rio Truss excluded
a['Store District'] = a['Store District'].fillna('')
TRUCK = {'Tractor TA', 'Tractor SA', 'Pickup', "24' TA Flatbed", "20' Slide Flatbed", "12' Slide Flatbed", "24' Conveyor",
         "16' Conveyor", "16' Van Body", "24' Van Body", "24' TA Slide Flatbed", 'Car', 'Cargo Van'}
EQUIP = {'Golf Cart', 'Scissor Lift', 'Light Tower', 'Sweeper', 'Order Picker', 'Electric Pallet Jack'}
def cls(t):
    if t in TRUCK: return 'Truck'
    if 'Truck Mounted' in t: return 'TMFL'
    if 'Trailer' in t: return 'Trailer'
    if t in EQUIP: return 'Equipment'
    return 'Forklift'
a['cls'] = a['Asset Template Name'].map(cls)
a['term'] = a.cls.map({'Truck': 60, 'Trailer': 60, 'Forklift': 84, 'TMFL': 84, 'Equipment': 84})
a['isd'] = pd.to_datetime(a['Asset Date of Origination'])

# ---------- locations ----------
ACTUAL_FIXED = {'355CD', '355'}
def district(r):
    d, s = r['Store District'], r['Store Number']
    if s in ('355CD', '355', '033'): return 'Region 05'
    if s == '762': return 'Region 02'
    if d.startswith('Region'): return d
    return d  # Distribution, Millwork, Headquarters
a['district'] = a.apply(district, axis=1)
locs = a.groupby('Store Number').agg(name=('Store Name', 'first'), district=('district', 'first')).reset_index()
def kind(r):
    s = r['Store Number']
    if s in ACTUAL_FIXED: return 'Central'
    if s.startswith('76'): return 'Loaner pool'
    if r.district.startswith('Region'): return 'Store'
    return r.district  # Distribution / Millwork / Headquarters (actual cost)
locs['kind'] = locs.apply(kind, axis=1)
DORDER = [f'Region 0{i}' for i in range(1, 9)] + ['Distribution', 'Millwork', 'Headquarters']
locs['dord'] = locs.district.map(lambda d: DORDER.index(d))
locs = locs.sort_values(['dord', 'Store Number'])
LOCKIND = dict(zip(locs['Store Number'], locs.kind))

# ---------- depreciation ----------
dp = pd.read_excel(A + 'f93ae8cd-BA_-_McCoys_Asset_Depreciation_Forecast_Fleet_Detail.xlsx', sheet_name=1, dtype=str)
dp.columns = [c.strip() for c in dp.columns]
dp['id'] = dp['Asset Identifier'].str.strip(); dp['amt'] = dp['Depreciation Amount'].astype(float)
dep = dp.groupby('id').amt.sum().to_dict()
dep_src = {k: 'March 2026 file' for k in dep}
for k, v in {'758005': 166407.25, '758011': 166407.25, '626005': 71123.99, '626003': 71123.99}.items():
    dep[k] = v / 60; dep_src[k] = 'Cost ÷ 60 (provided)'
# user fill-ins (Fleet_Forecast_2027_Missing_Inputs.xlsx, returned) if present
fill = INPUTS + 'missing_inputs.xlsx'
lease_fill = {}
if os.path.exists(fill):
    md = pd.read_excel(fill, sheet_name='Missing depreciation', dtype={'Asset': str})
    for _, r in md.iterrows():
        term = int(r['Depreciation Term (months)'])
        if pd.notna(r.get('OR Monthly Depreciation (fill in)')): dep[r.Asset] = float(r['OR Monthly Depreciation (fill in)']); dep_src[r.Asset] = 'Provided monthly amount'
        elif pd.notna(r.get('Original Cost (fill in)')): dep[r.Asset] = float(r['Original Cost (fill in)']) / term; dep_src[r.Asset] = f'Cost ÷ {term} (provided)'
    ml = pd.read_excel(fill, sheet_name='Missing lease rates', dtype={'Asset': str})
    for _, r in ml.iterrows():
        if pd.notna(r.get('Monthly Lease Rate (fill in)')): lease_fill[r.Asset] = float(r['Monthly Lease Rate (fill in)'])

# ---------- leases ----------
ls = pd.read_excel(A + 'a848caef-Current_Assets_8.xlsx', dtype=str)
lease = dict(zip(ls['Asset Name'], pd.to_numeric(ls['Asset Lease Fixed Rate'], errors='coerce')))
lease_note = {}
for k, v in lease.items():
    if pd.isna(v):
        if k in lease_fill: lease[k] = lease_fill[k]
        else: lease[k] = 447.11; lease_note[k] = 'Rate missing; 2020 Toyota batch rate $447.11 used'
lease['747104'] = 250.0; lease_note['747104'] = 'Holding rental $250/mo'
lease['747106'] = 150.0; lease_note['747106'] = 'Holding rental $150/mo'

# ---------- flat rates (stores) ----------
FLAT = {'10K Yard Forklift': 1305.79, "12' Slide Flatbed": 1192.80, "16' Conveyor": 1841.10, "16' Van Body": 1274.35,
        "20' Slide Flatbed": 1216.97, "24' Conveyor": 2264.57, "24' TA Flatbed": 1981.47, "24' TA Slide Flatbed": 2489.40,
        "24' Van Body": 1691.42, '3K Yard Forklift': 845.43, '4K Yard Forklift': 935.65, '4-Way Truck Mounted Forklift': 1343.17,
        '6K Yard Forklift': 1007.80, '8K Boxcar Forklift': 1343.23, '8K Yard Forklift': 1262.59, 'Electric Pallet Jack': 201.04,
        'Flatbed Trailer': 914.99, 'Millwork Trailer': 415.81, 'Open Air Trailer': 183.99, 'Order Picker': 1025.11,
        'Pickup': 995.46, 'Pup Trailer': 153.75, 'Reach Truck': 1400.63, 'Tractor SA': 1715.35, 'Tractor TA': 1741.26,
        'Truck Mounted Forklift': 1152.43}

# ---------- fees (registration, taxes, Geotab, Lytx; monthly) ----------
fee_files = [A + '3f23cdde-Reg_Tax_Geo_LYTX.xlsx', A + '04beebbc-Current_Assets_Reg_Tax_Geo_Lytx.xlsx'] + ([INPUTS + 'fees.xlsx'] if os.path.exists(INPUTS + 'fees.xlsx') else [])
def ldfee(f):
    x = pd.read_excel(f, sheet_name=0, dtype={'Asset Name': str}); x.columns = [c.strip() for c in x.columns]
    return x.rename(columns={'Registration Monthly': 'Registraiton Fee', 'Taxas Monthly': 'Monthly Taxes'})
FC = ['Registraiton Fee', 'Monthly Taxes', 'Geotab', 'Lytx']
fee = None
for f in fee_files:   # later files win; a blank in a later file keeps the earlier file's value
    if not os.path.exists(f): continue
    x = ldfee(f).drop_duplicates('Asset Name', keep='last').set_index('Asset Name')
    for c in FC + ['Asset Original Value']: x[c] = pd.to_numeric(x[c], errors='coerce') if c in x else float('nan')
    x = x[FC + ['Asset Original Value']]
    fee = x if fee is None else x.combine_first(fee)
FEE_NA = set(fee.index[fee[FC[0]].isna() | fee[FC[1]].isna()])   # still blank: estimated from cost below
for c in FC: fee[c] = fee[c].fillna(0)
HASFEE = set(fee.index)

# ---------- maintenance ----------
def ldm(f):
    d = pd.read_excel(A + f, dtype={'Asset Name': str, 'Subcontractor Invoice Reference Number': str})
    d.columns = [c.strip() for c in d.columns]
    d['dt'] = pd.to_datetime(d['Subcontractor Invoice Transaction Date'], format='ISO8601'); return d
T = 'Subcontractor Invoice Total'
mt = pd.concat([ldm('b4b61a20-Asset_Spend_Data_2025_11.xlsx'), ldm('101fbf67-Asset_Spend_Data_2026_.xlsx')])
mt = mt.drop_duplicates(['Subcontractor Invoice Reference Number', 'Asset Name', 'Work Order ID', T])
mtw = mt[(mt.dt >= '2025-10-01') & (mt.dt < '2026-10-01')]
isreg = mtw['Work Order Category Full Path'].fillna('').str.contains('Regist', case=False)
# registration work orders leave maintenance only where the fee file supplies registration (avoids double count)
mtw = mtw[~(isreg & mtw['Asset Name'].isin(HASFEE))]
# rentals (work order categories "Truck - Rental" / "Forklift - Rental") are broken out of maintenance (Nov 9)
_cat = mtw['Work Order Category Full Path'].fillna('')
_isrent = _cat.str.contains(r'- Rental', regex=True)
rentT = mtw[_isrent & _cat.str.contains('Truck - Rental')].groupby('Asset Name')[T].sum().to_dict()
# 011 South Austin turned in its recurring forklift rental; it does not continue into 2027 (user, Nov 9)
_sa = mtw['Work Order Store Identifier'].fillna('').str.startswith('011')
print('011 forklift rental dropped', round(mtw[_isrent & _cat.str.contains('Forklift - Rental') & _sa][T].sum()))
rentF = mtw[_isrent & _cat.str.contains('Forklift - Rental') & ~_sa].groupby('Asset Name')[T].sum().to_dict()
print('rentals in window: truck', round(sum(rentT.values())), 'forklift', round(sum(rentF.values())))
maint = mtw[~_isrent].groupby('Asset Name')[T].sum().to_dict()
# tire share of each asset's maintenance (for the Adaptive forklift tires account)
_tire = mtw[~_isrent & _cat.str.contains('Tires', case=False)].groupby('Asset Name')[T].sum().to_dict()
TIRE = {k: round(_tire.get(k, 0) / v, 4) for k, v in maint.items() if v > 0 and _tire.get(k, 0) > 0}
ms = mt[(mt.dt >= '2025-01-01') & (mt.dt < '2026-10-01') & ~mt['Work Order Category Full Path'].fillna('').str.contains('Regist', case=False)]
def season(x):
    m = x.groupby([x.dt.dt.year, x.dt.dt.month])[T].sum().unstack(0)
    i = (m / m.mean()).mean(axis=1).reindex(range(1, 13))
    return [float(v) for v in (i / i.mean())]

# ---------- miles & fuel ----------
mi = pd.read_pickle(S + 'miles_all.pkl')
POOL = {'761', '762', '763', '764', '765', '766', '767', '768'}
mi['key'] = mi.apply(lambda r: r.sfx if r.asset in POOL else r.asset, axis=1)
mi = mi[mi.month.isin(WIN)]
mpv = mi.pivot_table(index='key', columns='month', values='mi', aggfunc='sum')
fu = pd.read_pickle(S + 'fuel_all.pkl')
fu['Units'] = pd.to_numeric(fu.Units, errors='coerce'); fu['asset'] = fu['Custom Vehicle/Asset ID'].str.strip()
p = fu['Product Description'].fillna('')
fu = fu[p.str.contains('Diesel|Unleaded|Mid/Plus|Premium/Super|E-15|E-85|Biodiesel|Exhaust', case=False) & (fu['Unit of Measure'] == 'GA')]
gal = fu.groupby('asset').Units.sum().to_dict()

# ---------- propane ----------
lp = pd.read_excel(A + 'd30b360a-LP_Gallons_by_Location_2025_2026.xlsx', sheet_name='Monthly Detail', header=4, dtype={'Location': str})
lp.columns = [str(c).strip() for c in lp.columns]
lp = lp.dropna(subset=['Location']).set_index('Location')[WIN].sum(axis=1)
LP_GROUPS = {'033': ['033', '355', '355CD'], '359': ['359', '489']}   # one LP account shared by several locations

# ---------- assemble ----------
rows = []
_fv = a.set_index('Asset Name').join(fee[['Asset Original Value']], how='left')
_fv = _fv[_fv['Asset Original Value'] > 0]
def peer_cost(ty, yr):
    # same asset type, closest purchase years: widen the window until at least 3 peers (or the nearest year that has any)
    g = _fv[_fv['Asset Template Name'] == ty]; yrs = g.isd.dt.year
    if not len(g): return None, ''
    if (yrs == yr).sum(): return float(g[yrs == yr]['Asset Original Value'].median()), f'median cost of {int((yrs == yr).sum())} same-type assets bought in {yr}'
    for k in range(1, 40):
        m = (yrs - yr).abs() <= k
        if m.sum() >= 3 or (m.sum() and k >= 4):
            if m.sum() == 1: return float(g[m]['Asset Original Value'].iloc[0]), f'cost of the closest-age same-type asset (bought {int(yrs[m].iloc[0])})'
            return float(g[m]['Asset Original Value'].median()), f'median cost of {int(m.sum())} same-type assets bought {yr - k}-{yr + k}'
    near = yrs.iloc[(yrs - yr).abs().argsort().iloc[0]]
    return float(g[yrs == near]['Asset Original Value'].median()), f'median cost of same-type assets bought in {near} (closest year)'
for _, r in a.iterrows():
    aid, loc, c = r['Asset Name'], r['Store Number'], r.cls
    owned = r['Asset Ownership'] != 'Operating Lease'
    elig = [1 if (owned and 1 <= (2027 - r.isd.year) * 12 + (m - r.isd.month) + 1 <= r.term) else 0 for m in range(1, 13)]
    depm = float(dep.get(aid, 0.0)) if owned else 0.0
    rec = dict(id=aid, loc=loc, type=r['Asset Template Name'], cls=c, make=r['Asset Manufacturer Name'],
               own='Owned' if owned else 'Leased', status=r['Asset Status'], isd=r.isd.strftime('%Y-%m-%d'), term=int(r.term),
               dep=round(depm, 2), depElig=elig, depSrc=dep_src.get(aid, '') if owned and depm else '',
               lease=round(float(lease.get(aid, 0) or 0), 2) if not owned else 0.0, leaseNote=lease_note.get(aid, '') if not owned else '',
               flat=FLAT.get(r['Asset Template Name']), maint=round(float(maint.get(aid, 0.0)), 2), tire=TIRE.get(aid, 0.0), rentT=round(float(rentT.get(aid, 0.0)), 2), rentF=round(float(rentF.get(aid, 0.0)), 2),
               miles=0.0, gal=0.0, lp=0.0, milesNote='', galNote='', fuelNote='')
    if aid in fee.index:
        rec['reg'] = round(float(fee.loc[aid, FC[0]]) + float(fee.loc[aid, FC[1]]), 4); rec['tel'] = float(fee.loc[aid, FC[2]]) + float(fee.loc[aid, FC[3]]); rec['geo'] = float(fee.loc[aid, FC[2]]); rec['lytx'] = float(fee.loc[aid, FC[3]])
        if aid in FEE_NA:   # fee file left this blank (no original value): same formulas, cost from depreciation x term
            pv, how = peer_cost(r['Asset Template Name'], r.isd.year)
            cost = pv if pv else depm * int(r.term); how = how if pv else 'depreciation x term'
            rec['reg'] = round(cost * (0.0005 if c in ('Truck', 'Trailer') else 0) + cost / 1500, 4)
            rec['feeNote'] = f'Registration/taxes estimated (no original value in fee file): cost ${cost:,.0f}, {how}'
    else:
        rec['reg'] = 0.0; rec['tel'] = 0.0; rec['geo'] = 0.0; rec['lytx'] = 0.0
    if c == 'Truck':
        src = '692009' if aid == '626003' else aid
        if src in mpv.index:
            s = mpv.loc[src].dropna()
            rec['miles'] = float(s.mean() * 12) if len(s) else 0.0
            if 0 < len(s) < 12: rec['milesNote'] = f'{len(s)} of 12 months reported; annualized from average'
        if aid == '626003': rec['milesNote'] = 'Uses predecessor 692009 history'
        rec['gal'] = float(gal.get(src, 0.0))
        if aid == '758005':
            s = mpv.loc['758005', ['2026-07', '2026-08', '2026-09']]; mpg = mpv.loc['758005'].sum() / gal['758005']
            rec['miles'] = float(s.mean() * 12); rec['gal'] = rec['miles'] / mpg; rec['milesNote'] = 'New Apr 2026; Jul–Sep run rate'
    rows.append(rec)
df = pd.DataFrame(rows)

# Region 5 new-truck proxies (same as the Region 5 analysis)
for aid, like in [('758011', '758005'), ('626005', '697033')]:
    s = df[df.id == like].iloc[0]; df.loc[df.id == aid, ['miles', 'gal']] = [s.miles, s.gal]
    df.loc[df.id == aid, 'milesNote'] = f'New; uses 355CD {s.type} average ({like})'

# gallons per mile; type medians fill gaps
tr = df.cls == 'Truck'
df['gpm'] = np.where((df.miles > 0) & (df.gal > 0), df.gal / df.miles.replace(0, np.nan), np.nan)
med_gpm = df[tr].groupby('type').gpm.median(); all_gpm = df[tr].gpm.median()
med_mi = df[tr & (df.miles > 0)].groupby('type').miles.median(); all_mi = df[tr & (df.miles > 0)].miles.median()
for i, r in df[tr].iterrows():
    tm = med_gpm.get(r.type, np.nan); tm = all_gpm if pd.isna(tm) else tm
    if r.miles == 0 and r.gal == 0:
        mm = med_mi.get(r.type, np.nan); mm = all_mi if pd.isna(mm) else mm
        df.loc[i, ['miles', 'gal', 'gpm']] = [mm, mm * tm, tm]; df.loc[i, 'milesNote'] = f'No miles or fuel history; {r.type} company median used'
    elif r.miles == 0:
        df.loc[i, ['miles', 'gpm']] = [r.gal / tm, tm]; df.loc[i, 'milesNote'] = 'No miles reported; miles estimated from fuel at type median MPG'
    elif r.gal == 0:
        df.loc[i, ['gal', 'gpm']] = [r.miles * tm, tm]; df.loc[i, 'galNote'] = 'No WEX fuel; gallons estimated from miles at type median MPG'
    elif r.gpm < tm * 0.6 or r.gpm > tm * 1.7:
        df.loc[i, 'gpm'] = tm; df.loc[i, 'galNote'] = f'MPG outside type range; changes use {r.type} median'
df['gpm'] = df.gpm.fillna(0)

# propane split evenly across propane forklifts in each LP account
fk = df.cls == 'Forklift'
lp_missing = []
done = set()
for acct, members in LP_GROUPS.items():
    m = fk & df['loc'].isin(members); n = m.sum()
    if n: df.loc[m, 'lp'] = float(lp.get(acct, 0)) / n
    done |= set(members)
for loc in df['loc'].unique():
    if loc in done: continue
    m = fk & (df['loc'] == loc); n = m.sum()
    if not n: continue
    if loc in lp.index: df.loc[m, 'lp'] = float(lp[loc]) / n
    else: lp_missing.append(loc)
df.loc[fk & (df['loc'] == '900'), 'fuelNote'] = 'Airport tractor (moves company planes); Jet A not tracked; very low use, no fuel forecast'

# ---------- Distribution onsite tanks (Distribution_Fuel_Forecast_2027.xlsx) ----------
# Tank diesel is added to each truck's actual WEX gallons (replacing estimates), split by miles;
# tank propane replaces the AmeriGas split for those forklifts. Monthly shape follows the forecast.
df['galW'] = None; df['lpW'] = None
df['dc'] = 0; DC_FUEL = DC_LP = None
DCF = A + '5b77a9c6-Distribution_Fuel_Forecast_2027.xlsx'
if os.path.exists(DCF):
    inp = pd.read_excel(DCF, sheet_name='Inputs', header=None)
    def block(start):
        out = {}
        for r in range(start, start + 3):
            name = str(inp.iloc[r, 0]); vals = [float(v) for v in inp.iloc[r, 1:13]]
            out[name.split()[0]] = vals
        return out
    tank_d, tank_p = block(5), block(12)   # rows: 353, 356, 359 (incl. 489)
    DC_FUEL = [round(float(v), 4) for v in inp.iloc[8, 1:13]]; DC_LP = [round(float(v), 4) for v in inp.iloc[15, 1:13]]
    print('DC prices', DC_FUEL, DC_LP)
    groups = {'353': ['353'], '356': ['356'], '359': ['359', '489']}
    df.loc[df['loc'].isin(['353', '356', '359', '489']), 'dc'] = 1
    for acct, members in groups.items():
        dm, pm = tank_d[acct], tank_p[acct]
        tr_m = (df.cls == 'Truck') & df['loc'].isin(members)
        wex = df.loc[tr_m].apply(lambda r: 0.0 if 'No WEX fuel' in (r.galNote or '') else r.gal, axis=1)
        miles = df.loc[tr_m, 'miles']; share = miles / miles.sum()
        for i in df.index[tr_m]:
            tank = sum(dm) * share[i]; w = wex[i]
            total = w + tank
            monthly = [w / 12 + share[i] * dm[k] for k in range(12)]
            df.at[i, 'gal'] = total
            df.at[i, 'galW'] = [m * 12 / total for m in monthly] if total else None
            df.at[i, 'gpm'] = total / df.at[i, 'miles'] if df.at[i, 'miles'] else df.at[i, 'gpm']
            df.at[i, 'galNote'] = 'Onsite tank diesel (Distribution 2027 forecast) split by miles' + (', plus WEX purchases' if w else '')
        fk_m = (df.cls == 'Forklift') & df['loc'].isin(members); n = fk_m.sum()
        if n:
            tot = sum(pm)
            for i in df.index[fk_m]:
                df.at[i, 'lp'] = tot / n; df.at[i, 'lpW'] = [v * 12 / tot for v in pm]
    print('tank diesel gal', round(sum(sum(v) for v in tank_d.values())), 'tank LP gal', round(sum(sum(v) for v in tank_p.values())))

# flags
def flags(r):
    f = []
    if r.status == 'Loaner': f.append('Loaner unit')
    if r.own == 'Owned' and r.dep == 0 and any(r.depElig): f.append('In term but no depreciation amount (awaiting input)')
    if r.id not in HASFEE: f.append('Registration, taxes, Geotab and Lytx pending company-wide fee file')
    if isinstance(r.get('feeNote'), str) and r.get('feeNote'): f.append(r.get('feeNote'))
    for k in ('milesNote', 'galNote', 'fuelNote', 'leaseNote'):
        if r[k]: f.append(r[k])
    if r.cls == 'Forklift' and r['loc'] in lp_missing and r['loc'] != '900': f.append('No propane data for this location')
    if r.id == '747104': f.append('Owned wrap depreciation ($132.37/mo) excluded')
    return f
df['flags'] = df.apply(flags, axis=1)

FUEL27 = [5.198, 5.016, 4.834, 4.590, 4.455, 4.320, 4.185, 4.095, 4.005, 3.915, 3.825, 3.735]
LP27 = [1.624, 1.637, 1.626, 1.610, 1.604, 1.585, 1.591, 1.578, 1.567, 1.540, 1.524, 1.525]
cols = ['id', 'loc', 'type', 'cls', 'make', 'own', 'status', 'isd', 'term', 'dep', 'depElig', 'depSrc', 'lease', 'flat',
        'maint', 'tire', 'rentT', 'rentF', 'miles', 'gal', 'gpm', 'lp', 'reg', 'tel', 'geo', 'lytx', 'flags', 'galW', 'lpW', 'dc']
# ---------- planned new store: 122 Canyon Lake (Region 07), opens mid-June 2027, costs from July ----------
NEW_LOC = dict(id='122', name='Canyon Lake', district='Region 07', kind='Store')
NEW_FLEET = [("24' TA Flatbed", 2), ('Truck Mounted Forklift', 2), ('6K Yard Forklift', 3), ('8K Yard Forklift', 1), ('Pickup', 1)]
START = 6   # July
act = [0] * START + [1] * (12 - START)
r7 = df['loc'].isin(locs.loc[(locs.district == 'Region 07') & (locs.kind == 'Store'), 'Store Number'])
stores = df['loc'].isin(locs.loc[locs.kind == 'Store', 'Store Number'])
isd = pd.to_datetime(df['isd'])
def med(mask, col):
    s = df.loc[mask, col]; s = s[s > 0]; return float(s.median()) if len(s) else 0.0
planned = []
for ty, n in NEW_FLEET:
    same = df['type'] == ty; cls = df.loc[same, 'cls'].iloc[0]
    use = (r7 & same) if (r7 & same & (df.miles > 0 if cls == 'Truck' else True)).sum() >= 2 else (stores & same)
    newer = same & (isd >= '2024-01-01')
    miles = med(use, 'miles') if cls == 'Truck' else 0.0
    gpm = med(use, 'gpm') if cls == 'Truck' else 0.0
    lp = med(use & (df.cls == 'Forklift'), 'lp') if cls == 'Forklift' else 0.0
    dep = med(newer & (df.own == 'Owned'), 'dep'); maint = med(newer, 'maint') or med(same, 'maint')
    reg = med(newer, 'reg'); geo = med(newer, 'geo') if cls == 'Truck' else 0.0; lytx = med(newer, 'lytx') if cls == 'Truck' else 0.0
    term = int(df.loc[same, 'term'].iloc[0])
    for k in range(n):
        planned.append(dict(id=f'122-NEW-{len(planned) + 1:02d}', loc='122', type=ty, cls=cls, make='Planned', own='Owned', status='Planned',
            isd='2027-07-01', term=term, dep=round(dep, 2), depElig=act, depSrc='Estimate: median of same-type units bought 2024 or later',
            lease=0.0, flat=FLAT.get(ty), maint=round(maint, 2), miles=round(miles, 1), gal=round(miles * gpm, 1), gpm=gpm, lp=round(lp, 1),
            reg=round(reg, 4), tel=geo + lytx, geo=geo, lytx=lytx, act=act, galW=None, lpW=None, dc=0,
            flags=[f'Planned: 122 Canyon Lake opens mid-June 2027; costs start July. Miles, fuel and propane are the '
                   + ('Region 07' if use is not None and (use & r7).any() else 'company store') + ' median for this type; depreciation, maintenance and fees are the median of same-type units bought 2024 or later.']))
print('planned assets', len(planned), [(p['type'], p['miles'], p['gal'], p['lp'], p['dep'], p['maint'], p['reg']) for p in planned[::1]])

# ---------- loaner pools 765 and 768: no roster assets, so fuel is estimated from the other pools ----------
pool_gal = df[df['loc'].isin(POOL)].groupby('loc').gal.sum()
POOL_EST = float(pool_gal.mean())
EXTRA_LOCS = []
for pid, reg_name, dist in [('765', 'Region 5', 'Region 05'), ('768', 'Region 8', 'Region 08')]:
    if pid in set(df['loc']): continue
    EXTRA_LOCS.append(dict(id=pid, name=reg_name, district=dist, kind='Loaner pool'))
    planned.append(dict(id=f'{pid}-FUEL', loc=pid, type='Fuel-only account', cls='Truck', make='Estimate', own='Owned', status='Estimate',
        isd='2027-01-01', term=60, dep=0.0, depElig=[0] * 12, depSrc='', lease=0.0, flat=None, maint=0.0, miles=0.0, gal=round(POOL_EST, 1), gpm=0.0, lp=0.0,
        reg=0.0, tel=0.0, geo=0.0, lytx=0.0, galW=None, lpW=None, dc=0,
        flags=[f'Fuel only: no {pid} assets on the roster; fuel is the {pid} card account actual (Nov 2025-Oct 2026).']))
print('pool gal by pool', pool_gal.round().to_dict(), 'estimate', round(POOL_EST))

out = dict(generated='2026-10-30',
           locations=[dict(id=r['Store Number'], name=r['name'], district=r.district, kind=r.kind) for _, r in locs.iterrows()] + [NEW_LOC] + EXTRA_LOCS,
           districts=DORDER, assets=json.loads(df[cols].to_json(orient='records')) + planned,
           flat=FLAT, fuelRate=FUEL27, lpRate=LP27, dcFuelRate=DC_FUEL, dcLpRate=DC_LP, mileRate=0.425,
           maintSeason=dict(company=season(ms)))
# ---------- fuel = actual gallons by location, Nov 2025 - Oct 2026 (fuel card accounts + onsite tanks; user, Nov 4) ----------
# Each location's trucks are scaled so its monthly gallons equal its row (fuel stays with the card account that bought it).
# Tank rows go to 353 / 356 / 359 (359 includes 489). 068 is closed and dropped. Accounts with no fueled truck get a fuel-only line.
# All gallons are priced at the blended 2027 rate.
# stores and everyone else: 2027 blended rate; Distribution (353/356/359/489): 2027 diesel forecast prices (Julian, Nov 9)
out['fuelRate'] = FUEL27[:]; out['dcFuelRate'] = DC_FUEL[:]; out['ratesVer'] = 'blended-stores-dc-diesel-2026-11-09'
fl = pd.read_excel(INPUTS + 'fuel_by_location.xlsx', sheet_name='Gallons', header=2, dtype={'Location': str})
fl = fl[fl['Source'].isin(['Fuel card', 'On-site tank'])].copy()
fl['code'] = fl['Location'].str.extract(r'\((\d+)\)')[0].fillna(fl['Location']).str.strip()
MCOLS = [c for c in fl.columns if str(c).endswith('2027') and len(str(c)) == 8]
GROUP = {'489': '359'}
tgt = {}
for _, r in fl.iterrows():
    c = GROUP.get(r.code, r.code)
    if c == '068': continue
    tgt[c] = [x + float(v) for x, v in zip(tgt.get(c, [0.0] * 12), r[MCOLS])]
grp = lambda l: GROUP.get(l, l)
known = {l['id'] for l in out['locations']}
OTHER = []
for c, m in tgt.items():
    has = [a for a in out['assets'] if grp(a['loc']) == c and a['cls'] == 'Truck' and a['gal'] > 0 and not a.get('act')]
    if has: continue
    if c not in known:
        OTHER.append(c); out['locations'].append(dict(id=c, name='Fuel account ' + c, district='Other fuel accounts', kind='Fuel account'))
    out['assets'] = [a for a in out['assets'] if a['id'] != f'{c}-FUEL']
    out['assets'].append(dict(id=f'{c}-FUEL', loc=c, type='Fuel-only account', cls='Truck', make='Fuel account', own='Owned', status='Fuel only',
        isd='2027-01-01', term=60, dep=0.0, depElig=[0] * 12, depSrc='', lease=0.0, flat=None, maint=0.0, miles=0.0, gal=1.0, gpm=0.0, lp=0.0,
        reg=0.0, tel=0.0, geo=0.0, lytx=0.0, galW=None, lpW=None, dc=0,
        flags=[f'Fuel only: card account {c} has actual fuel (Nov 2025-Oct 2026) but no trucks on the roster.']))
if OTHER: out['districts'] = out['districts'] + ['Other fuel accounts']
before = {}; after = {}
for a in out['assets']:
    if a['cls'] != 'Truck' or a.get('act'): continue
    c = grp(a['loc']); before[c] = before.get(c, 0) + a['gal']
for a in out['assets']:
    if a['cls'] != 'Truck' or a.get('act') or not a['gal']: continue
    c = grp(a['loc']); m = tgt.get(c, [0.0] * 12); share = a['gal'] / before[c]
    old = a['gal']; mg = [x * share for x in m]; tot = sum(mg)
    a['gal'] = round(tot, 1); a['galW'] = [x * 12 / tot for x in mg] if tot else None
    if a.get('miles'): a['gpm'] = tot / a['miles']
    after[c] = after.get(c, 0) + tot
    if not tot: a['flags'] = a['flags'] + [f'No fuel: no {c} account or tank row in the actual fuel file (Nov 2025-Oct 2026)']
FUEL_SRC = dict(total=round(sum(sum(v) for v in tgt.values())), months=[round(sum(v[i] for v in tgt.values())) for i in range(12)])
out['fuelSrc'] = FUEL_SRC
print('fuel by location: target', FUEL_SRC['total'], 'assigned', round(sum(after.values())), 'other accounts', OTHER,
      'zeroed', sorted(c for c in before if c not in tgt and before[c] > 0))
# ---------- depreciation GL lines by asset type (Depreciation_Type.xlsx, Nov 9) ----------
_dt = pd.read_excel(INPUTS + 'depreciation_type.xlsx', header=4)
out['depType'] = {str(k).strip(): str(v).strip() for k, v in zip(_dt.iloc[:, 0], _dt.iloc[:, 1]) if str(k).strip() not in ('', 'nan')}
print('depreciation lines', sorted(set(out['depType'].values())))

# ---------- maintenance GL lines by asset type (Maintenace_Type.xlsx, Nov 9) ----------
_mt = pd.read_excel(INPUTS + 'maintenance_type.xlsx', header=4)
out['maintType'] = {str(k).strip(): str(v).strip() for k, v in zip(_mt.iloc[:, 0], _mt.iloc[:, 1]) if str(k).strip() not in ('', 'nan')}
_unm = sorted({a['type'] for a in out['assets'] if a['maint'] and a['type'] not in out['maintType']})
print('maintenance lines', sorted(set(out['maintType'].values())), 'unmapped with maintenance', _unm)

# ---------- sale proceeds GL lines by asset type (Proceeds_Type.xlsx, Nov 9) ----------
_pt = pd.read_excel(INPUTS + 'proceeds_type.xlsx', header=4)
_norm = lambda v: re.sub(r'\s+', ' ', re.sub(r'(?i)^sale proceeds\s*-\s*', 'Sale proceeds - ', str(v).strip()))
out['procType'] = {str(k).strip(): _norm(v) for k, v in zip(_pt.iloc[:, 0], _pt.iloc[:, 1]) if str(k).strip() not in ('', 'nan')}
print('proceeds lines', sorted(set(out['procType'].values())))

# ---------- lease GL lines by asset type (Lease_Type.xlsx, Nov 9) ----------
_lt = pd.read_excel(INPUTS + 'lease_type.xlsx', header=4)
out['leaseType'] = {str(k).strip(): str(v).strip() for k, v in zip(_lt.iloc[:, 0], _lt.iloc[:, 1]) if str(k).strip() not in ('', 'nan')}
print('lease lines', sorted(set(out['leaseType'].values())), 'unmapped leased', sorted({a['type'] for a in out['assets'] if a['lease'] and a['type'] not in out['leaseType']}))

# ---------- Adaptive export mapping (Adaptive_Import_Template.xlsx, Nov 9) ----------
_fa = pd.read_excel(INPUTS + 'adaptive_template.xlsx', sheet_name='Fleet Accounts', dtype=str)
_key = lambda s: re.sub(r'\s+', ' ', str(s).strip().lower())
_acc = {}
for _n, _c in zip(_fa.iloc[:, 0], _fa.iloc[:, 1]):
    if str(_n).strip() in ('', 'nan'): continue
    _acc.setdefault(_key(_n), []).append(str(_c).strip())
_lv = pd.read_excel(INPUTS + 'adaptive_template.xlsx', sheet_name='Levels', dtype=str)
_lvl = {}
for _s, _c in zip(_lv.iloc[:, 0], _lv.iloc[:, 1]):
    _code = str(_s).strip().split()[0]
    _lvl.setdefault(_code, []).append((str(_s).strip(), str(_c).strip()))
LEVEL = {}
for l in out['locations']:
    i = l['id']
    if i == '355CD': LEVEL[i] = next(c for s, c in _lvl['355'] if 'Central' in s)
    elif i == '355': LEVEL[i] = next(c for s, c in _lvl['355'] if 'Reload' in s)
    elif i in _lvl: LEVEL[i] = _lvl[i][0][1]
FLEET_LEVEL = _lvl['735'][0][1]
# department fuel accounts are not on the Levels tab; codes from Julian (Nov 9). They use the charges format.
LEVEL.update({'510': '510 Information Services Dept', '530': '530 People Development', '620': '620 Marketing Dept', '740': '740 Risk Management'})
for l in out['locations']:
    if l['id'] not in LEVEL: LEVEL[l['id']] = FLEET_LEVEL
    if l['id'] in ('510', '530', '620', '740'): l['name'] = LEVEL[l['id']].split(' ', 1)[1]
CHG_KINDS = {'Store', 'Loaner pool', 'Headquarters', 'Fuel account'}
out['adaptive'] = dict(accounts=_acc, levels=LEVEL, fleetLevel=FLEET_LEVEL,
                       fmt={l['id']: ('charge' if l['kind'] in CHG_KINDS else 'op') for l in out['locations']},
                       header=[str(c) for c in pd.read_excel(INPUTS + 'adaptive_template.xlsx', sheet_name='Adaptive Import', header=None, dtype=str).iloc[0]])
print('adaptive: accounts', len(_acc), 'levels', len(LEVEL), 'charge-format locs', sum(v == 'charge' for v in out['adaptive']['fmt'].values()),
      'to 735 (no level)', [l['id'] for l in out['locations'] if l['id'] not in _lvl and l['id'] not in ('355CD', '355')])

# ---------- 2027 replacements and sale proceeds (Nov 8) ----------
import sys as _sys; _sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import replacements as _R
out['repl'], _miss, _s = _R.build(INPUTS + 'replacements_2027.csv', INPUTS + 'remarketing.xlsx', out['assets'])
out['replMissing'] = _miss
print('replacements', len(out['repl']), 'not on roster', _miss, 'sales used', len(_s))
json.dump(out, open(OUT + 'data.json', 'w'), separators=(',', ':'))

# ---------- checks ----------
print('assets', len(df), 'locations', len(locs), df.cls.value_counts().to_dict())
print(locs.groupby(['district', 'kind']).size().to_string())
print('LP accounts missing for forklift locations:', lp_missing)
print('dep in-term missing:', int(((df.own == 'Owned') & (df.dep == 0) & df.depElig.apply(any)).sum()))
print('assets with fees:', int(df.id.isin(HASFEE).sum()))
print('truck miles', round(df.miles.sum()), 'gal', round(df.gal.sum()), 'lp', round(df.lp.sum()), 'maint', round(df.maint.sum()))
print('flag counts', pd.Series([x for f in df['flags'] for x in f]).str.split(';').str[0].value_counts().head(12).to_string())
print('json KB', os.path.getsize(OUT + 'data.json') // 1024)
