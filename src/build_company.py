"""Company-wide 2027 fleet forecast data model (all locations except 950 Rio Truss)."""
import pandas as pd, numpy as np, json, os
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
fee_files = [A + '3f23cdde-Reg_Tax_Geo_LYTX.xlsx'] + ([INPUTS + 'fees.xlsx'] if os.path.exists(INPUTS + 'fees.xlsx') else [])
fee = pd.concat([pd.read_excel(f, sheet_name=0, dtype={'Asset Name': str}) for f in fee_files])
fee.columns = [c.strip() for c in fee.columns]
fee = fee.drop_duplicates('Asset Name', keep='last').set_index('Asset Name')
FC = ['Registraiton Fee', 'Monthly Taxes', 'Geotab', 'Lytx']
for c in FC: fee[c] = pd.to_numeric(fee[c], errors='coerce').fillna(0)
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
maint = mtw.groupby('Asset Name')[T].sum().to_dict()
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
for _, r in a.iterrows():
    aid, loc, c = r['Asset Name'], r['Store Number'], r.cls
    owned = r['Asset Ownership'] != 'Operating Lease'
    elig = [1 if (owned and 1 <= (2027 - r.isd.year) * 12 + (m - r.isd.month) + 1 <= r.term) else 0 for m in range(1, 13)]
    depm = float(dep.get(aid, 0.0)) if owned else 0.0
    rec = dict(id=aid, loc=loc, type=r['Asset Template Name'], cls=c, make=r['Asset Manufacturer Name'],
               own='Owned' if owned else 'Leased', status=r['Asset Status'], isd=r.isd.strftime('%Y-%m-%d'), term=int(r.term),
               dep=round(depm, 2), depElig=elig, depSrc=dep_src.get(aid, '') if owned and depm else '',
               lease=round(float(lease.get(aid, 0) or 0), 2) if not owned else 0.0, leaseNote=lease_note.get(aid, '') if not owned else '',
               flat=FLAT.get(r['Asset Template Name']), maint=round(float(maint.get(aid, 0.0)), 2),
               miles=0.0, gal=0.0, lp=0.0, milesNote='', galNote='', fuelNote='')
    if aid in fee.index:
        rec['reg'] = round(float(fee.loc[aid, FC[0]]) + float(fee.loc[aid, FC[1]]), 4); rec['tel'] = float(fee.loc[aid, FC[2]]) + float(fee.loc[aid, FC[3]])
    else:
        rec['reg'] = 0.0; rec['tel'] = 0.0
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
    for k in ('milesNote', 'galNote', 'fuelNote', 'leaseNote'):
        if r[k]: f.append(r[k])
    if r.cls == 'Forklift' and r['loc'] in lp_missing and r['loc'] != '900': f.append('No propane data for this location')
    if r.id == '747104': f.append('Owned wrap depreciation ($132.37/mo) excluded')
    return f
df['flags'] = df.apply(flags, axis=1)

FUEL27 = [5.198, 5.016, 4.834, 4.590, 4.455, 4.320, 4.185, 4.095, 4.005, 3.915, 3.825, 3.735]
LP27 = [1.624, 1.637, 1.626, 1.610, 1.604, 1.585, 1.591, 1.578, 1.567, 1.540, 1.524, 1.525]
cols = ['id', 'loc', 'type', 'cls', 'make', 'own', 'status', 'isd', 'term', 'dep', 'depElig', 'depSrc', 'lease', 'flat',
        'maint', 'miles', 'gal', 'gpm', 'lp', 'reg', 'tel', 'flags', 'galW', 'lpW', 'dc']
out = dict(generated='2026-10-30',
           locations=[dict(id=r['Store Number'], name=r['name'], district=r.district, kind=r.kind) for _, r in locs.iterrows()],
           districts=DORDER, assets=json.loads(df[cols].to_json(orient='records')),
           flat=FLAT, fuelRate=FUEL27, lpRate=LP27, dcFuelRate=DC_FUEL, dcLpRate=DC_LP, mileRate=0.425,
           maintSeason=dict(company=season(ms)))
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
