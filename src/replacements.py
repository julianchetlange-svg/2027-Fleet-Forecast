"""2027 replacements: old unit cost stops after delivery, new unit straight-line depreciation, sale price forecast
from 2025-2026 healthy sales by type + brand + age (most specific level with >= 3 sales)."""
import pandas as pd, numpy as np, re

SALE_CAT = {  # sales-file Type -> forecast category
    'Truck Flatbed': '24TAFB', 'Truck - Lumber Dump': 'SLIDE', 'Truck-Tractor': 'TRACTOR', 'Forklift - Truck Mounted': 'TMFL',
    'Forklift - 8,000lb': 'FK8', 'Forklift-6,000lb': 'FK6', 'Forklift - 10,000lb': 'FK10', 'Pick Up-Light Duty': 'PICKUP',
    'Trailer-Flatbed': 'FBTRL', 'Trailer-Dry Van': 'VANTRL', 'Truck-Conveyor': 'CONV', 'Truck - Van Body': 'VANBODY',
    'Car - Light Duty': 'CAR', 'Van- Light Duty': 'CARGOVAN', 'Trailer - Cargo': 'MWTRL', 'Scissorlift': 'SCISSOR'}
ROSTER_CAT = {  # roster type -> (category, fallback note)
    "24' TA Flatbed": ('24TAFB', ''), "20' Slide Flatbed": ('SLIDE', ''), "12' Slide Flatbed": ('SLIDE', ''), "24' TA Slide Flatbed": ('SLIDE', ''),
    'Tractor TA': ('TRACTOR', ''), 'Tractor SA': ('TRACTOR', ''), 'Truck Mounted Forklift': ('TMFL', ''), '4-Way Truck Mounted Forklift': ('TMFL', ''),
    '8K Yard Forklift': ('FK8', ''), '6K Yard Forklift': ('FK6', ''), 'Pickup': ('PICKUP', ''),
    '10K Yard Forklift': ('FK8', 'Only 1 sale of a 10K forklift; priced from 8K forklift sales, likely low'),
    '13K Yard Forklift': ('FK8', 'No 13K forklift sales; priced from 8K forklift sales, likely low'),
    '15.5K Yard Forklift': ('FK8', 'No 15.5K forklift sales; priced from 8K forklift sales, likely low')}
BAD = 'Catastroph|Castroph|Castastroph|Wreck|Total|Fire|Frame|Structural|Mechanical Failure|Mechanical Hailure|Mechanical Issues'
BANDS = [0, 4, 6, 8, 10, 13, 99]

def band(age):
    for k in range(len(BANDS) - 1):
        if BANDS[k] <= age < BANDS[k + 1]: return k
    return len(BANDS) - 2

def load_sales(path):
    rows = []
    for s in ['Sales History 2026', 'Sales History 2025']:
        d = pd.read_excel(path, sheet_name=s, header=4, dtype={'Asset Number': str}); d.columns = [str(c).strip() for c in d.columns]; rows.append(d)
    d = pd.concat(rows)
    d['sold'] = pd.to_datetime(d['Date Sold'], errors='coerce'); d['price'] = pd.to_numeric(d['Sales Price'], errors='coerce')
    d = d[(d.sold >= '2025-01-01') & d.price.notna()]
    d = d[~d['Reason'].astype(str).str.contains(BAD, case=False)]
    d['my'] = d['Make'].astype(str).str.extract(r'^(\d{4})')[0].astype(float)
    d['oem'] = d['Make'].astype(str).str.replace(r'^\d{4}\s*', '', regex=True).str.strip().str.title()
    d['age'] = d.sold.dt.year + (d.sold.dt.month - 1) / 12 - d.my
    d['cat'] = d['Type'].map(SALE_CAT)
    d = d[d.cat.notna() & d.age.notna()]
    d['band'] = d.age.apply(band)
    return d

def price(sales, cat, oem, age):
    b = band(age); lab = f'{BANDS[b]}-{BANDS[b+1]} yrs' if BANDS[b + 1] < 99 else f'{BANDS[b]}+ yrs'
    s = sales[sales.cat == cat]
    for m, desc in [((s.oem == oem) & (s.band == b), f'{oem}, {lab}'), (s.oem == oem, f'{oem}, any age'), (s.band == b, f'all brands, {lab}'), (s.oem == s.oem, 'all brands, any age')]:
        x = s[m]
        if len(x) >= 3: return float(x.price.median()), f'{desc} ({len(x)} sales)'
    return (float(s.price.median()), f'all brands, any age ({len(s)} sales)') if len(s) else (0.0, 'no sales history')

def build(csv, sales_path, assets, start_year=2027):
    sales = load_sales(sales_path)
    r = pd.read_csv(csv, dtype={'Unit': str, 'Store': str})
    M = {a['id']: a for a in assets}
    out, missing = [], []
    for _, x in r.iterrows():
        a = M.get(x.Unit)
        if a is None: missing.append(x.Unit); continue
        m = int(x['Replacement Month']) - 1
        cost = float(str(x['Replace Cost']).replace('$', '').replace(',', '').strip())
        term = 60 if a['cls'] in ('Truck', 'Trailer') else 84
        cat, note = ROSTER_CAT.get(a['type'], (None, 'No sales category for this type'))
        oem = str(x.Manufacturer).strip().title(); my = int(x.Year)
        px, basis = [], []
        for k in range(12):   # sale month = delivery month + 1, for each possible delivery month
            sm = k + 1; age = start_year + sm / 12 - my
            p, b = price(sales, cat, oem, age) if cat else (0.0, 'no sales category')
            px.append(round(p, 2)); basis.append(b)
        out.append(dict(id=a['id'], m=m, cost=cost, term=term, fleet=x.Fleet, oem=oem, year=my, status=str(x.Status).replace('✓', '').strip(),
                        measure=str(x['Current Measure']), threshold=str(x.Threshold), lc=str(x['LC Exposure %']),
                        newReg=round(cost / 1500 + (cost * 0.0005 if a['cls'] == 'Truck' else 0), 4),
                        px=px, basis=basis, note=note, leased=a['own'] != 'Owned'))
    return out, missing, sales
