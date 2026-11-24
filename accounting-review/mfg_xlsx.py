import json, pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
d = json.load(open('company/mfg.json'))
R = pd.DataFrame(d['rows']); R.columns = ['acct', 'level', 'x'] + [f'm{i}' for i in range(1, 13)]
D = pd.DataFrame(d['out'], columns=['loc', 'level', 'cat', 'src', 'amt', 'id'])
lv = sorted(set(D.level)); R = R[R.level.isin(lv)].copy(); R['tot'] = R[[f'm{i}' for i in range(1, 13)]].sum(axis=1)
acc = d['acc']; c2a = {k: v[0] for k, v in acc.items()}
D['acct'] = D.cat.str.lower().map(c2a)
REQ = [a for a in sorted(R.acct.unique()) if a.startswith(('6602', '6610', '6625'))]
OTH = [a for a in sorted(R.acct.unique()) if a not in REQ]
H = Font(bold=True, color='FFFFFF'); HF = PatternFill('solid', fgColor='1F3864'); B = Font(bold=True); TF = PatternFill('solid', fgColor='D9E1F2')
money = '$#,##0;[Red]-$#,##0'; thin = Side(style='thin', color='BFBFBF')
wb = Workbook(); ws = wb.active; ws.title = '2027 by Level & Account'
ws['A1'] = "McCoy's Fleet: 2027 budget by Manufacturing & Distribution level and account (annual)"; ws['A1'].font = Font(bold=True, size=13)
ws['A2'] = 'Source: 2027 fleet forecast Adaptive export (Nov 19 build, with the agreed accounting revision). Ad valorem tax is excluded (property-tax account).'
def table(r0, accts, title):
    ws.cell(r0, 1, title).font = B; r0 += 1
    hdr = ['Account'] + lv + ['Total']
    for j, h in enumerate(hdr, 1):
        c = ws.cell(r0, j, h); c.font = H; c.fill = HF; c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    ws.row_dimensions[r0].height = 45
    first = r0 + 1
    for i, a in enumerate(accts):
        r = first + i; ws.cell(r, 1, a)
        for j, l in enumerate(lv, 2):
            v = R[(R.acct == a) & (R.level == l)].tot.sum(); c = ws.cell(r, j, round(float(v), 2) if v else 0); c.number_format = money
        c = ws.cell(r, len(lv) + 2, f'=SUM({L(2)}{r}:{L(len(lv)+1)}{r})'); c.number_format = money; c.font = B
    rt = first + len(accts); ws.cell(rt, 1, 'Total').font = B
    for j in range(2, len(lv) + 3):
        c = ws.cell(rt, j, f'=SUM({L(j)}{first}:{L(j)}{rt-1})'); c.number_format = money; c.font = B; c.fill = TF
    return rt + 2
r = table(4, REQ, '6602 Fleet Expense, 6610 Truck Rental, 6625 Forklift Rental')
table(r, OTH, 'Other fleet accounts at these levels (not in this request): fuel, forklift lease, Geotab, sale proceeds')
ws.column_dimensions['A'].width = 40
for j in range(2, len(lv) + 3): ws.column_dimensions[L(j)].width = 15
ws.freeze_panes = 'B6'
# build-up sheet
w2 = wb.create_sheet('Build-up by source')
w2['A1'] = 'What makes up each 2027 account (all Mfg/DC levels combined)'; w2['A1'].font = Font(bold=True, size=13)
srcs = ['Unit', 'Unit (replaced in 2027)', 'Planned new unit', 'Book-only depreciation (not on roster)', 'Budget adjustment (accounting revision)']
hdr = ['Account'] + srcs + ['Total']
for j, h in enumerate(hdr, 1):
    c = w2.cell(3, j, h); c.font = H; c.fill = HF; c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
w2.row_dimensions[3].height = 45
Dq = D[D.acct.isin(REQ)]
for i, a in enumerate(REQ):
    rr = 4 + i; w2.cell(rr, 1, a)
    for j, s in enumerate(srcs, 2):
        v = Dq[(Dq.acct == a) & (Dq.src == s)].amt.sum(); c = w2.cell(rr, j, round(float(v), 2)); c.number_format = money
    c = w2.cell(rr, len(srcs) + 2, f'=SUM(B{rr}:{L(len(srcs)+1)}{rr})'); c.number_format = money; c.font = B
rt = 4 + len(REQ); w2.cell(rt, 1, 'Total').font = B
for j in range(2, len(srcs) + 3):
    c = w2.cell(rt, j, f'=SUM({L(j)}4:{L(j)}{rt-1})'); c.number_format = money; c.font = B; c.fill = TF
notes = ['Unit: each unit on the Nov 18 roster at its 2027 location (maintenance from its 12-month repair history; depreciation from the books, stopping at month 61 / 85).',
 'Unit (replaced in 2027): old unit through its replacement delivery month, then the new unit (depreciation at cost / 60 or 84, median new-unit maintenance).',
 'Planned new unit: 359 15.5K forklift from January.',
 'Book-only depreciation: assets on the depreciation books but not on the fleet roster, mostly Chameleon rolling tarps at 353/356/359 (Sep 2026 actual depreciation, to estimated end dates).',
 'Budget adjustment: the agreed revision (maintenance from the 2026 outlook less replacement savings, spread by each level\'s share; rentals set from the Jan-Oct 2026 rental invoices by location; registration & inspections at the 2026 outlook).']
for i, n in enumerate(notes): w2.cell(rt + 2 + i, 1, n)
w2.column_dimensions['A'].width = 40
for j in range(2, len(srcs) + 3): w2.column_dimensions[L(j)].width = 18
# long format
w3 = wb.create_sheet('Monthly detail')
hdr = ['Level', 'Account'] + ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'] + ['Total']
for j, h in enumerate(hdr, 1): c = w3.cell(1, j, h); c.font = H; c.fill = HF
for i, x in enumerate(R.sort_values(['level', 'acct']).itertuples(), 2):
    w3.cell(i, 1, x.level); w3.cell(i, 2, x.acct)
    for k in range(12): c = w3.cell(i, 3 + k, round(getattr(x, f'm{k+1}'), 2)); c.number_format = money
    c = w3.cell(i, 15, f'=SUM(C{i}:N{i})'); c.number_format = money
w3.column_dimensions['A'].width = 32; w3.column_dimensions['B'].width = 38; w3.freeze_panes = 'C2'
wb.save('/mnt/user-data/outputs/Fleet_2027_MfgDC_by_Level_and_Account.xlsx')
print(R[R.acct.isin(REQ)].groupby('level').tot.sum().round(), R[R.acct.isin(REQ)].tot.sum())

# ---------- 2027 vs 2026 by level and account (accounting's DC variance check) ----------
from openpyxl import load_workbook
wb = load_workbook('/mnt/user-data/outputs/Fleet_2027_MfgDC_by_Level_and_Account.xlsx')
a = pd.read_excel('company/inputs/dc_fy2026_by_level.xlsx', header=None, skiprows=2)
a.columns = ['level', 'acctname', 'fy26', 'fy27', 'var', 'pct', 'chk']; a['level'] = a.level.replace('', None).ffill()
NAME = {'Forklift Repair & Maintenance': '6602_ForkliftRepair_Maintenance', 'Forklift Repairs & Tires': '6602_ForkliftRepairs_Tires', 'Forklifts - 7 yr': '6602_Forklifts_7yr',
        'Trailer Repair & Tires': '6602_TrailerRepair_Tires', 'Truck Repair & Maintenance': '6602_TruckRepair_Maintenance', 'Vehicle Heavy - > 13,000 - 5 yr': '6602_VehicleHeavy__13_000_5yr',
        'Vehicle Light - < 13,000 - 5 yr': '6602_VehicleLight__13_000_5yr', 'Vehicle Registrations & Inspections': '6602_VehicleRegistrations_Inspections',
        'Vehicle - 18 Wheeler - 5 yr': '6602_Vehicle_18Wheeler_5yr', 'Fleet Cameras': '6602_FleetCameras', 'Truck Rental': '6610_TruckRental', 'Forklift Rental': '6625_ForkliftRental',
        'Truck Leasing Expense': '6602_TruckLeasingExpense', 'Auto Repair & Maintenance': '6602_AutoRepair_Maintenance'}
a['acct'] = a.acctname.str.split(':').str[1].map(NAME)
f26 = a.groupby(['level', 'acct']).fy26.sum()
o27 = R[R.acct.isin(REQ)].groupby(['level', 'acct']).tot.sum()
keys = sorted(set(f26.index) | set(o27.index))
def note(l, ac, v26, v27):
    if l.startswith('355 Centralized'): return '355CD disbanded Jan 1, 2027: units moved to stores (flat rate), 355 Pharr Reload or 450, or were sold/returned. Clear this level for 2027.'
    if ac in ('6602_Forklifts_7yr', '6602_VehicleHeavy__13_000_5yr', '6602_VehicleLight__13_000_5yr', '6602_Vehicle_18Wheeler_5yr'):
        return 'Depreciation: 2026 budget has none at this level. Open question: budget it here from this forecast, or where accounting budgets depreciation today.'
    if ac == '6602_FleetCameras': return 'Intentional: DC and Millwork are straight P&L, so Lytx cameras hit the location starting 2027.'
    if ac == '6602_ForkliftRepair_Maintenance': return 'Not used: all forklift maintenance is in Forklift Repairs & Tires. Clear any amount from the earlier upload.'
    if ac == '6602_ForkliftRepairs_Tires': return 'All forklift maintenance (2026 amount, adjusted for units that moved between locations).'
    if ac == '6625_ForkliftRental' and l.startswith('359'): return 'Burnet forklift rental ended: the rental unit was purchased.'
    if l.startswith('479'): return 'No fleet units assigned to this level in the forecast; keep the amount accounting has.'
    if ac.startswith(('6602_Truck', '6602_Trailer', '6602_Auto', '6610', '6625')) and abs(v27 - v26) > 1000:
        return '2026 amount adjusted for units that moved between locations (e.g. 755031/755036 359 to 489, 377068/377069 489 to 359, 382004 356 to 359, 750027 355CD to 735).'
    if abs(v27 - v26) <= 1000: return 'Same as 2026.'
    return ''
w4 = wb.create_sheet('2027 vs 2026', 0)
w4['A1'] = "2027 vs 2026 by Manufacturing & Distribution level and account (6602, 6610, 6625)"; w4['A1'].font = Font(bold=True, size=13)
w4['A2'] = "FY2026 from accounting's DC Vehicle Expense Check. FY2027 from the fleet forecast (Nov 24, with 355CD disbanded Jan 1, 2027). Maintenance, registration and rentals at each level are the 2026 amount adjusted only for units that moved between locations; the 359 forklift rental is removed."
hdr = ['Level', 'Account', 'FY2026', 'FY2027', '2027 v 2026 $', '2027 v 2026 %', 'Note']
for j, h in enumerate(hdr, 1): c = w4.cell(4, j, h); c.font = H; c.fill = HF
LVN = {'356 Midland - DC': '356 Midland - DC', '450 Millwork - NB': '450 Millwork - NB'}
r = 5
for (l, ac) in keys:
    v26 = float(f26.get((l, ac), 0)); v27 = float(o27.get((l, ac), 0))
    if not v26 and not v27: continue
    w4.cell(r, 1, l); w4.cell(r, 2, ac)
    c = w4.cell(r, 3, round(v26, 2)); c.number_format = money
    c = w4.cell(r, 4, round(v27, 2)); c.number_format = money
    c = w4.cell(r, 5, f'=D{r}-C{r}'); c.number_format = money
    c = w4.cell(r, 6, f'=IF(C{r}=0,"",E{r}/C{r})'); c.number_format = '0.0%'
    w4.cell(r, 7, note(l, ac, v26, v27)); r += 1
w4.cell(r, 1, 'Total').font = B
for j, col in ((3, 'C'), (4, 'D'), (5, 'E')):
    c = w4.cell(r, j, f'=SUM({col}5:{col}{r-1})'); c.number_format = money; c.font = B; c.fill = TF
c = w4.cell(r, 6, f'=IF(C{r}=0,"",E{r}/C{r})'); c.number_format = '0.0%'; c.font = B
r += 1; w4.cell(r, 1, 'Total excluding depreciation').font = B
dep = "{'6602_Forklifts_7yr','6602_VehicleHeavy__13_000_5yr','6602_VehicleLight__13_000_5yr','6602_Vehicle_18Wheeler_5yr'}"
for j, col in ((3, 'C'), (4, 'D')):
    f = '-'.join([f'SUMIF($B$5:$B${r-2},"{x}",{col}$5:{col}${r-2})' for x in ['6602_Forklifts_7yr', '6602_VehicleHeavy__13_000_5yr', '6602_VehicleLight__13_000_5yr', '6602_Vehicle_18Wheeler_5yr']])
    c = w4.cell(r, j, f'={col}{r-1}-{f}'); c.number_format = money; c.font = B
c = w4.cell(r, 5, f'=D{r}-C{r}'); c.number_format = money; c.font = B
c = w4.cell(r, 6, f'=IF(C{r}=0,"",E{r}/C{r})'); c.number_format = '0.0%'; c.font = B
for col, wd in zip('ABCDEFG', (30, 40, 14, 14, 15, 13, 110)): w4.column_dimensions[col].width = wd
w4.freeze_panes = 'C5'
wb.save('/mnt/user-data/outputs/Fleet_2027_MfgDC_by_Level_and_Account.xlsx')
