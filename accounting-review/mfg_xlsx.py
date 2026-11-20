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
