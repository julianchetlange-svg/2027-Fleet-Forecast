from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
F='Arial'; b=Font(name=F,bold=True); n=Font(name=F); blue=Font(name=F,color='0000FF'); hdr=Font(name=F,bold=True,color='FFFFFF')
fill=PatternFill('solid',fgColor='0D5540'); yel=PatternFill('solid',fgColor='FFF2A8'); tot=PatternFill('solid',fgColor='E9EFE9')
money='$#,##0;($#,##0);-'; pct='0.0%;(0.0%);-'
grp={'Maintenance - Truck':(1540803,138506,4476478),'Maintenance - LD':(54411,0,225462),'Maintenance - Forklift':(779839,109318,4295947),
 'Maintenance - Trailer':(770106,50366,479308),'Registration & taxes':(195942,33250,905315),'Truck rental':(217487,0,51328),'Forklift rental':(6400,0,30780)}
rows=[('6602','Truck Repair & Maintenance','Maintenance - Truck',7119753,6155786,-87474,9239,'2026 outlook, less replacement savings, plus store 122 / new 359 forklift'),
 ('6602','Auto Repair & Maintenance','Maintenance - LD',125398,279872,-8294,440,'2026 outlook, less replacement savings, plus store 122'),
 ('6602','Forklift maintenance, repairs & tires','Maintenance - Forklift',5679344,5185104,-208586,15136,'2026 outlook, less replacement savings, plus store 122 / new 359 forklift'),
 ('6602','Trailer Repair & Tires','Maintenance - Trailer',1780625,1299779,0,0,'2026 outlook (no trailers on the 2027 replacement list)'),
 ('6602','Vehicle Registrations & Inspections','Registration & taxes',432055,1134507,0,0,'Hold at 2026 outlook; the $749K ad valorem tax in the plan moves to the property-tax account'),
 ('6610','Truck Rental','Truck rental',337770,9305,0,0,'2026 rental invoices Jan-Oct x 12/10, less 950 Rio Truss (not in the plan). Split by where the invoices were charged'),
 ('6625','Forklift Rental','Forklift rental',199805,5825,0,0,'2026 rental invoices Jan-Oct x 12/10, less 011 South Austin (turned in), 950 Rio Truss and the 359 Burnet rental (unit purchased)')]
wb=Workbook(); ws=wb.active; ws.title='2027 Revised Estimate'
H=['Account','Adaptive category','Fleet forecast line','FY2026 Outlook','FY2027 Plan (current)','Replacement savings','New units (122, 359)','Other adjustment','2027 Revised','Revised vs 2026 $','Revised vs 2026 %','Change vs current plan','Mfg/DC share (incl. 355CD/355)','Flat-rate share','Mfg/DC $','Flat rate $','Basis']
for c,h in enumerate(H,1):
    x=ws.cell(1,c,h); x.font=hdr; x.fill=fill; x.alignment=Alignment(wrap_text=True,horizontal='center',vertical='center')
ws.row_dimensions[1].height=42
for i,(acct,cat,line,o26,p27,sav,new,basis) in enumerate(rows,2):
    adj={'Truck rental':round(268814.96-337770),'Forklift rental':round(37180.11-199805)}.get(line,0)
    sv=sav
    vals=[acct,cat,line,o26,p27,sv,new,adj]
    for c,v in enumerate(vals,1):
        x=ws.cell(i,c,v); x.font=blue if c>=4 else n
        if c>=4: x.number_format=money
    ws.cell(i,9,f'=D{i}+F{i}+G{i}+H{i}'); ws.cell(i,10,f'=I{i}-D{i}'); ws.cell(i,11,f'=IF(D{i}=0,0,J{i}/D{i})'); ws.cell(i,12,f'=I{i}-E{i}')
    g=grp[line]; t=sum(g)
    m=round((g[0]+g[1])/t,4)
    x=ws.cell(i,13,m); x.font=blue; x.number_format=pct
    x=ws.cell(i,14,f'=1-M{i}'); x.font=n; x.number_format=pct
    ws.cell(i,15,f'=I{i}*M{i}'); ws.cell(i,16,f'=I{i}*N{i}')
    ws.cell(i,17,basis).font=n
    for c in (9,10,12,15,16): ws.cell(i,c).number_format=money; ws.cell(i,c).font=n
    ws.cell(i,11).number_format=pct; ws.cell(i,11).font=n
    for c in (9,15,16): ws.cell(i,c).fill=yel
r=len(rows)+2
ws.cell(r,1,'Total').font=b
for c in (4,5,6,7,8,9,10,12,15,16):
    L=ws.cell(1,c).column_letter; x=ws.cell(r,c,f'=SUM({L}2:{L}{r-1})'); x.font=b; x.number_format=money
ws.cell(r,11,f'=IF(D{r}=0,0,J{r}/D{r})').number_format=pct; ws.cell(r,11).font=b
for c in range(1,18): ws.cell(r,c).fill=tot
ws.cell(r+1,1,'Not in this table: ad valorem (property) tax on the fleet, $749,097, should be budgeted in the property-tax account rather than 6602.').font=Font(name=F,italic=True)
ws.cell(r+3,1,'Blue cells are inputs. Shares come from the 2027 fleet forecast (actual cost by location group). Mfg/DC = 353, 356, 359, 489, 450, plus 355CD and 355 Reload (per accounting). Flat rate = stores, loaner pools and HQ departments.').font=Font(name=F,italic=True,size=9)
ws.cell(2,6).comment=Comment('Old unit maintenance less the median for new units of the same type (bought 2024+), for the months after each 2027 delivery. 77 replacements.','Fleet forecast')
widths=[8,34,22,13,13,13,13,12,13,13,10,13,12,10,13,13,70]
for c,w in enumerate(widths,1): ws.column_dimensions[ws.cell(1,c).column_letter].width=w
ws.freeze_panes='D2'
# evidence sheet
e=wb.create_sheet('Why the plan is low')
E=[('Item','Amount','Detail'),
('Maintenance: plan vs 2026 outlook (4 lines)',12920541-14705120,'Plan $12.92M vs outlook $14.71M'),
('Invoices on units no longer on the roster, Oct 2025-Sep 2026',1405897,'Rolled off the forecast because it is built unit-by-unit from the current roster. $727K on units since sold, $598K on units disposed or returned, $80K on units excluded (e.g. Rio Truss). The fleet size is steady, so this spend does not go away; it moves to the units that replace them.'),
('Registration/inspection work orders excluded from maintenance',52263,'Moved out because registration comes from the fee file'),
('Rental work orders moved to rental lines',38401,''),
('Work-order system total, Oct 2025-Sep 2026',14379050,'Close to the 2026 GL outlook for the four maintenance lines ($14.71M); the remaining gap is GL-only items and the Oct-Dec outlook'),
('',None,''),
('Registration & taxes: plan',1134507,'From the fee file: registration $351,705 + ad valorem tax $749,097 + inspections/other'),
('Registration only (fee file)',351705,'Plus inspection work orders (~$52K) = ~$404K, in line with the 2026 outlook of $432,055'),
('Ad valorem tax (fee file)',749097,'0.8% of delivered cost. Not in the 2026 6602 outlook, so it is budgeted elsewhere'),
('',None,''),
('Rentals: 2026 outlook',537575,'Truck $337,770 + forklift $199,805'),
('Rentals in the asset-based work-order export, Oct 2025-Sep 2026',38401,'Most rental invoices have no asset attached, so they did not export and the plan missed them'),
('Rental invoices, Jan-Oct 2026 (full rental report)',354822,'Truck $224,012 + forklift $130,809; 322 invoices, cancelled ones excluded'),
('  Annualized (x 12/10)',425786,'About $112K below the 2026 outlook of $537,575; accounting to confirm what else is in the GL (accruals, other vendors)'),
('  Less 011 South Austin forklift rental, turned in',-19691*1.2,'$1,790/month'),
('  Less 950 Rio Truss',-6207*1.2,'Rio Truss is not in the 2027 fleet plan'),
('  Less 359 Burnet forklift rental, Jan-Jul 2026',-73928*1.2,'Ended in July; the rental unit was purchased'),
('  2027 rental estimate',305996,'Truck $268,815 + forklift $37,180; Mfg/DC $223,887, flat rate $82,108'),
('Rental timing',None,'Heaviest Jan-Jun (359 Burnet, 450 Millworks, 353 Dayton); Jul-Oct runs about $16K/month')]
for i,row in enumerate(E,1):
    for c,v in enumerate(row,1):
        x=e.cell(i,c,v); x.font=b if i==1 else n
        if c==2 and isinstance(v,(int,float)): x.number_format=money
for c in range(1,4): e.cell(1,c).fill=fill; e.cell(1,c).font=hdr
e.column_dimensions['A'].width=58; e.column_dimensions['B'].width=14; e.column_dimensions['C'].width=110
import os; wb.save(os.path.join(os.path.dirname(os.path.abspath(__file__)),'Fleet_2027_Budget_Variance_Review.xlsx'))
