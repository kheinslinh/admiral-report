#!/usr/bin/env python3
"""Build the raw-data Excel report from data.json (output of prepare_data.py).

Usage: python3 -I build_excel.py data.json out.xlsx
Then recalculate: python3 <xlsx skill>/scripts/recalc.py out.xlsx 90

Sheets: Summary (client KPIs, prev vs current), Data Sources (auto vs manual map), Inquiries,
Campaign Cost, Lost Reasons, Organic Facebook (Manual), Leads_Data. Every KPI is a formula on
Leads_Data / Campaign Cost, so the team can edit statuses and the workbook recalculates.
Colour code: green = Rework auto, yellow = Rework + team update / inputs, orange = manual from Meta.
"""
import sys, json, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation

data = json.load(open(sys.argv[1])); OUT = sys.argv[2]
rows = data['leads']; ADSETS = data['adsets']; PK = data['packages']; prev = data['previous']
PER = data['period_label']; PREV = data.get('prev_label', 'Previous month')
F = 'Arial'; NAVY = '1F3864'; AUTO = 'E2EFDA'; AUTOT = '375623'; MAN = 'FCE4D6'; MANT = 'C55A11'; INP = 'FFF2CC'; GREY = 'F2F2F2'
font = lambda **k: Font(name=F, **k)
thin = Side(style='thin', color='D9D9D9'); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
fill = lambda c: PatternFill('solid', fgColor=c)
wb = openpyxl.Workbook()

def title(ws, t, sub):
    ws['A1'] = t; ws['A1'].font = font(bold=True, size=16, color=NAVY)
    ws['A2'] = sub; ws['A2'].font = font(size=10, color='595959'); ws.sheet_view.showGridLines = False
def section(ws, r, text, n=8):
    ws.cell(r, 1, text).font = font(bold=True, size=11, color='FFFFFF')
    for c in range(1, n + 1): ws.cell(r, c).fill = fill(NAVY)
def header(ws, r, vals):
    for i, v in enumerate(vals, 1):
        c = ws.cell(r, i, v); c.font = font(bold=True, size=10); c.fill = fill(GREY); c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical='center')
def cell(ws, r, c, v, fmt=None, bold=False, src=None, wrap=False):
    x = ws.cell(r, c, v); x.font = font(size=10, bold=bold); x.border = BOX
    if fmt: x.number_format = fmt
    if wrap: x.alignment = Alignment(wrap_text=True, vertical='top')
    if src: x.fill = fill({'auto': AUTO, 'manual': MAN, 'input': INP}[src])
    return x
def tag(ws, r, c, src):
    t = {'auto': ('Rework – auto', AUTO, AUTOT), 'manual': ('Manual – Meta', MAN, MANT), 'mixed': ('Rework + team update', INP, '7F6000')}[src]
    x = ws.cell(r, c, t[0]); x.fill = fill(t[1]); x.font = font(size=9, bold=True, color=t[2]); x.border = BOX
    x.alignment = Alignment(horizontal='center')

# ---------------- Leads_Data
D = 'Leads_Data'
ld = wb.active; ld.title = D
title(ld, f"Lead data – {data['client']['short']}, {PER}", 'One row per Messenger inquiry. Source: Rework Prospectors (page inbox) + team status tracking + Sales pipeline.')
header(ld, 4, ['Date', 'Time', 'Week', 'Source', 'Ad Set', 'Full Name', 'Package', 'Status', 'Lost Reason', 'Lost Reason Category',
               'Contact Type', 'Contact Info', 'Notes', 'Status Reviewed', 'Inbox Conversation ID'])
rows = sorted(rows, key=lambda x: (x['date'] or '', x['time'] or ''))
for i, x in enumerate(rows):
    r = 5 + i
    d = datetime.datetime.strptime(x['date'], '%Y-%m-%d') if x['date'] else None
    t = datetime.datetime.strptime(x['time'], '%H:%M').time() if x['time'] else None
    vals = [d, t, f'=IF(A{r}="","",A{r}-WEEKDAY(A{r},3))', x['source'], x['adset'], x['name'], x['package'], x['status'], x['reason'],
            x['reason_category'], x['contact_type'], x['contact'], x['notes'], x['reviewed'], x['conversation_id']]
    fm = ['mmm d, yyyy', 'h:mm AM/PM', 'mmm d'] + [None] * 11 + ['0']
    for c, (v, f) in enumerate(zip(vals, fm), 1): cell(ld, r, c, v, f)
    if x['reviewed'] == 'No': ld.cell(r, 14).fill = fill(INP)
    if x.get('adset_estimated'): ld.cell(r, 5).fill = fill(INP)
LAST = 4 + len(rows)
for c, w in zip('ABCDEFGHIJKLMNO', [13, 10, 9, 16, 44, 26, 13, 15, 30, 24, 14, 22, 40, 10, 14]): ld.column_dimensions[c].width = w
ld.freeze_panes = 'A5'; ld.auto_filter.ref = f'A4:O{LAST}'
dv = DataValidation(type='list', formula1='"Intake,Qualified,Pending Payment,Paid Booking,Declined,Cancelled"'); ld.add_data_validation(dv); dv.add(f'H5:H{LAST + 200}')
dv2 = DataValidation(type='list', formula1='"' + ','.join([p['code'] for p in PK] + ['Not specified']) + '"'); ld.add_data_validation(dv2); dv2.add(f'G5:G{LAST + 200}')
ld['E4'].comment = Comment('Yellow = ad set assigned by the date the conversation started (Meta did not tag the ad). Estimate.', 'Rework')
ld['N4'].comment = Comment('No = imported from the Rework inbox and not yet reviewed by the sales team (status defaults to Intake).', 'Rework')
R = lambda col: f"{D}!${col}$5:${col}${LAST + 200}"
SRC, ADS, PKG, ST, LRC, REV = R('D'), R('E'), R('G'), R('H'), R('J'), R('N')

# ---------------- Campaign Cost
ad = wb.create_sheet('Campaign Cost')
title(ad, f"Campaign cost – {data['client']['short']} ad sets, {PER}", 'Source: Rework Meta Ads > Ad sets (lifetime per ad set). CPM, CPC, frequency and CTR are calculated.')
section(ad, 4, 'Ad set performance', 17)
header(ad, 5, ['Campaign', 'Ad Set', 'Objective', 'Start', 'End', 'Spend', 'Impressions', 'Reach', 'Clicks', 'CTR', 'CPM', 'CPC', 'Frequency',
               'Inquiries', 'Cost per Inquiry', 'Paid Bookings', 'Meta-reported Results'])
for i, a in enumerate(ADSETS):
    r = 6 + i
    vals = [a['campaign'], a['adset'], a['objective'], datetime.datetime.fromisoformat(a['start_at']), datetime.datetime.fromisoformat(a['end_at']),
            a['spend'], a['impressions'], a['reach'], a['clicks']]
    for c, v in enumerate(vals, 1):
        cell(ad, r, c, v, 'mmm d' if c in (4, 5) else '#,##0.00' if c == 6 else '#,##0' if c in (7, 8, 9) else None, src='auto')
    cell(ad, r, 10, f'=IF(G{r}=0,"",I{r}/G{r})', '0.00%'); cell(ad, r, 11, f'=IF(G{r}=0,"",F{r}/G{r}*1000)', '#,##0.00')
    cell(ad, r, 12, f'=IF(I{r}=0,"",F{r}/I{r})', '#,##0.00'); cell(ad, r, 13, f'=IF(H{r}=0,"",G{r}/H{r})', '0.00')
    cell(ad, r, 14, f'=COUNTIFS({ADS},B{r})', '#,##0'); cell(ad, r, 15, f'=IF(N{r}=0,"",F{r}/N{r})', '#,##0.00')
    cell(ad, r, 16, f'=COUNTIFS({ADS},B{r},{ST},"Paid Booking")', '#,##0'); cell(ad, r, 17, a.get('meta_results'), '#,##0', src='manual')
TR = 6 + len(ADSETS)
cell(ad, TR, 1, 'Total', bold=True)
for c in range(2, 18): cell(ad, TR, c, None)
for c, L in [(6, 'F'), (7, 'G'), (8, 'H'), (9, 'I'), (14, 'N'), (16, 'P'), (17, 'Q')]:
    cell(ad, TR, c, f'=SUM({L}6:{L}{TR - 1})', '#,##0.00' if c == 6 else '#,##0', bold=True)
cell(ad, TR, 10, f'=IF(G{TR}=0,"",I{TR}/G{TR})', '0.00%', bold=True); cell(ad, TR, 11, f'=IF(G{TR}=0,"",F{TR}/G{TR}*1000)', '#,##0.00', bold=True)
cell(ad, TR, 12, f'=IF(I{TR}=0,"",F{TR}/I{TR})', '#,##0.00', bold=True); cell(ad, TR, 15, f'=IF(N{TR}=0,"",F{TR}/N{TR})', '#,##0.00', bold=True)
ad['H5'].comment = Comment('Total = sum of ad-set reach; people reached by several ad sets are counted more than once.', 'Rework')
ad['Q5'].comment = Comment('MANUAL: results reported by Meta Ads Manager for the same dates. Not stored in Rework analytics.', 'Rework')
r = TR + 2; tag(ad, r, 1, 'auto'); ad.cell(r, 2, 'Spend, impressions, reach, clicks: Rework Meta Ads. Inquiries & paid bookings: counted from Leads_Data.').font = font(size=9)
tag(ad, r + 1, 1, 'manual'); ad.cell(r + 1, 2, 'Meta-reported Results (column Q): copy from Meta Ads Manager.').font = font(size=9)
for c, w in zip('ABCDEFGHIJKLMNOPQ', [40, 52, 12, 9, 9, 11, 11, 10, 10, 8, 9, 8, 9, 9, 11, 9, 12]): ad.column_dimensions[c].width = w
ad.freeze_panes = 'C6'

# ---------------- Organic (Manual)
og = wb.create_sheet('Organic Facebook (Manual)')
title(og, f'Organic Facebook – {PER}', 'MANUAL: Rework does not store organic page insights. Copy from Meta Business Suite > Insights.')
section(og, 4, 'A. Page performance', 5); header(og, 5, ['KPI', PREV, 'This month', 'Change'])
OK = data['organic']['kpis']
for i, k in enumerate(OK):
    r = 6 + i; cell(og, r, 1, k['label']); cell(og, r, 2, k['prev'], '#,##0'); cell(og, r, 3, k['current'], '#,##0', src='manual')
    cell(og, r, 4, f'=IF(OR(C{r}="",B{r}=0),"",(C{r}-B{r})/B{r})', '0.0%')
pr = 6 + len(OK) + 1
section(og, pr, 'B. Content published', 12)
header(og, pr + 1, ['Date', 'Post title', 'Content pillar', 'Series', 'Format', 'Views', 'Reach', 'Reactions', 'Shares', 'Comments', 'Total engagement', 'Engagement rate'])
posts = data['organic'].get('posts', [])
for j in range(max(15, len(posts))):
    r = pr + 2 + j; p = posts[j] if j < len(posts) else {}
    for c, key in enumerate(['date', 'title', 'pillar', 'series', 'format', 'views', 'reach', 'reactions', 'shares', 'comments'], 1):
        cell(og, r, c, p.get(key), src='manual')
    cell(og, r, 11, f'=IF(B{r}="","",SUM(H{r}:J{r}))', '#,##0'); cell(og, r, 12, f'=IF(OR(B{r}="",G{r}=0),"",K{r}/G{r})', '0.00%')
for c, w in zip('ABCDEFGHIJKL', [30, 34, 22, 16, 12, 9, 9, 10, 9, 10, 12, 12]): og.column_dimensions[c].width = w

# ---------------- Lost Reasons
lr = wb.create_sheet('Lost Reasons')
title(lr, f'Why inquiries did not convert – {PER}', 'Source: Status and Lost Reason on Leads_Data (Rework Sales lost reasons once deals exist for every inquiry).')
header(lr, 4, ['Lost reason category', 'Inquiries', '% of lost'])
CATS = list(data.get('lost_reason_rules', {}).keys()) + ['Other', 'No reason recorded']
TL = 5 + len(CATS)
for i, c_ in enumerate(CATS):
    r = 5 + i; cell(lr, r, 1, c_); cell(lr, r, 2, f'=COUNTIFS({LRC},A{r})', '#,##0'); cell(lr, r, 3, f'=IF($B${TL}=0,"",B{r}/$B${TL})', '0.0%')
cell(lr, TL, 1, 'Total lost', bold=True); cell(lr, TL, 2, f'=SUM(B5:B{TL - 1})', '#,##0', bold=True); cell(lr, TL, 3, None)
header(lr, TL + 2, ['Lead', 'Package', 'Source', 'Lost reason (as recorded)'])
for i, x in enumerate([x for x in rows if x['status'] in ('Declined', 'Cancelled')]):
    for c, v in enumerate([x['name'], x['package'], x['source'], x['reason'] or '(not recorded)'], 1): cell(lr, TL + 3 + i, c, v)
for c, w in zip('ABCD', [34, 14, 18, 44]): lr.column_dimensions[c].width = w

# ---------------- Inquiries
iq = wb.create_sheet('Inquiries')
title(iq, f'Inquiries – {PER}', 'Source: Rework Prospectors (page Messenger inbox). All counts are formulas on Leads_Data.')
section(iq, 4, 'A. By source', 3); header(iq, 5, ['Source', 'Inquiries', '% of total'])
for i, s_ in enumerate(['Meta Ads', 'Organic Facebook']):
    r = 6 + i; cell(iq, r, 1, s_); cell(iq, r, 2, f'=COUNTIFS({SRC},A{r})', '#,##0'); cell(iq, r, 3, f'=IF($B$8=0,"",B{r}/$B$8)', '0.0%')
cell(iq, 8, 1, 'Total', bold=True); cell(iq, 8, 2, '=SUM(B6:B7)', '#,##0', bold=True); cell(iq, 8, 3, None)
section(iq, 10, 'B. By status', 3); header(iq, 11, ['Status', 'Inquiries', '% of total'])
SL = ['Intake', 'Qualified', 'Pending Payment', 'Paid Booking', 'Declined', 'Cancelled']
for i, s_ in enumerate(SL):
    r = 12 + i; cell(iq, r, 1, s_); cell(iq, r, 2, f'=COUNTIFS({ST},A{r})', '#,##0'); cell(iq, r, 3, f'=IF($B$8=0,"",B{r}/$B$8)', '0.0%')
RVW = 12 + len(SL); cell(iq, RVW, 1, '…of which not yet reviewed'); cell(iq, RVW, 2, f'=COUNTIFS({REV},"No")', '#,##0', src='input'); cell(iq, RVW, 3, f'=IF($B$8=0,"",B{RVW}/$B$8)', '0.0%')
pk0 = RVW + 2; section(iq, pk0, 'C. By package', 3); header(iq, pk0 + 1, ['Package', 'Inquiries', '% of total'])
for i, s_ in enumerate([p['code'] for p in PK] + ['Not specified']):
    r = pk0 + 2 + i; cell(iq, r, 1, s_); cell(iq, r, 2, f'=COUNTIFS({PKG},A{r})', '#,##0'); cell(iq, r, 3, f'=IF($B$8=0,"",B{r}/$B$8)', '0.0%')
h0 = pk0 + 4 + len(PK); section(iq, h0, 'D. By hour of day', 3); header(iq, h0 + 1, ['Hour', 'Inquiries'])
for h in range(24):
    r = h0 + 2 + h; cell(iq, r, 1, f'{h:02d}:00'); nxt = '1' if h == 23 else f'TIME({h + 1},0,0)'
    cell(iq, r, 2, f'=COUNTIFS({R("B")},">="&TIME({h},0,0),{R("B")},"<"&{nxt})', '#,##0')
for c, w in zip('ABC', [30, 12, 12]): iq.column_dimensions[c].width = w

# ---------------- Summary
sm = wb.create_sheet('Summary', 0)
title(sm, f"{data['client']['name']} – Monthly Report", PER + '  ·  Prepared from Rework (Prospectors, Sales, Leads, Meta Ads)')
section(sm, 4, 'Package prices (inputs)', 7); header(sm, 5, ['', 'Package', 'Price', 'Note'])
for i, p in enumerate(PK):
    r = 6 + i; cell(sm, r, 2, p['label']); cell(sm, r, 3, p['price'], '#,##0', src='input'); cell(sm, r, 4, 'Edit if a promo price applies')
MIX = 6 + len(PK)
known = '+'.join(f'COUNTIFS({PKG},"{p["code"]}")' for p in PK)
cell(sm, MIX, 2, 'Average price where package not recorded')
cell(sm, MIX, 3, '=IFERROR((' + '+'.join(f'COUNTIFS({PKG},"{p["code"]}")*$C${6 + i}' for i, p in enumerate(PK)) + f')/({known}),AVERAGE($C$6:$C${MIX - 1}))', '#,##0')
cell(sm, MIX, 4, 'Weighted by inquiries with a known package')
def value_of(extra=''):
    parts = [f'COUNTIFS({PKG},"{p["code"]}"{extra})*$C${6 + i}' for i, p in enumerate(PK)]
    parts.append(f'COUNTIFS({PKG},"Not specified"{extra})*$C${MIX}')
    return '=' + '+'.join(parts)
k0 = MIX + 2
section(sm, k0, 'Client KPIs – the five questions agreed for monthly reporting', 7)
header(sm, k0 + 1, ['#', 'Question', 'KPI', PREV, 'This month', 'Change', 'Data source'])
S = k0 + 2
Q = [
 (1, 'How many inquiries came in?', 'Total inquiries', prev.get('inquiries'), '=Inquiries!B8', 'auto'),
 (1, '', '…from Meta Ads', prev.get('meta'), '=Inquiries!B6', 'auto'),
 (1, '', '…from Organic Facebook', prev.get('organic'), '=Inquiries!B7', 'auto'),
 (2, 'What are they worth if all of them buy a package?', 'Potential value', None, value_of(), 'mixed'),
 (3, 'How many became paying customers?', 'Paid bookings', prev.get('paid'), f'=COUNTIFS({ST},"Paid Booking")', 'mixed'),
 (3, '', 'Revenue from paid bookings', None, value_of(f',{ST},"Paid Booking"'), 'mixed'),
 (3, '', 'Pending payment', None, f'=COUNTIFS({ST},"Pending Payment")', 'mixed'),
 (3, '', 'Qualified (incl. later stages)', None, f'=COUNTIFS({ST},"Qualified")+COUNTIFS({ST},"Pending Payment")+COUNTIFS({ST},"Paid Booking")', 'mixed'),
 (4, 'Why did the others not buy?', 'Lost inquiries', None, f"='Lost Reasons'!B{TL}", 'mixed'),
 (4, '', 'Top lost reason', None, f"=IFERROR(INDEX('Lost Reasons'!A5:A{TL - 2},MATCH(MAX('Lost Reasons'!B5:B{TL - 2}),'Lost Reasons'!B5:B{TL - 2},0)),\"\")", 'mixed'),
 (5, 'What did the campaigns cost?', 'Ad spend', None, f"='Campaign Cost'!F{TR}", 'auto'),
 (5, '', 'Cost per inquiry from ads', None, f"='Campaign Cost'!O{TR}", 'auto'),
 (5, '', 'Cost per paid booking', None, f"=IF(E{S + 4}=0,\"\",'Campaign Cost'!F{TR}/E{S + 4})", 'mixed'),
 (5, '', 'Budget used', None, f"=IF({data.get('budget') or 0}=0,\"\",'Campaign Cost'!F{TR}/{data.get('budget') or 1})", 'auto'),
]
for i, (n_, q, k_, pv, f_, src) in enumerate(Q):
    r = S + i
    money = k_ in ('Potential value', 'Revenue from paid bookings', 'Ad spend') or k_.startswith('Cost')
    fmt = '0.0%' if k_ == 'Budget used' else '#,##0.00' if k_.startswith('Cost') else '#,##0'
    cell(sm, r, 1, n_ if q else None); cell(sm, r, 2, q, bold=True, wrap=True); cell(sm, r, 3, k_ + (' (₱)' if money else ''))
    cell(sm, r, 4, pv, fmt); cell(sm, r, 5, f_, fmt, bold=True)
    cell(sm, r, 6, f'=IF(OR(D{r}="",D{r}=0,NOT(ISNUMBER(E{r}))),"",(E{r}-D{r})/D{r})', '0.0%'); tag(sm, r, 7, src)
r0 = S + len(Q) + 1
section(sm, r0, 'Organic Facebook (manual – from Meta Business Suite)', 7)
header(sm, r0 + 1, ['', '', 'KPI', PREV, 'This month', 'Change', 'Data source'])
for i, k in enumerate(OK[:3]):
    r = r0 + 2 + i; cell(sm, r, 1, None); cell(sm, r, 2, None); cell(sm, r, 3, k['label'])
    cell(sm, r, 4, f"='Organic Facebook (Manual)'!B{6 + i}", '#,##0')
    cell(sm, r, 5, f"=IF('Organic Facebook (Manual)'!C{6 + i}=\"\",\"\",'Organic Facebook (Manual)'!C{6 + i})", '#,##0', bold=True)
    cell(sm, r, 6, f'=IF(OR(E{r}="",D{r}=0),"",(E{r}-D{r})/D{r})', '0.0%'); tag(sm, r, 7, 'manual')
r = r0 + 6; section(sm, r, 'Data quality note', 7)
sm.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=7)
sm.cell(r + 1, 1, f'=Inquiries!B{RVW}&" of "&Inquiries!B8&" inquiries have not been reviewed by the sales team yet and default to Intake. Qualified, paid and lost counts, and the package mix behind Potential value, will change once they are updated."').font = font(size=10, italic=True)
sm.cell(r + 1, 1).alignment = Alignment(wrap_text=True, vertical='top'); sm.row_dimensions[r + 1].height = 32
for c, w in zip('ABCDEFG', [5, 34, 32, 12, 14, 10, 20]): sm.column_dimensions[c].width = w

# ---------------- Data Sources
rm = wb.create_sheet('Data Sources', 1)
title(rm, 'Where each number comes from', 'Green = Rework collects it automatically · Yellow = in Rework, but depends on sales-team updates · Orange = copy manually from Meta')
header(rm, 4, ['Report section', 'Metric', 'Source', 'Collected', 'How / where', 'To automate later'])
RM = [
 ('Q1 Inquiries', 'Inquiries by source, week, hour', 'Rework Prospectors (page inbox)', 'auto', 'Every new Messenger conversation on the page; ad vs organic is tagged by Meta.', '—'),
 ('Q1 Inquiries', 'Ad set per inquiry', 'Rework Prospectors + Meta Ads', 'mixed', 'From the ad that opened the conversation; untagged ones assigned by start date (yellow).', 'Create a lead/deal for every inquiry so the ad set is stored on the record.'),
 ('Q2 Potential value', 'Package asked about', 'Sales team', 'mixed', 'Package column on Leads_Data.', 'Add a Package field to the lead service and deal form.'),
 ('Q2 Potential value', 'Package prices', 'Client price list', 'mixed', 'Summary price inputs.', 'Add products/prices to the pipeline.'),
 ('Q3 Paying customers', 'Status', 'Sales team → Rework Sales pipeline', 'mixed', 'Status column on Leads_Data, or move the deal in the pipeline.', 'Deal-stage → lead-stage automation.'),
 ('Q4 Lost reasons', 'Lost reason', 'Sales team → Rework Sales lost reasons', 'mixed', 'Record a reason for every Declined / Cancelled inquiry.', 'Shared lost-reason list prefixed with the BU name.'),
 ('Q5 Campaign cost', 'Spend, impressions, reach, clicks per ad set', 'Rework Meta Ads', 'auto', 'Meta Ads > Ad sets, filtered to the client.', '—'),
 ('Q5 Campaign cost', 'Meta-reported results', 'Meta Ads Manager', 'manual', 'Ads Manager > Ad sets > Results, same dates.', '—'),
 ('Organic', 'Page views, reach, interactions, clicks, visits, follows', 'Meta Business Suite', 'manual', 'Insights > Overview.', 'Not available in Rework.'),
 ('Organic', 'Posts and their performance', 'Meta Business Suite', 'manual', 'Content > Posts.', 'Not available in Rework.'),
 ('Organic', 'Followers, age, top cities', 'Meta Business Suite', 'manual', 'Insights > Audience.', 'Not available in Rework.'),
]
for i, (a, b, c, s_, h, t) in enumerate(RM):
    r = 5 + i
    for col, v in enumerate([a, b, c], 1): cell(rm, r, col, v, wrap=True)
    tag(rm, r, 4, s_); cell(rm, r, 5, h, wrap=True); cell(rm, r, 6, t, wrap=True)
for c, w in zip('ABCDEF', [20, 36, 30, 20, 52, 44]): rm.column_dimensions[c].width = w

order = ['Summary', 'Data Sources', 'Inquiries', 'Campaign Cost', 'Lost Reasons', 'Organic Facebook (Manual)', D]
wb._sheets = [wb[n] for n in order]
for ws in wb:
    ws.sheet_properties.tabColor = {'Organic Facebook (Manual)': MANT, 'Data Sources': '7F7F7F'}.get(ws.title, NAVY)
wb.save(OUT); print('saved', OUT, len(rows), 'lead rows')
