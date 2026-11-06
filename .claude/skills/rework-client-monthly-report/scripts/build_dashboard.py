#!/usr/bin/env python3
"""Build the client-ready 5-tab HTML dashboard from data.json (output of prepare_data.py).

Usage: python3 -I build_dashboard.py data.json out.html

Every number is computed from data.json. A value that is null / unknown shows as a TBU pill.
Narrative text (headline, insights, plan) comes from config.narrative - write it AFTER looking at
the numbers this script prints, then re-run.
Tabs: 01 Overview · 02 Lead Generation · 03 Conversion & Revenue · 04 Channel Performance · 05 Insights & Next Steps
"""
import sys, json, os, calendar, datetime, html
from collections import Counter

d = json.load(open(sys.argv[1])); OUT = sys.argv[2]
HERE = os.path.dirname(os.path.abspath(__file__))
head = open(os.path.join(HERE, '..', 'assets', 'dashboard_head.html'), encoding='utf-8').read()

cl = d['client']; PER = d['period_label']; PREV = d.get('prev_label', 'Prev.'); PS = PREV[:3]
N = d.get('narrative', {}); L = d['leads']; A = d['adsets']; PK = d['packages']; prev = d.get('previous', {})
y, m = map(int, d['month'].split('-')); ND = calendar.monthrange(y, m)[1]; MON = calendar.month_abbr[m]
TBU = '<span class="tbu">TBU</span>'
esc = html.escape

def num(v, dec=0): return TBU if v is None else f'{v:,.{dec}f}'
def peso(v, dec=0): return TBU if v is None else f'₱{v:,.{dec}f}'
def pct(a, b, dec=0): return TBU if (a is None or not b) else f'{a / b * 100:.{dec}f}%'
def short_money(v):
    if v is None: return TBU
    return f'₱{v / 1e6:.2f}M' if v >= 1e6 else f'₱{v / 1e3:.0f}K' if v >= 1e5 else f'₱{v:,.0f}'
def change(cur, pv):
    if pv in (None, 0) or cur is None: return ''
    c = (cur - pv) / pv * 100
    return f'<span class="{"up" if c >= 0 else "down"}">{"▲" if c >= 0 else "▼"} {abs(c):.0f}%</span>'

# ---------------- aggregates
tot = len(L); meta = sum(1 for x in L if x['source'] == 'Meta Ads'); org = tot - meta
st = Counter(x['status'] for x in L)
paid = st['Paid Booking']; pend = st['Pending Payment']; qual = st['Qualified'] + pend + paid; pay_stage = pend + paid
lost = st['Declined'] + st['Cancelled']; unreviewed = sum(1 for x in L if x['reviewed'] == 'No')
# nothing reviewed and no deal moved -> conversion is unknown, not zero
CONV = unreviewed < tot or any(x['status'] != 'Intake' for x in L)
price = {p['code']: p['price'] for p in PK}
pkc = Counter(x['package'] for x in L)
known = sum(pkc[c] for c in price)
avg_unknown = (sum(pkc[c] * price[c] for c in price) / known) if known else sum(price.values()) / len(price)
def value(rows): return sum(price.get(x['package'], avg_unknown) for x in rows)
potential = value(L); revenue = value([x for x in L if x['status'] == 'Paid Booking'])
if not CONV: paid = pend = qual = pay_stage = lost = revenue = None
pend_value = value([x for x in L if x['status'] == 'Pending Payment']) if CONV else None
spend = sum(a['spend'] for a in A); impr = sum(a['impressions'] for a in A); reach = sum(a['reach'] for a in A); clicks = sum(a['clicks'] for a in A)
ctr = clicks / impr * 100 if impr else None
cpi = spend / meta if meta else None
paid_ads = sum(1 for x in L if x['status'] == 'Paid Booking' and x['source'] == 'Meta Ads')
cppb = spend / paid_ads if (paid_ads and CONV) else None
budget = d.get('budget')
inq_by_set = Counter(x['adset'] for x in L if x['adset'])
paid_by_set = Counter(x['adset'] for x in L if x['adset'] and x['status'] == 'Paid Booking')
lost_cat = Counter(x['reason_category'] for x in L if x['reason_category'])

daily_meta = [0] * ND; daily_org = [0] * ND; hours = [0] * 24
for x in L:
    if not x['date']: continue
    dt = datetime.date.fromisoformat(x['date'])
    if (dt.year, dt.month) != (y, m): continue
    (daily_meta if x['source'] == 'Meta Ads' else daily_org)[dt.day - 1] += 1
    if x['time']: hours[int(x['time'][:2])] += 1
# weeks = Monday-start calendar weeks clipped to the month
weeks, day = [], 1
while day <= ND:
    wd = datetime.date(y, m, day).weekday(); end = min(ND, day + (6 - wd))
    lab = f'{MON} {day}–{end}' if end != day else f'{MON} {day}'
    weeks.append([lab, sum(daily_meta[day - 1:end]), sum(daily_org[day - 1:end])]); day = end + 1
peaks = sorted(range(24), key=lambda h: -hours[h])[:2]

print(f'inquiries {tot} (meta {meta}, organic {org}) | qualified {qual} payment {pay_stage} paid {paid} lost {lost} unreviewed {unreviewed}')
print(f'potential {potential:,.0f} revenue {revenue} | spend {spend:,.2f} cpi {cpi and round(cpi, 2)} ctr {ctr and round(ctr, 2)} reach {reach:,}')
print('by ad set:', {a['short']: (inq_by_set[a['adset']], round(a['spend'] / inq_by_set[a['adset']], 2) if inq_by_set[a['adset']] else None) for a in A})
print('weeks:', weeks, '| peak hours:', [(h, hours[h]) for h in peaks], '| lost:', dict(lost_cat), '| packages:', dict(pkc))

# ---------------- html pieces
def kpi(color, label, val, sub): return f'<div class="kpi"><span class="l"><i style="background:var(--{color})"></i>{label}</span><span class="v">{val}</span><span class="d">{sub}</span></div>'
def footer(src, i): return f'<footer><span>{src}</span><span>0{i} / 05</span></footer>'
GRAD = ['#9c90f5,#7b6cf0', '#8fbaf8,#6aa2f6', '#86d79e,#58c37a', '#f6c08d,#f1a35b']
def funnel(stages):
    out = []
    for i, (lab, v, note) in enumerate(stages):
        w = [100, 78, 56, 36][i]
        out.append(f'<div class="fz"><div style="width:{w}%;background:linear-gradient(180deg,{GRAD[i]})"></div></div>'
                   f'<div class="flab"><span>{lab}</span><b>{num(v)}</b>{f"<em>{note}</em>" if note else ""}</div>')
    return '<div class="funnel">' + ''.join(out) + '</div>'
def hbars(items, color):
    mx = max([v for _, v, _ in items] + [1])
    return '<div class="hbars">' + ''.join(
        f'<div class="hb"><span>{esc(lab)}</span><div class="tr"><div class="fi" style="width:{v / mx * 100:.1f}%;background:{c or color}"></div></div><b>{v}</b></div>'
        for lab, v, c in items) + '</div>'
def fdate(s): return datetime.datetime.fromisoformat(s)
def drange(a):
    s, e = fdate(a.get('display_start') or a['start_at']), fdate(a['end_at'])
    return f'{calendar.month_abbr[s.month]} {s.day}–{e.day}' if s.month == e.month else f'{calendar.month_abbr[s.month]} {s.day} – {calendar.month_abbr[e.month]} {e.day}'

up_prev = f'{change(tot, prev.get("inquiries"))}vs. {prev.get("inquiries")} in {PS}' if prev.get('inquiries') else ''
budget_txt = f'{pct(spend, budget)} of {peso(budget)} budget' if budget else 'Total for the month'
nar = lambda k, default='': N.get(k) or default

# 01 overview
COLORS = ['purple', 'green', 'blue', 'orange']
ins = ''.join(f'<li><span class="ic" style="background:var(--{COLORS[i % 4]}-soft);color:var(--{COLORS[i % 4]})">{i + 1}</span><div><b>{esc(t["title"])}</b><span>{esc(t["text"])}</span></div></li>'
              for i, t in enumerate(N.get('insights', [])))
p1 = f'''
<section class="page" id="overview">
  <div class="title"><h1>{esc(nar("headline", f"{cl['short']} – {PER}"))}</h1><p>{esc(nar("subhead", f"{tot} inquiries this month."))}</p></div>
  <div class="grid kpis">
    {kpi("purple", "Inquiries", num(tot), up_prev)}
    {kpi("green", "Potential value", short_money(potential), "If every inquiry booked")}
    {kpi("orange", "Paid bookings", num(paid), f"{peso(revenue)} · {num(pend)} pending payment" if CONV else "Sales team to update")}
    {kpi("blue", "Ad spend", peso(spend), budget_txt)}
    {kpi("gray", "Cost per inquiry", peso(cpi, 2), "From Meta Ads")}
  </div>
  <div class="card">
    <div class="chead"><h3>Performance Overview</h3>
      <div class="seg" role="group" aria-label="Chart view"><button type="button" data-mode="cum" aria-pressed="true">Cumulative</button><button type="button" data-mode="daily" aria-pressed="false">Daily</button></div>
    </div>
    <div class="perf">
      <div class="series">
        <div><div class="l"><i style="background:var(--purple)"></i>Total inquiries</div><div class="v">{tot}</div></div>
        <div><div class="l"><i style="background:var(--blue)"></i>Meta Ads</div><div class="v">{meta}</div></div>
        <div><div class="l"><i style="background:var(--green)"></i>Organic Facebook</div><div class="v">{org}</div></div>
      </div>
      <svg id="perf" viewBox="0 0 760 300" role="img" aria-label="Inquiries over the month; cumulative total reaches {tot}"></svg>
    </div>
  </div>
  <div class="grid g2">
    <div class="card">
      <div class="chead"><h3>Inquiry Source</h3></div>
      <div class="donut">
        <svg viewBox="0 0 220 220" role="img" aria-label="Meta Ads {pct(meta, tot)}, Organic Facebook {pct(org, tot)}">
          <defs><filter id="glow" x="-20%" y="-20%" width="140%" height="140%"><feGaussianBlur stdDeviation="6" result="b"/><feMerge><feMergeNode in="b"/><feMergeNode in="SourceGraphic"/></feMerge></filter></defs>
          <g id="donut" transform="translate(110 110)"></g>
          <text x="110" y="108" text-anchor="middle" style="font-size:28px;font-weight:800;fill:var(--ink)">{tot}</text>
          <text x="110" y="130" text-anchor="middle" style="font-size:12px">Total inquiries</text>
        </svg>
        <div class="lg">
          <div><i style="background:var(--blue)"></i><span>Meta Ads</span><b>{pct(meta, tot)}</b></div>
          <div><i style="background:var(--green)"></i><span>Organic Facebook</span><b>{pct(org, tot)}</b></div>
        </div>
      </div>
    </div>
    <div class="card">
      <div class="chead"><h3>Leads Funnel</h3></div>
      {funnel([("Inquiries", tot, ""), ("Qualified", qual, (f"{pct(qual, tot)} of inquiries" if CONV else "")), ("Payment stage", pay_stage, (f"{pct(pay_stage, qual)} of qualified" if CONV else "")), ("Paid bookings", paid, (f"{peso(revenue)} revenue" if CONV else ""))])}
      {f'<p class="note">{unreviewed} of {tot} inquiries are still in follow-up; qualified and paid counts are {TBU}.</p>' if unreviewed else ''}
    </div>
  </div>
  {f'<div class="card"><div class="chead"><h3>Key Insights</h3></div><ul class="ins">{ins}</ul></div>' if ins else ''}
  {footer(f"Sources: Rework (Messenger inbox, Sales pipeline, Meta Ads) · Meta Business Suite · As of {d.get('as_of', '')}", 1)}
</section>'''

# 02 leads
pk_h = sorted(range(24), key=lambda h: -hours[h])[:2]
p2 = f'''
<section class="page" id="leads" hidden>
  <div class="title"><h2>Lead Generation</h2><p>{esc(nar("leads_intro"))}</p></div>
  <div class="grid g3">
    {kpi("purple", "Total inquiries", num(tot), up_prev)}
    {kpi("blue", "From Meta Ads", num(meta), f"{pct(meta, tot)}" + (f" · {PS}: {prev['meta']}" if prev.get('meta') is not None else ""))}
    {kpi("green", "From organic Facebook", num(org), f"{pct(org, tot)}" + (f" · {PS}: {prev['organic']}" if prev.get('organic') is not None else ""))}
  </div>
  <div class="card">
    <div class="chead"><h3>Weekly Trend</h3><div class="lg" style="flex-direction:row;gap:18px"><div style="grid-template-columns:14px auto"><i style="background:var(--blue)"></i><span>Meta Ads</span></div><div style="grid-template-columns:14px auto"><i style="background:var(--green)"></i><span>Organic</span></div></div></div>
    <svg id="weekly" viewBox="0 0 760 260" role="img" aria-label="Weekly inquiries: {', '.join(f'{w[0]}: {w[1] + w[2]}' for w in weeks)}"></svg>
  </div>
  <div class="card">
    <div class="chead"><h3>Peak Hours</h3></div>
    <svg id="hours" viewBox="0 0 760 170" role="img" aria-label="Inquiries by hour, peaking at {pk_h[0]:02d}:00 ({hours[pk_h[0]]}) and {pk_h[1]:02d}:00 ({hours[pk_h[1]]})"></svg>
    <p class="note">{esc(nar("hours_note", f"Busiest hours: {pk_h[0]:02d}:00 and {pk_h[1]:02d}:00. Fast replies at these hours have the most impact."))}</p>
  </div>
  {footer(f"Source: Rework – {cl.get('page', cl['short'] + ' Facebook page')} Messenger inbox", 2)}
</section>'''

# 03 conversion
pk_colors = ['var(--purple)', 'var(--blue)', 'var(--green)', 'var(--orange)']
pk_items = [(f"{p['code']} ({peso(p['price'])})", pkc[p['code']], pk_colors[i % 4]) for i, p in enumerate(PK)] + [('Not yet confirmed', pkc['Not specified'], '#d9dce6')]
mix = ', '.join(f"{pkc[c] / known * 100:.0f}% {c}" for c in price) if known else (f"{PK[0]['label'].lower()} price" if len(PK) == 1 else 'the average list price')
lost_items = [(c, n, '#d9dce6' if c == 'No reason recorded' else None) for c, n in lost_cat.most_common()]
p3 = f'''
<section class="page" id="conversion" hidden>
  <div class="title"><h2>Conversion &amp; Revenue</h2><p>Potential value assumes every inquiry books one package at list price: {", ".join(f"{p['code']} {peso(p['price'])}" for p in PK)}.</p></div>
  <div class="grid g3">
    {kpi("green", "Potential value", peso(potential), f"{tot} inquiries at list price")}
    {kpi("orange", "Revenue from paid bookings", peso(revenue), (f"{paid} paid booking{'s' if paid != 1 else ''}" if CONV else "Sales team to update"))}
    {kpi("purple", "Pending payment", num(pend), f"Up to {peso(pend_value)} more" if pend else "")}
  </div>
  <div class="grid g2">
    <div class="card">
      <div class="chead"><h3>Pipeline Funnel</h3></div>
      {funnel([("Inquiries", tot, ""), ("Qualified (incl. later stages)", qual, (f"{pct(qual, tot)} of inquiries" if CONV else "")), ("Pending payment + paid", pay_stage, (f"{pct(pay_stage, qual)} of qualified" if CONV else "")), ("Paid bookings", paid, f"{paid} of {pay_stage} at payment stage" if CONV else "")])}
    </div>
    <div class="card">
      <div class="chead"><h3>Package Interest</h3></div>
      {hbars(pk_items, 'var(--purple)')}
      {f'<p class="note">Package for {pkc["Not specified"]} inquiries is {TBU}. Potential value uses ' + (f'the known mix for these ({mix}).' if known else f'the {mix} for these.') + '</p>' if pkc['Not specified'] else ''}
    </div>
  </div>
  <div class="card">
    <div class="chead"><h3>Lost Reasons</h3><span class="note">{num(lost)} lost inquiries · open ones {TBU}</span></div>
    {hbars(lost_items, 'var(--orange)') if lost_items else f'<p class="note">No lost inquiries recorded yet {TBU}.</p>'}
    {f'<p class="note">{esc(N["lost_detail"])}</p>' if N.get('lost_detail') else ''}
  </div>
  {footer(f"Sources: Rework – Messenger inbox, Sales pipeline · {cl['short']} price list", 3)}
</section>'''

# 04 channels
DOT = ['green', 'purple', 'blue', 'orange', 'gray']
rows_sorted = sorted(A, key=lambda a: (a['spend'] / inq_by_set[a['adset']]) if inq_by_set[a['adset']] else 1e12)
best = rows_sorted[0]['adset'] if rows_sorted and inq_by_set[rows_sorted[0]['adset']] else None
tr = []
for i, a in enumerate(rows_sorted):
    n = inq_by_set[a['adset']]; c = a['spend'] / n if n else None
    tr.append(f'<tr><td><span class="dot" style="background:var(--{DOT[i % 5]})"></span>{esc(a.get("short") or a["adset"])}</td><td class="nw">{drange(a)}</td>'
              f'<td class="r">{peso(a["spend"], 2)}</td><td class="r">{num(a["reach"])}</td><td class="r">{pct(a["clicks"], a["impressions"], 2)}</td>'
              f'<td class="r">{n}</td><td class="r{" best" if a["adset"] == best else ""}">{peso(c, 2) if c else "—"}</td><td class="r">{num(a.get("meta_results"))}</td></tr>')
mr = [a.get('meta_results') for a in A]
tr.append(f'<tr class="total"><td>Total</td><td></td><td class="r">{peso(spend, 2)}</td><td class="r">{num(reach)}</td><td class="r">{pct(clicks, impr, 2)}</td>'
          f'<td class="r">{sum(inq_by_set[a["adset"]] for a in A)}</td><td class="r">{peso(cpi, 2)}</td><td class="r">{num(sum(mr)) if all(v is not None for v in mr) else TBU}</td></tr>')
ok = d['organic']['kpis']
org_rows = ''.join(f'<tr><td>{esc(k["label"])}</td><td class="r">{num(k["prev"])}</td><td class="r">{num(k["current"])}</td><td class="r">{change(k["current"], k["prev"]) or TBU}</td></tr>' for k in ok)
p4 = f'''
<section class="page" id="channels" hidden>
  <div class="title"><h2>Channel Performance</h2><p>{esc(nar("channels_intro", f"{len(A)} ad sets ran this month."))}</p></div>
  <div class="grid kpis">
    {kpi("blue", "Ad spend", peso(spend), budget_txt)}
    {kpi("purple", "Ad reach", num(reach), "Summed per ad set")}
    {kpi("green", "Avg. CTR", pct(clicks, impr, 2), f"{impr:,} impressions")}
    {kpi("gray", "Cost per inquiry", peso(cpi, 2), f"{meta} inquiries from ads")}
    {kpi("orange", "Cost per paid booking", peso(cppb), (f"{paid_ads} paid booking{'s' if paid_ads != 1 else ''} from ads" if CONV else "Sales team to update"))}
  </div>
  <div class="card">
    <div class="chead"><h3>Meta Ads · Campaign Efficiency</h3></div>
    <div class="tablewrap"><table>
      <thead><tr><th>Ad set</th><th>Dates</th><th class="r">Spend</th><th class="r">Reach</th><th class="r">CTR</th><th class="r">Inquiries*</th><th class="r">Cost / inquiry</th><th class="r">Meta results</th></tr></thead>
      <tbody>{''.join(tr)}</tbody>
    </table></div>
    <p class="note">*Messenger conversations started from each ad set; where Meta did not tag the ad, the conversation is assigned to the ad set running that day.{(" " + esc(N["adsets_note"])) if N.get("adsets_note") else ""}</p>
  </div>
  <div class="card">
    <div class="chead"><h3>Organic Facebook</h3><span class="note">Figures from Meta Business Suite{"" if all(k["current"] is not None for k in ok) else " " + TBU}</span></div>
    <div class="tablewrap"><table>
      <thead><tr><th>KPI</th><th class="r">{esc(PREV)}</th><th class="r">This month</th><th class="r">Change</th></tr></thead>
      <tbody>{org_rows}</tbody>
    </table></div>
    {f'<p class="note">{esc(d["organic"]["prev_note"])}</p>' if d['organic'].get('prev_note') else ''}
  </div>
  {footer(f"Sources: Rework – Meta Ads (synced {d.get('as_of', '')}) · Meta Ads Manager · Meta Business Suite", 4)}
</section>'''

# 05 next
li = lambda xs: ''.join(f'<li>{x}</li>' for x in xs) or f'<li>{TBU}</li>'   # worked/attention allow <b> markup
plan = ''.join(f'<li><div><b>{esc(p["title"])}</b>{esc(p["text"])}</div><span class="when">{esc(p.get("when", ""))}</span></li>' for p in N.get('plan', []))
p5 = f'''
<section class="page" id="next" hidden>
  <div class="title"><h2>Insights &amp; Next Steps</h2><p>{esc(nar("next_intro", "What to keep, what to fix, and the plan for next month."))}</p></div>
  <div class="grid g2">
    <div class="card col"><h4 style="color:#2f9a55">What worked</h4><ul>{li(N.get("worked", []))}</ul></div>
    <div class="card col"><h4 style="color:#c97a2f">Needs attention</h4><ul>{li(N.get("attention", []))}</ul></div>
  </div>
  <div class="card"><div class="chead"><h3>Next Month's Action Plan</h3></div><ol class="plan">{plan or f"<li><div>{TBU}</div></li>"}</ol></div>
  {footer(f"Prepared for {esc(cl['name'])}", 5)}
</section>'''

body = f'''
<body>
<div class="wrap">
<header class="masthead">
  <div class="brand"><div class="logo" style="font-size:{13 if len(cl.get('logo', cl['short'])) > 3 else 15}px">{esc(cl.get('logo', cl['short'])[:4])}</div><div><div class="t">{esc(cl['name'])}</div><div class="s">Monthly Marketing &amp; Sales Report</div></div></div>
  <div class="period"><i></i>{esc(PER)}</div>
</header>
<nav aria-label="Report sections">
  <a href="#overview"><span class="n">01</span><span class="h">Overview</span><span class="d">Executive KPIs, monthly highlights, key insights</span></a>
  <a href="#leads"><span class="n">02</span><span class="h">Lead Generation</span><span class="d">Inquiries, sources, weekly trend, peak hours</span></a>
  <a href="#conversion"><span class="n">03</span><span class="h">Conversion &amp; Revenue</span><span class="d">Funnel, paid bookings, packages, lost reasons</span></a>
  <a href="#channels"><span class="n">04</span><span class="h">Channel Performance</span><span class="d">Meta Ads, campaign efficiency, organic Facebook</span></a>
  <a href="#next"><span class="n">05</span><span class="h">Insights &amp; Next Steps</span><span class="d">What worked, what needs attention, action plan</span></a>
</nav>
{p1}{p2}{p3}{p4}{p5}
</div>'''

DATA = dict(meta=daily_meta, org=daily_org, hours=hours, weeks=weeks, peaks=peaks, mon=MON, nd=ND,
            ann=d.get('annotations', []), donut=[meta, org])
script = '''
<script>
(function(){
  var D=__DATA__;
  var NS='http://www.w3.org/2000/svg';
  function el(n,a,t){var e=document.createElementNS(NS,n);for(var k in a)e.setAttribute(k,a[k]);if(t!=null)e.textContent=t;return e;}
  var C={purple:'#7b6cf0',blue:'#6aa2f6',green:'#58c37a',orange:'#f1a35b',grid:'#f0f1f5',axis:'#e3e5ec'};
  var meta=D.meta,org=D.org,ND=D.nd;
  function cum(a){var s=0;return a.map(function(v){return s+=v;});}
  function nice(v,n){var raw=v/n,p=Math.pow(10,Math.floor(Math.log10(raw||1))),f=raw/p;var s=(f<=1?1:f<=2?2:f<=2.5?2.5:f<=5?5:10)*p;return s;}
  function smooth(pts){var d='M'+pts[0][0]+','+pts[0][1];
    for(var i=1;i<pts.length;i++){var p0=pts[i-1],p1=pts[i],cx=(p0[0]+p1[0])/2;d+=' C'+cx+','+p0[1]+' '+cx+','+p1[1]+' '+p1[0]+','+p1[1];}
    return d;}
  var svg=document.getElementById('perf');
  function drawPerf(mode){
    while(svg.firstChild)svg.removeChild(svg.firstChild);
    var defs=el('defs',{});svg.appendChild(defs);
    [['gP',C.purple],['gB',C.blue],['gG',C.green]].forEach(function(g){var lg=el('linearGradient',{id:g[0],x1:0,x2:0,y1:0,y2:1});lg.appendChild(el('stop',{offset:'0%','stop-color':g[1],'stop-opacity':.22}));lg.appendChild(el('stop',{offset:'100%','stop-color':g[1],'stop-opacity':0}));defs.appendChild(lg);});
    var tot=meta.map(function(v,i){return v+org[i];});
    var S=mode==='cum'?[[cum(tot),C.purple,'gP'],[cum(meta),C.blue,'gB'],[cum(org),C.green,'gG']]:[[tot,C.purple,'gP'],[meta,C.blue,'gB'],[org,C.green,'gG']];
    var top=Math.max.apply(null,S[0][0].concat([1])),stepY=nice(top,4),max=Math.ceil(top/stepY)*stepY,L=40,R=750,T=14,B=262,W=ND-1;
    for(var v=0;v<=max;v+=stepY){var y=B-v/max*(B-T);svg.appendChild(el('line',{x1:L,x2:R,y1:y,y2:y,stroke:C.grid}));svg.appendChild(el('text',{x:L-10,y:y+4,'text-anchor':'end'},v));}
    for(var d=0;d<ND;d++){var x=L+d/W*(R-L);svg.appendChild(el('line',{x1:x,x2:x,y1:T,y2:B,stroke:C.grid,'stroke-dasharray':'2 4'}));}
    [1,8,15,22,29].filter(function(d){return d<=ND;}).forEach(function(d){var x=L+(d-1)/W*(R-L);svg.appendChild(el('text',{x:x,y:B+22,'text-anchor':'middle'},D.mon+' '+(d<10?'0':'')+d));});
    S.forEach(function(s){
      var pts=s[0].map(function(v,i){return [L+i/W*(R-L),B-v/max*(B-T)];});
      var p=smooth(pts);
      svg.appendChild(el('path',{d:p+' L'+R+','+B+' L'+L+','+B+' Z',fill:'url(#'+s[2]+')'}));
      svg.appendChild(el('path',{d:p,fill:'none',stroke:s[1],'stroke-width':2.5,'stroke-linecap':'round'}));
      var last=pts[pts.length-1];svg.appendChild(el('circle',{cx:last[0],cy:last[1],r:4,fill:'#fff',stroke:s[1],'stroke-width':2.5}));
    });
    D.ann.forEach(function(a){var x=L+(a.day-1)/W*(R-L);svg.appendChild(el('line',{x1:x,x2:x,y1:T,y2:B,stroke:C.purple,'stroke-opacity':.5,'stroke-dasharray':'4 4'}));svg.appendChild(el('text',{x:x+6,y:T+10,'class':'ann'},a.label));});
  }
  drawPerf('cum');
  document.querySelectorAll('.seg button').forEach(function(b){b.addEventListener('click',function(){document.querySelectorAll('.seg button').forEach(function(x){x.setAttribute('aria-pressed',x===b);});drawPerf(b.dataset.mode);});});
  var dg=document.getElementById('donut'),parts=[[D.donut[0],C.blue],[D.donut[1],C.green]].filter(function(p){return p[0]>0;}),r=84,w=26,tot=D.donut[0]+D.donut[1],a0=-Math.PI/2,gap=parts.length>1?.03:0;
  parts.forEach(function(p){var a=p[0]/tot*Math.PI*2,s=a0+gap/2,e=a0+a-gap/2-(parts.length===1?0.0001:0),large=(e-s)>Math.PI?1:0;
    var x1=Math.cos(s)*r,y1=Math.sin(s)*r,x2=Math.cos(e)*r,y2=Math.sin(e)*r;
    dg.appendChild(el('path',{d:'M'+x1+','+y1+' A'+r+','+r+' 0 '+large+' 1 '+x2+','+y2,fill:'none',stroke:p[1],'stroke-width':w,'stroke-linecap':'round',filter:'url(#glow)','stroke-opacity':.9}));a0+=a;});
  var wk=D.weeks,ws=document.getElementById('weekly');
  var L=40,R=750,T=20,B=222,wtop=Math.max.apply(null,wk.map(function(w){return w[1]+w[2];}).concat([1])),wst=nice(wtop,3),max=Math.ceil(wtop/wst)*wst,step=(R-L)/wk.length,cw=Math.min(70,step*0.6);
  for(var v=0;v<=max;v+=wst){var y=B-v/max*(B-T);ws.appendChild(el('line',{x1:L,x2:R,y1:y,y2:y,stroke:C.grid}));ws.appendChild(el('text',{x:L-10,y:y+4,'text-anchor':'end'},v));}
  wk.forEach(function(w,i){var x=L+i*step+(step-cw)/2,hm=w[1]/max*(B-T),ho=w[2]/max*(B-T);
    ws.appendChild(el('rect',{x:x,y:B-hm,width:cw,height:hm,fill:C.blue,rx:hm>0&&ho===0?6:0}));
    ws.appendChild(el('rect',{x:x,y:B-hm-ho,width:cw,height:ho,fill:C.green,rx:6}));
    if(hm>0&&ho>0)ws.appendChild(el('rect',{x:x,y:B-hm,width:cw,height:Math.min(6,hm),fill:C.blue}));
    ws.appendChild(el('text',{x:x+cw/2,y:B-hm-ho-8,'text-anchor':'middle','class':'val'},w[1]+w[2]));
    ws.appendChild(el('text',{x:x+cw/2,y:B+22,'text-anchor':'middle'},w[0]));});
  var h=D.hours,hs=document.getElementById('hours'),hb=134,bw=22,gp=8.6,hmax=Math.max.apply(null,h.concat([1]));
  h.forEach(function(v,i){var x=14+i*(bw+gp),hh=v/hmax*(hb-18),pk=D.peaks.indexOf(i)>=0;
    hs.appendChild(el('rect',{x:x,y:hb-hh,width:bw,height:Math.max(hh,1),rx:5,fill:pk?C.purple:'#d8dcf3'}));
    if(i%3===0)hs.appendChild(el('text',{x:x+bw/2,y:hb+20,'text-anchor':'middle'},(i<10?'0':'')+i+':00'));
    if(pk)hs.appendChild(el('text',{x:x+bw/2,y:hb-hh-6,'text-anchor':'middle','class':'val'},v));});
  var ids=['overview','leads','conversion','channels','next'];
  function show(){var id=(location.hash||'#overview').slice(1);if(ids.indexOf(id)<0)id='overview';
    ids.forEach(function(p){document.getElementById(p).hidden=(p!==id);});
    document.querySelectorAll('nav a').forEach(function(a){if(a.getAttribute('href')==='#'+id)a.setAttribute('aria-current','page');else a.removeAttribute('aria-current');});
    window.scrollTo(0,0);}
  window.addEventListener('hashchange',show);show();
})();
</script>
</body>
</html>
'''.replace('__DATA__', json.dumps(DATA))

title = f"{cl['short']} Monthly Report – {calendar.month_name[m]} {y}"
open(OUT, 'w', encoding='utf-8').write(head.replace('{{TITLE}}', esc(title)) + body + script)
print('saved', OUT)
