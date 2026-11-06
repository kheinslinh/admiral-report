#!/usr/bin/env python3
"""Merge Rework exports into one data.json for the monthly client report.

Inputs (all saved from Rework MCP tool results to local files):
  --channels   Prospectors channel list JSON (rework_prospect_channel_list, limit 5000)
  --account    service_account_id of the client's Facebook page (e.g. 1176 = ECI)
  --month      YYYY-MM of the report
  --adsets     adsets.json (see SKILL.md, step 1D) - campaign + ad set metrics from Rework Meta Ads
  --config     config.json (client, prices, previous month, budget, organic, narratives)
  --tracking   optional: the team's manual tracking workbook (.xlsx/.xlsm) with a
               'Lead Generation' sheet (Date, Time, Source, Campaign ID, Full Name, Contact Type,
               Contact Information, Last Contact Date, Last Contact Time, Notes, Status, 3x follow-up,
               Final Status, Reasons for Decline) - its statuses override the inbox defaults
  --deals      optional: Sales deal list JSON (rework_deal_list) for the client's pipeline
  --out        data.json to write

Run with `python3 -I`.
"""
import argparse, datetime, json, re

STAGE_TO_STATUS = {  # Sales stage name keywords -> report status (edit per pipeline)
    'payment': 'Pending Payment', 'confirmation': 'Pending Payment',
    'qualified': 'Qualified', 'quotation': 'Qualified', 'offer': 'Qualified', 'negotiation': 'Qualified',
    'consultation': 'Qualified', 'decision': 'Qualified', 'treatment plan': 'Qualified',
}

def norm(s):
    return re.sub(r'\s+', ' ', str(s or '').strip().lower())

def reason_category(reason, status, rules):
    if status not in ('Declined', 'Cancelled'):
        return ''
    r = reason.lower()
    if not r:
        return 'No reason recorded'
    for cat, words in rules.items():
        if any(w in r for w in words):
            return cat
    return 'Other'

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--channels', required=True)
    ap.add_argument('--account', type=int, required=True)
    ap.add_argument('--month', required=True)
    ap.add_argument('--adsets', required=True)
    ap.add_argument('--config', required=True)
    ap.add_argument('--tracking')
    ap.add_argument('--deals')
    ap.add_argument('--tz', type=float, default=8.0, help='UTC offset of the client (PH = 8)')
    ap.add_argument('--out', required=True)
    a = ap.parse_args()

    cfg = json.load(open(a.config))
    adsets = json.load(open(a.adsets))
    tz = datetime.timezone(datetime.timedelta(hours=a.tz))
    y, m = map(int, a.month.split('-'))

    raw = json.load(open(a.channels))
    chans = raw['channels'] if isinstance(raw, dict) else raw
    page = [c for c in chans if c.get('service_account_id') == a.account]
    def started(c):
        return datetime.datetime.fromtimestamp(c['since'], tz)
    month_ch = sorted([c for c in page if (started(c).year, started(c).month) == (y, m)], key=lambda c: c['since'])
    pm = (y, m - 1) if m > 1 else (y - 1, 12)
    prev_ch = [c for c in page if (started(c).year, started(c).month) == pm]

    # manual tracking (optional)
    tracked = {}
    if a.tracking:
        import openpyxl
        ws = openpyxl.load_workbook(a.tracking, data_only=True)['Lead Generation']
        for r in range(4, ws.max_row + 1):
            name = ws.cell(r, 5).value
            if not name:
                continue
            v = [ws.cell(r, c).value for c in range(1, 17)]
            camp = str(v[3] or '')
            pkg = next((p['code'] for p in cfg['packages'] if camp.upper().startswith(p['code'])), 'Not specified')
            tracked[norm(name)] = dict(date=v[0], time=v[1], source=v[2], campaign_hint=camp, contact_type=v[5], contact=v[6],
                                       notes=v[9] or '', status=v[10] or 'Intake', reason=v[15] or '', package=pkg)

    deals = {}
    if a.deals:
        dj = json.load(open(a.deals))
        for d in dj.get('deals', dj):
            nm = re.sub(r'\s+via\s+.*$|\s*#\d+$', '', d['name'], flags=re.I)
            deals[norm(nm)] = d

    # ad-set attribution by date: among message ad sets live at the start time, the highest spend per day wins
    def window(x):
        s = datetime.datetime.fromisoformat(x['start_at']).replace(tzinfo=tz)
        e = datetime.datetime.fromisoformat(x['end_at']).replace(tzinfo=tz)
        return s, e
    msg_sets = [x for x in adsets if x.get('brings_messages', True)]
    def daily_spend(x):
        s, e = window(x)
        return x['spend'] / max((e - s).total_seconds() / 86400, 0.5)
    hint_map = {}
    for x in adsets:
        for h in x.get('tracking_ids', []):
            hint_map[h] = x['adset']

    rules = cfg.get('lost_reason_rules', {})
    leads, used = [], set()
    for c in month_ch:
        k = norm(c['name']); d = started(c)
        t = tracked.get(k)
        src = 'Meta Ads' if c.get('subtype') == 'ad' else 'Organic Facebook'
        row = dict(date=d.strftime('%Y-%m-%d'), time=d.strftime('%H:%M'), source=src, adset='', name=c['name'],
                   package='Not specified', status='Intake', reason='', contact_type='', contact='',
                   notes='Imported from Rework inbox - status not yet reviewed', reviewed='No', conversation_id=c['id'],
                   adset_estimated=False)
        if t:
            used.add(k)
            row.update(source=t['source'] or src, package=t['package'], status=t['status'], reason=t['reason'],
                       contact_type=t['contact_type'] or '', contact=t['contact'] or '', notes=t['notes'], reviewed='Yes')
            m_ = re.search(r'(\d{15,})', t['campaign_hint'])
            if m_ and m_.group(1) in hint_map:
                row['adset'] = hint_map[m_.group(1)]
        dl = deals.get(k)
        if dl and not t:
            row['notes'] = 'Deal in Sales pipeline'
            if dl.get('status') == 10: row['status'] = 'Paid Booking'
            elif dl.get('status') == -10: row['status'] = 'Declined'
            else:
                sn = str(dl.get('stage_name') or (dl.get('stage') or {}).get('name', '') if isinstance(dl.get('stage'), dict) else dl.get('stage_name') or '').lower()
                row['status'] = next((v for k_, v in STAGE_TO_STATUS.items() if k_ in sn), 'Intake')
        if row['source'] == 'Meta Ads' and not row['adset']:
            # several message ad sets live that day -> the one spending most per day most likely delivered the ad
            live = [x for x in msg_sets if window(x)[0] <= d <= window(x)[1]]
            if live:
                row['adset'] = max(live, key=daily_spend)['adset']
            row['adset_estimated'] = bool(row['adset'])
        row['reason_category'] = reason_category(row['reason'], row['status'], rules)
        leads.append(row)
    # tracked leads that are not in the inbox (e.g. lead-form submissions)
    for k, t in tracked.items():
        if k in used:
            continue
        td = t['date']
        if isinstance(td, str):
            try: td = datetime.datetime.strptime(td.strip().rstrip(','), '%b %d, %Y')
            except ValueError: td = None
        tt = t['time']
        leads.append(dict(date=td.strftime('%Y-%m-%d') if td else None, time=tt.strftime('%H:%M') if hasattr(tt, 'strftime') else '', source=t['source'], adset=hint_map.get((re.search(r'(\d{15,})', t['campaign_hint']) or [None, ''])[1], ''),
                          name=k.title(), package=t['package'], status=t['status'], reason=t['reason'],
                          reason_category=reason_category(t['reason'], t['status'], rules), contact_type=t['contact_type'] or '',
                          contact=t['contact'] or '', notes=t['notes'], reviewed='Yes', conversation_id='', adset_estimated=False))

    prev = cfg.setdefault('previous', {})
    prev.setdefault('inquiries', len(prev_ch))
    prev.setdefault('meta', sum(1 for c in prev_ch if c.get('subtype') == 'ad'))
    prev.setdefault('organic', sum(1 for c in prev_ch if c.get('subtype') != 'ad'))

    out = dict(cfg, month=a.month, leads=leads, adsets=adsets)
    json.dump(out, open(a.out, 'w'), indent=1, ensure_ascii=False, default=str)
    nd = sum(1 for l in leads if not l['date'])
    print(f"{len(leads)} leads ({nd} without inbox date), prev month {prev['inquiries']}, {len(adsets)} ad sets -> {a.out}")

if __name__ == '__main__':
    main()
