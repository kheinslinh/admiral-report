---
name: rework-client-monthly-report
description: Build a client's monthly marketing & sales report from live Rework data (Prospectors inbox, Sales/Leads, Meta Ads) - a raw-data Excel workbook with formulas plus a client-ready 5-tab HTML dashboard (Overview, Lead Generation, Conversion & Revenue, Channel Performance, Insights & Next Steps) with TBU placeholders for what must still be copied from Meta. Use when asked for a monthly report / dashboard / "báo cáo tháng" for an Astra business unit (ECI, FitMD, J-MaVe, ...) or any Rework client whose leads arrive through a Facebook page inbox and Meta Ads.
---

# Rework client monthly report

Produces two files for one client and one month:

| Output | Script | For |
|---|---|---|
| `data.json` | `scripts/prepare_data.py` | merged inquiries + ad sets + config (working file, contains personal data, never commit) |
| `<CLIENT>_Monthly_Report_<Month>_<Year>.xlsx` | `scripts/build_excel.py` | the team: raw lead rows + formula KPIs, colour-coded by data source |
| `<CLIENT>_Report_<Mon><Year>.html` | `scripts/build_dashboard.py` | the client's CEO: 5-tab light dashboard, readable in a few minutes |

The report answers the five questions agreed with the client:
1. How many inquiries came in (by source)? 2. What are they worth if all of them buy? 3. How many became paying customers?
4. Why did the others not buy? 5. What did the campaigns cost?

Work in the scratchpad. Run every script with `python3 -I`. Example inputs for ECI, October 2026 are in
`examples/eci-2026-10/` (config + ad sets; no lead data).

---

## Step 1 - Collect the data from Rework

Use the client's workspace connectors (for Astra: `Astra`, `Astra_-_Prospectors`). Save every large tool result to a
file (the tool-results dump path is fine) - the scripts read files, not chat output.

### 1A. Inquiries - Prospectors inbox (source of truth for counts)
1. `rework_prospect_service_list` → find the Socials service (Astra: 6254).
2. `rework_prospect_channel_list(limit=5000, serviceId=<socials>)` → save the JSON.
   Each channel = one Messenger conversation: `service_account_id` = the Facebook page
   (Astra: ECI 1176, FitMD 1072), `since` = conversation start (unix; client timezone PH = UTC+8),
   `subtype` = `ad` (opened from an ad) or `organic`.
3. Find the page id: group the list by `service_account_id` and check names, or open one channel.
4. `prepare_data.py` keeps only the page's channels whose start falls in the month, and counts the previous month
   automatically.

Do NOT call `rework_prospect_channel_get` for every channel to read the ad title (100+ calls). Ad-set attribution uses
manual tracking IDs first, then the date window (Step 1D).

### 1B. Team tracking workbook (optional, recommended)
If the client keeps a manual sheet (ECI: `Lead Generation` tab of the monthly .xlsm - Date, Time, Source, Campaign ID,
Full Name, Contact Type, Contact Information, ..., Status, ..., Reasons for Decline), pass it with `--tracking`.
Names are matched to inbox conversations (case/space-insensitive); its Status, package and lost reason override the
inbox default (`Intake`, `Not specified`). The Campaign ID column prefix gives the package (e.g. `PCO - 1202509...`)
and the long number maps to an ad set through `tracking_ids`.

### 1C. Sales / Leads
- `rework_pipeline_list` → the client's pipeline; `rework_deal_list(pipeline)` → save JSON, pass with `--deals`
  (deal won → Paid Booking, lost → Declined, for inquiries not in the tracking sheet).
- Lead stages are kept in sync from deals by automations (see "Related setup"), so deals are enough.
- Report statuses used everywhere: `Intake / Qualified / Pending Payment / Paid Booking / Declined / Cancelled`.
  Edit `STAGE_TO_STATUS` in `prepare_data.py` if a pipeline uses other stage names.

### 1D. Meta Ads - ad sets (`adsets.json`)
`rework_metaads_campaign_insights` returns 500 - don't use it. Instead:
1. `rework_metaads_overview` → campaign / ad set / ad tree; keep the ones whose name contains the client
   (e.g. `ECI`) or that are boosted posts of the client's page. `rework_metaads_campaign_list` → start/end times.
2. `rework_analytics_query(engine="ad_performance", groupBy="adgroup_id", measure=<m>, measureOp="sum")` once each for
   `spend`, `impressions`, `reach`, `clicks`. Values are LIFETIME per ad set (filters and `results`/`leads` measures
   return nothing) - fine when the ad set ran inside the month; otherwise note it.
3. Write `adsets.json`, one object per ad set:
```json
{"campaign": "...", "adset": "<full ad set name - id>", "short": "Messenger – November training",
 "objective": "Messages", "start_at": "2026-10-01T00:00", "end_at": "2026-10-19T23:59",
 "display_start": "2026-10-13T00:00", "spend": 389.56, "impressions": 7832, "reach": 7130, "clicks": 284,
 "meta_results": null, "brings_messages": true, "tracking_ids": ["120250944352200043"]}
```
- `brings_messages`: true if the ad opens Messenger (Messages objective, boosted posts with a Send-message button).
  Only these receive date-window attribution. Lead-form / engagement sets get inquiries only through `tracking_ids`.
- Date-window rule: an untagged ad conversation goes to the message ad set whose window contains its start time;
  when several overlap, the one with the highest spend per day wins. Stretch the main Messenger set's `start_at` back to day 1 so early ad
  conversations are not left blank, and put the real start in `display_start` for the dashboard.
- An ad set that runs past month end: multiply spend/impressions/reach/clicks by the share of its days inside the
  month, set `end_at` to the month end, and mark `short` with ` *` (explain in `adsets_note`).
- Conversations started by an ad that is not synced in Rework (e.g. a boosted post created in Meta only) cannot be
  costed - sample a few with `rework_prospect_channel_get` (`referral_sources[].ads_context_data.ad_title`) to check.
- `meta_results` stays `null` → orange cell in Excel, TBU on the dashboard (copied from Ads Manager by the client).

### 1E. config.json (per client, copy the example)
`client {name, short, page}`, `period_label`, `prev_label`, `as_of`, `budget`,
`packages [{code, label, price}]`, `previous {inquiries, meta, organic, paid}` (omit counts to auto-compute from the
inbox), `lost_reason_rules {category: [keywords]}`, `organic {prev_note, kpis [{label, prev, current}], posts []}`
(organic page insights are NOT in Rework - leave `current: null` until the client sends them),
`annotations [{day, label}]` (events on the trend chart), `narrative` (see Step 3).

---

## Step 2 - Run the scripts

```bash
SK=<repo>/.claude/skills/rework-client-monthly-report/scripts
python3 -I $SK/prepare_data.py --channels channels.json --account 1176 --month 2026-10 \
  --adsets adsets.json --config config.json [--tracking tracking.xlsm] [--deals deals.json] --out data.json
python3 -I $SK/build_excel.py data.json ECI_Monthly_Report_October_2026.xlsx
python3 <xlsx skill>/scripts/recalc.py ECI_Monthly_Report_October_2026.xlsx 90   # must report 0 errors
python3 -I $SK/build_dashboard.py data.json ECI_Report_Oct2026.html               # prints all key numbers
```
Check `prepare_data.py`'s line (leads, previous month, ad sets) and `build_dashboard.py`'s printout before writing
any narrative. Screenshot each tab with Playwright (`#overview #leads #conversion #channels #next`) and look at it.

## Step 3 - Narrative (config.narrative), then re-run build_dashboard.py
Keys: `headline`, `subhead`, `leads_intro`, `hours_note`, `channels_intro`, `adsets_note`, `lost_detail`,
`insights [{title, text}]` (4), `worked [html]` (4), `attention [html]` (4), `plan [{title, text, when}]` (5),
`next_intro`. Rules:
- Every number in the text must come from the printout (inquiries per ad set, cost per inquiry, lost-reason counts...).
  Re-read each sentence against it. Ad-set "Inquiries" = all conversations from that set, not only Intake.
- Plain, short sentences for a business owner. No Rework jargon. Missing facts stay TBU, never guessed.

---

## What the Excel contains
Summary (price inputs, 5-question KPIs vs previous month, organic block, data-quality note) · Data Sources (where each
number comes from + how to automate it) · Inquiries (by source / status / package / hour) · Campaign Cost (per ad set:
spend, impressions, reach, clicks, CTR, CPM, CPC, frequency, inquiries, cost per inquiry, paid bookings, Meta results)
· Lost Reasons · Organic Facebook (Manual) · Leads_Data (one row per inquiry, dropdowns for Status and Package).
All KPIs are formulas on Leads_Data, so the team updates statuses and the workbook recalculates.
Colours: green `E2EFDA` = Rework auto · yellow `FFF2CC` = Rework + team update / inputs / estimated ad set ·
orange `FCE4D6` = manual from Meta. Font Arial, headers navy `1F3864`.

Potential value = Σ package price per inquiry; an inquiry with no package uses the price weighted by the known
package mix (shown on Summary). Cost per inquiry = ad spend ÷ Meta-sourced inquiries.

## What stays manual (TBU)
- Meta-reported results per ad set (Ads Manager).
- Organic page insights: views, reach, interactions, link clicks, profile visits, follows, posts (Meta Business Suite).
- Status / package / lost reason for inquiries the sales team has not reviewed (`Status Reviewed = No`).

## Gotchas
- `since` is UTC epoch - always convert with the client offset (`--tz 8`) before bucketing by day/hour.
- Reach summed over ad sets double-counts people; label it "Summed per ad set".
- openpyxl: assign a fill only inside an `if` (StyleProxy is unhashable); recalc with LibreOffice before delivering.
- A `.xlsm` template without `vbaProject.bin` has no macros to lose - keep the macro content type or save as `.xlsx`.
- After changing the HTML, send it under a new file name; viewers cache the old one.
- Never commit `data.json`, exports or the generated reports - they hold names and phone numbers.

## Related setup (makes next month more automatic)
- Deal stage → lead stage automations per BU (`automation_editleadfieldsnode`, operation `edit_lead_stage`,
  fields `{"checks":{"stage":"1"},"data":{"stage":"<id>"},"fields":{"stage":{...}}}`, lead bound from
  `trigger.deal.fs_export.id`).
- Create a lead/deal for every inquiry with a Package field so the ad set, package and status live on the record.
