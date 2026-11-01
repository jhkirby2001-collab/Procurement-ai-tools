# Portfolio Trackers — Project Guide

Contract Tracker, Project Tracker and Staff Quick Start for James's contracts administration work. Built 2026-10-30 to 2026-11-01. See `README.md` for what each file is and how to rebuild it.

## Working rules

- **The working copies are the master records.** James keeps them in his OneDrive / Google Drive "Trackers" folder and staff edit them. When he asks for a change, ask for his latest copy first. Rebuilding from the scripts gives a fresh workbook without the team's edits; only rebuild when he wants the template changed, and say so.
- He starts change requests with "Contract Tracker update:" or "Project Tracker update:". The two are separate workbooks; don't mix them.
- **Everything configurable lives on the controls tab** (Contract Tracker: Alert Controls; Project Tracker: Lists / Priority Matrix). New statuses, thresholds and rules go there, never hard-coded into formulas. He explicitly wants to change things after meeting with his staff.
- Rules tables are evaluated **first match wins** in slot order. Colours are chosen by name from the palette (section 7); conditional formatting maps the names to fills.
- The contract data is public. Don't add privacy or confidentiality warnings about storing it.
- `_JHK3` suffix on every file. Keep public-facing markdown free of agency branding (repo-wide rule); the binary workbooks may name real sources.
- Avoid functions LibreOffice can't evaluate (XLOOKUP, FILTER, SORT, UNIQUE). Lists on dashboards use SMALL / MATCH on helper keys in the hidden `Dashboard Data` sheet.
- After any build: recalculate, confirm zero formula errors, and render the dashboards (LibreOffice → PDF → PNG) to check the layout. `TEST_TO=1` loads sample releases; `RENDER_DASH=<sheet>` hides every other sheet.

## Layout facts (Contract Tracker)

- Portfolio Tracker columns: A-Q from the source list; R-X inputs (Renewal Path, Action Status, Next Action, Due Date, Done?, Total Contract Value, Expended); Y-AL calculated; AM-AQ notes; AR Contract Type.
- Alert Controls: 1 general, 2 renewal paths, 3 action statuses / categories (col K), 4 expiration rules, 5 funds rules, 6 next action rules, 7 colour key, 8 summary lanes / owners, 9 dashboard settings, 10 contract types, 11 release types, 12 release statuses, 13 commitment rules, 14 release flag rules (conditions are a fixed list in col K171+).
- Data rows: tracker 2-300, Task Order Tracker 5-504.

## Open items

- Charts use the starting status lists and colours; adding a rule or lane updates the tables, but the chart data range has to be widened by hand.
- Release-flag conditions are a fixed list; a new kind of condition needs a formula change in `contract_tracker_taskorder_part_JHK3.py` (`rf_cond`).
- James planned to clean the contract data, hide helper columns and protect Alert Controls himself.
