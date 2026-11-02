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

- **Column Guide tabs must stay in step with the trackers.** When a tracker column is added, moved, renamed, or changes between input and formula, update `contract_tracker_column_guide_part_JHK3.py` / `project_tracker_column_guide_part_JHK3.py` and re-run the header check (guide name and type vs. the tracker header and its "(auto)" marker).
- Examples live on the Column Guide, never as sample rows on a tracker (sample rows would be counted on the dashboards). James agreed to this on 2026-11-01.

## Layout facts (Contract Tracker)

- Portfolio Tracker columns: A-Q from the source list; R-X inputs (Renewal Path, Action Status, Next Action, Due Date, Done?, Total Contract Value, Expended); Y-AL calculated; AM-AQ notes; AR Contract Type; AS Owning Dept Code; AT Ownership (auto); AU Spec Number; AV PO Contract Type; AW BWS File Notes.
- Rows: 2-87 original list (86 contracts), 88-140 shared-contracts list (53 contracts). The build takes the shared list as an optional third argument.
- Alert Controls: 1 general, 2 renewal paths, 3 action statuses / categories (col K), 4 expiration rules, 5 funds rules, 6 next action rules, 7 colour key, 8 summary lanes / owners, 9 dashboard settings, 10 contract types, 11 release types, 12 release statuses, 13 commitment rules, 14 release flag rules (conditions are a fixed list in col K171+), 15 shared contracts (my department code D184 = 88; owning-dept code list C187-C198).
- Data rows: tracker 2-300, Task Order Tracker 5-504.

## Open items

- Charts use the starting status lists and colours; adding a rule or lane updates the tables, but the chart data range has to be widened by hand.
- Release-flag conditions are a fixed list; a new kind of condition needs a formula change in `contract_tracker_taskorder_part_JHK3.py` (`rf_cond`).
- James planned to clean the contract data, hide helper columns and protect Alert Controls himself. He is the only person updating the tracker for now (staff not yet using it).
- Shared contracts (2026-11-02): James's rules for the load were to keep both PO 87372 rows (may be a separate service contract; he will review), fix text dates / year typos / text POs, keep the tracker's dates where the shared file differs, make contract-type spellings consistent, and carry the shared file's notes into BWS File Notes. Codes only for now; department names can come later. PO 28456 and 271670 are coded 85 in the shared file and PO 27036 carries Dept Code 38, so all three show as Shared until he confirms.
- The four text start dates on the original list (11/22/0024, 07/11/0007, 02/08/0024, 05/05/0023) were fixed to 2024 / 2007 / 2024 / 2023 on James's request (2026-11-02) and logged on Data Check. James confirmed PO 14285 (AMR) started in 2007.
