# Portfolio Trackers

Two Excel workbooks for running a contract and project portfolio, plus a one-page staff guide. Every alert, threshold, status list and colour is set on a controls tab, so the team can change the rules without editing formulas.

| File | What it is |
|---|---|
| [outputs/Contract_Tracker_2026_JHK3.xlsx](outputs/Contract_Tracker_2026_JHK3.xlsx) | **Contract Tracker.** Contract expirations, funds spent, next actions, and task orders / job order releases against master contracts. |
| [outputs/Project_Tracker_JHK3.xlsx](outputs/Project_Tracker_JHK3.xlsx) | **Project Tracker.** Blank template for portfolios, projects, tasks and issues. |
| [outputs/Contract_Tracker_Staff_Quick_Start_JHK3.docx](outputs/Contract_Tracker_Staff_Quick_Start_JHK3.docx) | One-page staff guide to the Contract Tracker: the four columns staff update, the weekly routine, and what the colours mean. |

---

## Contract Tracker

**Tabs (navy = contracts, teal = task orders, gold = controls)**

| Tab | Purpose |
|---|---|
| How To Use | Training guide, the reasoning behind every alert, sources, change log |
| Column Guide | Every Portfolio Tracker and Task Order Tracker column: who fills it in (original list, you, or a formula), what it means, how it is figured, an example, and which setting changes it |
| Dashboard | One page for the portfolio owner and leadership: health tiles, status mix, contracts ending by quarter, status by lane and owner, Top 10 most urgent |
| Summary | Counts by status, lane and owner; money totals |
| Portfolio Tracker | One row per contract (owned and shared), with Renewal Path, Action Status, Next Action, Contract Type and Owning Dept Code inputs, an Ownership column, and three calculated flags |
| DPS 365 Report | The monthly DPS report of every department contract ending within a year, linked to the Portfolio Tracker: on my tracker?, responsibility (my portfolio / visibility only), tracker status and next action, end-date differences, and a suggested monthly comment |
| Task Order How To | Guide to tracking task orders and job order releases |
| Task Order Dashboard | Ceiling, committed, available and paid; committed vs. available by master; releases by status, flag, month and type; Top 10 flagged releases |
| Task Order Tracker | One row per release, linked to its master contract by PO |
| Alert Controls | Every list, lead time, rule, threshold and colour (sections 1-19) |
| Original (as received) | The source contract list, unedited |
| Shared List (as received) | The shared-contracts list, unedited |
| DPS 365 (as received) | The DPS 365 report, unedited |
| Data Check | Every change made while loading the shared-contracts list (text dates and PO numbers converted, typos fixed, spellings made consistent, conflicts with the tracker) |

**How the alerts work**

- **Expiration Status.** Each contract gets a Start-By date: end date minus the lead time for its renewal path (extension 240 days, competitive rebid 540, sole source 365). Rules are checked in order and the first match wins: CLOSED, LAPSED, IN PROCESS, RED (≤180 days or Start-By passed), YELLOW (≤365 days or Start-By within 90), GREEN.
- **Funds Flag.** % expended, and spending compared with the share of the term already used (RED at 75% spent or 20+ points ahead of time; YELLOW at 50% or 10+ points).
- **Next Action Flag.** OVERDUE, DUE SOON (14 days), UPCOMING (30 days).
- **Priority.** The most urgent of the three flags (1 = most urgent).
- **Spending pace.** Burn Index = % spent / % of term used (amber 1.10, red 1.25), Projected Exhaustion Date, and "Funds Run Out Before End?" as the early warning for a vendor limit increase.
- **Supply coverage.** For lapsed or soon-to-lapse contracts: how the department keeps buying (Exhibit B, alternate PO, vendor honoring an extension, stock on hand), shown as Supply Status / Supply Risk. The Dashboard shows lapsed contracts by renewal movement and supply coverage.
- **Action Needed.** The whole row turns red (Act now) or amber (Plan) when a contract needs action; renewals in process with supply covered are not highlighted.
- **Task orders.** Ceiling = the master's Total Contract Value. Committed = releases whose status counts as committed. Available = ceiling − committed. The commitment flag goes YELLOW at 75%, RED at 90% and OVER CEILING at 100%. Each release gets one flag: EXCEEDS CEILING, NOT A MASTER, OUTLASTS MASTER, OVERPAID, OVERDUE or ENDING SOON.

- **Shared contracts.** Contracts owned by another department (by PO Dept Code) that the department draws on sit on the same list. Ownership reads Owned when the Owning Dept Code equals the department code on Alert Controls section 15, otherwise Shared. They get the same alerts, a "Owning dept renews" renewal path and their own Dashboard section.

All of these numbers are starting values that the team can edit on Alert Controls.

---

## Project Tracker

Blank template: Read Me, Column Guide, Dashboard, Project Tracker, Task Tracker, Issue Tracker, Priority Matrix, Guide, Lists. Task priority comes from Impact × Urgency; project health (Red / Amber / Green) comes from due dates, overdue tasks and open critical issues.

---

## Rebuilding the workbooks

The workbooks are generated by scripts, so they can be rebuilt exactly.

```bash
cd trackers/scripts

# Contract Tracker (from the source contract list and the shared-contracts list)
python build_contract_tracker_JHK3.py ../data/Contract_List_Source_JHK3.xlsx ../outputs/Contract_Tracker_2026_JHK3.xlsx \
    ../data/Shared_Contracts_BWS_Source_JHK3.xlsx ../data/DPS_365_Expiring_Report_JHK3.xlsx

# Keep a working copy's typed entries when rebuilding (REVIEWED_COL=1 if it has the Reviewed column in A):
REVIEWED_COL=1 CARRY_USER=<working copy>.xlsx python build_contract_tracker_JHK3.py ... 
python recalc.py <output>            # from the xlsx tooling
python insert_reviewed_column_JHK3.py <output> D2 Q2

# Project Tracker (blank)
python build_project_tracker_JHK3.py ../outputs/Project_Tracker_JHK3.xlsx

# Staff Quick Start (needs the npm package "docx")
node build_contract_quick_start_JHK3.js ../outputs/Contract_Tracker_Staff_Quick_Start_JHK3.docx
```

openpyxl writes formulas without cached values, so open and save each workbook in Excel (or recalculate it with LibreOffice headless) before reading it with pandas.

**Important:** the working copies that staff edit are the master records. Rebuilding from the scripts produces a fresh workbook with no team edits in it. To change a workbook that's in use, edit the working copy, or rebuild and then copy the team's inputs across.

| Script | Builds |
|---|---|
| `build_contract_tracker_JHK3.py` | Contract Tracker: How To Use, Summary, Portfolio Tracker, Alert Controls sections 1-8 |
| `contract_tracker_dashboard_part_JHK3.py` | Dashboard, Alert Controls section 9 |
| `contract_tracker_taskorder_part_JHK3.py` | Task Order Tracker, Contract Type column, Alert Controls sections 10-14 |
| `contract_tracker_todash_part_JHK3.py` | Task Order Dashboard, Task Order How To |
| `contract_tracker_testdata_part_JHK3.py` | Sample releases for testing (only with `TEST_TO=1`) |
| `contract_tracker_shared_part_JHK3.py` | Loads the shared-contracts list; Owning Dept Code / Ownership columns; Alert Controls section 15; Data Check and Shared List tabs |
| `contract_tracker_shared_report_part_JHK3.py` | Shared-contract sections on the Dashboard and Summary |
| `contract_tracker_supply_part_JHK3.py` | Burn Index, Projected Exhaustion Date, supply coverage, Action Needed row highlight, lapsed renewal x supply KPIs, Alert Controls sections 17-19 |
| `insert_reviewed_column_JHK3.py` | Adds the working copy's Reviewed column at A of the Portfolio Tracker with LibreOffice so every reference shifts as in Excel, and restores its freeze panes |
| `contract_tracker_dps365_part_JHK3.py` | DPS 365 Report tab, tracker columns On DPS 365 Report? / DPS 365 Comments / DPS Procurement Mgr, Alert Controls section 16, Dashboard DPS section |
| `contract_tracker_column_guide_part_JHK3.py` | Contract Tracker Column Guide |
| `build_project_tracker_JHK3.py` + `tracker_helpers_JHK3.py` | Project Tracker |
| `project_tracker_column_guide_part_JHK3.py` | Project Tracker Column Guide |
| `column_guide_JHK3.py` | Shared Column Guide writer used by both trackers |
| `build_contract_quick_start_JHK3.js` | Staff Quick Start |
