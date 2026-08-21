# Maverick Spend Conversion Tracker

*Exhibit B → awarded contract conversion. Headline metric: **Contract Conversion Rate (CCR)**.*

**Status:** complete and delivered (2026-08-21). Paused, safe to resume.
**Branch:** `claude/exhibit-b-consolidation-nqvtak`
**Owner:** James H. Kirby III, CSCP, MS-SCM

Answers one question: **did non-contract ("maverick") Exhibit B spend actually move
onto an awarded contract this year — and can we prove it?**

## Naming (settled 2026-08-21)

| | Name | Used where |
|---|---|---|
| **Project** | Maverick Spend Conversion Tracker | Repo, docs, workbook titles — public-facing and vendor-neutral |
| **Metric** | **Contract Conversion Rate (CCR)** | Leadership decks, quarterly reporting — the number that gets quoted |

"Exhibit B" is internal jargon and this repo is public, so it stays in subtitles and in the
workbook body where it correctly labels the *source data* — never as the project name.
The quotable sentence is: *"CCR for 2026 year to date is 0.33% — $175K of $53.5M."*

---

## The answer, as delivered

| | |
|---|---|
| Exhibit B items verifiably converted | **8** |
| Requisitions behind them | **22** (20 carry dollars, 2 are $0) |
| Maverick spend moved onto contract | **$175,233.91** |
| Distinct contracts landed on | **6** |
| Award methods | 6 BID · 1 EMERGENCY · 1 SOLE SOURCE |
| **Total Exhibit B spend, Jan 5 – Aug 20 2026** | **$53,461,392.14** (730 reqs) |
| **Contract Conversion Rate (CCR)** | **0.33%** |
| **Still bought outside a contract** | **$53,286,158.23 — 99.67%** |
| Ratio | $305 of Exhibit B spend for every $1 moved onto a contract |

**Framing decision (locked):** every exit from Exhibit B is a **win**, regardless of award
method. A sole-source or emergency award is still an awarded contract — justified,
documented, negotiated, auditable. Maverick spend is the *absence* of a contract vehicle.
Award method is reported as context, never as a pass/fail gate. Do not reintroduce a
"competitive vs non-competitive" scoring gate.

---

## Deliverables

| File | What it is |
|---|---|
| `outputs/Exhibit_B_Verified_Conversions_JHK3.xlsx` | **The one that went to leadership.** 3 tabs: Summary · Verified Conversions (proof table + whole-year context block) · Backing Data. |
| `outputs/Exhibit_B_Conversion_Analysis_JHK3.xlsx` | Fuller 11-tab analysis. Superseded for leadership use, kept for depth: pipeline forecast, claimed-vs-verified, reverse leakage, data-quality findings. |

## Scripts

| Script | Role |
|---|---|
| `scripts/exhibit_b_conversion_analysis_JHK3.py` | Match engine + the 11-tab workbook. **All matching logic lives here.** |
| `scripts/exhibit_b_verified_conversions_JHK3.py` | Imports the engine, builds the 3-tab leadership workbook. No duplicated logic. |

```bash
cd spend-analysis/scripts
python exhibit_b_verified_conversions_JHK3.py            # leadership workbook
python exhibit_b_conversion_analysis_JHK3.py             # full 11-tab analysis
python exhibit_b_verified_conversions_JHK3.py <src.xlsx> # explicit source
```

---

## ⚠ Source data is NOT in the repo

`spend-analysis/data/raw/` is gitignored ("organization's source extracts, never
committed") and **this repo is public**, so the source workbook was deliberately not
committed. It came in as a session upload and is gone once that container is reclaimed.

**To re-run, supply it again** — a two-tab workbook:
1. Exhibit B report for the period
2. Awarded contracts for the same period

Then either drop it in `spend-analysis/data/raw/` (the scripts glob for
`ExhibitB_Awarded_Contracts*.xlsx`), set `EXHIBIT_B_SRC=/path/to/file.xlsx`, or pass it as
argv[1]. Missing source produces a clear message, not a traceback.

Original file: `ExhibitB_Awarded_Contracts_1.1.2026__8.21.2026.xlsx` — Exhibit B tab
730 rows × 20 cols, Awarded tab 100 rows × 22 cols.

---

## Method (why the numbers are defensible)

**Unit of analysis is the sourcing effort, not the requisition.** 730 requisitions collapse
to 175 distinct sourcing efforts, keyed on the pending new-contract requisition (RX). One
effort absorbs 49 requisitions / $16.3M. Counting requisitions distorts everything.

**The RX field is free text and must be parsed.** Real values include `RX 596048`,
`RX676831- SPEC 1197406C`, concatenated `710856710875`, and
`Target Market 710856, Non-Target Market 710875`. Recovery gets 614 of 730 rows.

**Five evidence tiers; only A and B count.**

| Tier | Basis | Counts? |
|---|---|---|
| A | Requisition ID link (EB new-contract RX = awarded REQ), or reported contract# = awarded PO/CONTRACT_ID | ✅ |
| B | Same vendor + commodity token overlap ≥ 0.10 | ✅ |
| C | Same vendor only | ❌ disclosed, never in the headline |
| D | Description similarity ≥ 0.30, no vendor/ID link | ❌ |
| E | No linkage found | ❌ |

**Timing gate:** a Tier A/B match counts only if the award is dated *after* the Exhibit B
request. An earlier award is *reverse leakage* (spend that should have ridden an existing
contract), not a conversion.

**Self-reported fields are evidence of nothing.** `Contract Awarded?` and
`Contract Number Awarded` are department self-reports, treated as a claim to be tested.

---

## Findings worth keeping

1. **Tracking is broken both directions.** 48 requisitions self-reported as awarded; only
   6 verify (12.5%). Of 25 contract numbers departments supplied, **only 2 appear anywhere
   in the award file**. Separately, **16 genuine conversions were never flagged as awarded**.
   Not inflation — nobody is maintaining Exhibit B status.
2. **One contract is 30% of the exception report.** Comprehensive Custodial Services
   (Zones 1/2/4): 49 requisitions, $16.3M, still in Specification Development.
3. **$8.6M is one step from award** — 25 requisitions in Recommendation of Award or
   Signature Cycle. The honest rebuttal if 0.33% is called failure: most of it is timing.
4. **$23.0M won't clear this year** — still in Specification Development, 12+ months out.
5. **Three departments converged on one contract.** CPD, 2FM/DCASE and DPS each bought
   portable toilet service off-contract independently; all landed on Service Sanitation
   381051. A small consolidation story.
6. **`Contract Number Awarded` is unusable as a field** — 51× "TBD", 40× "na", 18× "0",
   17× "NO", plus "yes", "Pending", "UNKNOWN", "N?A".

---

## Known limits (state these before anyone quotes a number)

- **The floor, not the ceiling.** The two systems share no reliable common key; linkage
  depends on requisition numbers appearing inside a free-text field. Genuine conversions
  that left no trace are not counted.
- **"Not verified" ≠ "no award happened."** The award file holds 100 contracts. If that is
  not the complete population, an effort could have landed on a contract absent from it.
  Stress-tested: the top 10 unconverted efforts have no plausible counterpart (best fuzzy
  candidates 0.00–0.17), so the low rate is real, not a matching failure.
- **Recent months understate by construction** — an Exhibit B filed in August has had no
  time to reach award.
- **SPEC-number matching is a dead end** — 0 hits. Do not rebuild it.

---

## If picking this back up

Verification already done — every figure reconciled to raw: all 22 requisitions present
with matching amounts, none denied, every award dated after its Exhibit B, all 6
requisition-ID claims confirmed in both systems, all summary cells tie. Clean recalc.

Natural expansions:
1. **Quarterly re-run** as a standing KPI — scripts take a new source as argv[1]; CCR
   becomes a trend line rather than a single reading.
2. **Close the loop on the tracking failure** — carry the Exhibit B requisition number
   forward onto the award record so conversion becomes measurable automatically. This is
   the highest-value fix and it is a process change, not an analytics one.
3. **Pipeline aging** — how long does an effort sit in each Procurement Phase? Turns the
   forecast into a cycle-time diagnostic.
4. **Consolidation overlay** — reuse `spend_report_JHK3.py`'s savings engine against
   still-open efforts to size the prize of clearing the backlog.

**Environment note:** `libreoffice-calc` is not installed in fresh containers, so
`recalc.py` silently times out until you run `apt-get update && apt-get install -y
libreoffice-calc`. Do that before trusting any formula workbook.
