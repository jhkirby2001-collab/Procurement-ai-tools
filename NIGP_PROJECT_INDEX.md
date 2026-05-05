# NIGP Procurement Taxonomy & Classification Engine — Project Index

A clickable map of every file in this project. Click any filename below to jump directly to that file. Files are grouped by purpose, not by folder location.

**Project status:** COMPLETE (production run finalized 30 April 2026). Methodology version 1.1.

**One-line summary:** A reusable, NIGP-sourced commodity taxonomy and classification engine — validated on a 784,556-row 23-year public-sector procurement dataset, with 86.4% auto-classified by deterministic rule, 17.8% routed to a human review queue, end-to-end runtime 14 minutes 37 seconds.

---

## 1. Start here — leadership-facing deliverables

The four files a non-technical reader should look at first.

| File | What it is |
|---|---|
| [outputs/NIGP_Executive_Brief_JHK3.docx](outputs/NIGP_Executive_Brief_JHK3.docx) | Two-page Word brief for senior leadership. The shortest possible read. |
| [outputs/NIGP_Summary_for_Leadership_JHK3.xlsx](outputs/NIGP_Summary_for_Leadership_JHK3.xlsx) | Single-tab Excel summary — 17 categories with row counts, headline metric tiles, methodology footnotes. |
| [outputs/NIGP_Methodology_for_Leadership_JHK3.docx](outputs/NIGP_Methodology_for_Leadership_JHK3.docx) | The full methodology re-rendered for leadership: subheadings, bullet points, callout boxes. |
| [outputs/HOW_TO_USE_NIGP_Mapping_JHK3.md](outputs/HOW_TO_USE_NIGP_Mapping_JHK3.md) | Plain-language guide for procurement staff using the classified output in Excel. |

---

## 2. Full methodology and project objective

| File | What it is |
|---|---|
| [spend-analysis/METHODOLOGY_JHK3.md](spend-analysis/METHODOLOGY_JHK3.md) | The long-form methodology document — 13 sections plus locked-decisions appendix. The canonical narrative. |
| [spend-analysis/PROJECT_OBJECTIVE.md](spend-analysis/PROJECT_OBJECTIVE.md) | The original project objective statement that framed the work. |

---

## 3. The classified data — the actual deliverable

These two files are produced by running the classifier. They are NOT in git (too large), but they live on disk in this codespace.

| File | What it is | Status |
|---|---|---|
| `outputs/NIGP_Mapping_JHK3.csv` | 784,556 rows, every record with its three-level classification, confidence score, and reason. **225 MB.** | On disk — regenerable |
| `outputs/NIGP_Mapping_Review_Queue_JHK3.csv` | 139,868 review-flagged rows for procurement-staff triage. **30 MB.** | On disk — regenerable |

If either is ever lost, regenerate by running:
```
python spend-analysis/scripts/classifier_JHK3.py --batch
```

---

## 4. Rule files — what procurement staff edit to maintain the classifier

These CSVs are the source-of-truth for every classification decision. Edit them directly to add, retire, or change rules. No code changes required.

| File | What it is |
|---|---|
| [spend-analysis/data/reference/keyword_rules_DRAFT_JHK3.csv](spend-analysis/data/reference/keyword_rules_DRAFT_JHK3.csv) | 148 hand-curated keyword rules. The highest-volume hitters. |
| [spend-analysis/data/reference/keyword_rules_from_ai_JHK3.csv](spend-analysis/data/reference/keyword_rules_from_ai_JHK3.csv) | 6,766 AI-mined long-tail rules with full provenance metadata in each row. Frozen — not consulted at runtime. |
| [spend-analysis/data/reference/account_patterns_DRAFT_JHK3.csv](spend-analysis/data/reference/account_patterns_DRAFT_JHK3.csv) | Six FMPS account-code patterns (the 220xxx subgrant series). |
| [spend-analysis/data/reference/business_categories_JHK3.csv](spend-analysis/data/reference/business_categories_JHK3.csv) | Mapping from NIGP 3-digit Class to one of the 17 Business Categories — with judgment notes. |
| [spend-analysis/data/reference/business_categories_summary_JHK3.csv](spend-analysis/data/reference/business_categories_summary_JHK3.csv) | 17-row summary of the Business Category structure with definitions. |

---

## 5. NIGP reference catalog

The NIGP commodity-code lookup tables. These rarely change.

| File | What it is |
|---|---|
| [spend-analysis/data/reference/nigp_codes_3digit_JHK3.csv](spend-analysis/data/reference/nigp_codes_3digit_JHK3.csv) | 138 NIGP 3-digit Classes used by the classifier. |
| [spend-analysis/data/reference/nigp_codes_5digit_JHK3.csv](spend-analysis/data/reference/nigp_codes_5digit_JHK3.csv) | 470 NIGP 5-digit Class-Items used by the classifier. |
| [spend-analysis/data/reference/nigp_codes_10digit_JHK3.csv](spend-analysis/data/reference/nigp_codes_10digit_JHK3.csv) | 4,592 codes — full EY-supplied catalog at 10-digit granularity. |

---

## 6. Production classifier — runs every batch

| File | What it is |
|---|---|
| [spend-analysis/scripts/classifier_JHK3.py](spend-analysis/scripts/classifier_JHK3.py) | The production classifier. Rules-only at runtime — no API key needed. Supports two modes: `--batch` (whole file) and `--describe "text"` (single record). |

**To run a full batch:** `python spend-analysis/scripts/classifier_JHK3.py --batch`
**To classify one description:** `python spend-analysis/scripts/classifier_JHK3.py --describe "your description here"`

---

## 7. Build scripts — one-time setup, do not re-run unless rebuilding the project

| File | What it is |
|---|---|
| [spend-analysis/scripts/build_nigp_reference_JHK3.py](spend-analysis/scripts/build_nigp_reference_JHK3.py) | Built the three NIGP reference CSVs from the raw EY data. |
| [spend-analysis/scripts/build_business_categories_draft_JHK3.py](spend-analysis/scripts/build_business_categories_draft_JHK3.py) | Built the initial draft of the 17-category mapping. |
| [spend-analysis/scripts/build_keyword_rules_draft_JHK3.py](spend-analysis/scripts/build_keyword_rules_draft_JHK3.py) | Built the initial draft of the hand-curated keyword rules. |
| [spend-analysis/scripts/build_account_patterns_draft_JHK3.py](spend-analysis/scripts/build_account_patterns_draft_JHK3.py) | Built the initial draft of the account-code patterns. |
| [spend-analysis/scripts/profile_descriptions_JHK3.py](spend-analysis/scripts/profile_descriptions_JHK3.py) | One-time descriptive profiling of the four PO/invoice description fields. |
| [spend-analysis/scripts/build_leadership_deliverables_JHK3.py](spend-analysis/scripts/build_leadership_deliverables_JHK3.py) | Generates the three leadership documents in Section 1 (Excel + 2 Word docs). Re-run any time to regenerate them. |

---

## 8. AI mining — one-time bounded use, already completed

These two scripts ran exactly ONCE during the project build. Production classification is rules-only — these scripts are not part of the runtime path.

| File | What it is |
|---|---|
| [spend-analysis/scripts/ai_classify_JHK3.py](spend-analysis/scripts/ai_classify_JHK3.py) | The single AI batch run over 30,342 unique long-tail descriptions. **Do not re-run** (cost ~$27, output already preserved). |
| [spend-analysis/scripts/harvest_ai_patterns_JHK3.py](spend-analysis/scripts/harvest_ai_patterns_JHK3.py) | Promoted the AI proposals into the rule file using the high/medium/low confidence rules. |

---

## 9. Diagnostic outputs — useful but not deliverables

These files explain *how* the classifier made its decisions. They are NOT in git (too varied / regenerable) but live on disk in this codespace.

| File | What it is |
|---|---|
| `spend-analysis/data/processed/classifier_coverage_report_JHK3.csv` | Which rules fired how many times — used to identify rule gaps. |
| `spend-analysis/data/processed/classifier_account_pattern_report_JHK3.csv` | How often each account-code pattern fired. |
| `spend-analysis/data/processed/ai_classified_unique_descriptions_JHK3.csv` | All 30,342 AI proposals with confidence and reasoning — preserved for audit. |
| `spend-analysis/data/processed/description_column_profile_JHK3.csv` | One-time descriptive profile of the four description fields. |
| `spend-analysis/data/processed/top_description_patterns_JHK3.csv` | Most common description patterns from the EY raw data. |
| `spend-analysis/data/processed/lowinfo_description_examples_JHK3.csv` | Examples of "Misc" / "Per contract" type descriptions. |

---

## 10. Source data — never modified

| File | What it is |
|---|---|
| `spend-analysis/data/raw/ey raw data.xlsx` | Original EY consulting deliverable. **326 MB. Untouched.** Not in git (too large; raw data is never committed). |

---

## What is NOT part of this project

For clarity — these files exist in this repo but are unrelated to the NIGP project:

- The earlier EY benchmark / Top Consolidation Savings work (separate analytical effort)
- `outputs/research-reports/` market research deliverables
- `market_research_emergency_generator_maintenance_2026-04-22.md` (separate market-research engagement)
- The procurement-researcher agent definition under `.claude/agents/`

---

*Index maintained by James H. Kirby III, CSCP, MS-SCM. Last updated 1 May 2026.*
