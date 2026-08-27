# NIGP-Sourced Procurement Category Mapper — Project Guide

This repo is the home of the **NIGP-Sourced Procurement Category Mapper**, an independent public-sector procurement classification tool. The build is complete and in production; ongoing work is incremental polish and Phase 2 (interactive drill-down).

**Owner:** James H. Kirby III, CSCP, MS-SCM (jhkirby2001@yahoo.com)
**Output naming convention:** `_JHK3` suffix on every deliverable file
**Status:** Production. Latest commit `e349d65` on `origin/main`. Active branch work `claude/github-identification-ivqudt` (2026-08-07: fence fix + interactive-robustness rule expansion 246→280 + expanded operate-and-trust guide; 2026-08-18: Spend Report path — every row now lands in a real category, the "Unclassified" bucket replaced by a single catch-all **"General & Other Procurement"**, a new **Examples** column shows representative descriptions per category, **Grants** scoped to fire only on genuine grant signals on uploads (no memorized-exact misfires), curated rules 293→336; 2026-08-20: **prescriptive savings engine** — addressable-spend + benchmark-rate cost-avoidance with **hard (cashable) + cost-avoidance split**, ranked recommendations, **Savings Opportunity Summary** + **Top Consolidation Opportunities** + **Sources & Methodology** tabs, **professional matplotlib charts**, and **General & Other absorbed via best-fit** into the 17 categories; 2026-08-21: **plain-language pass** — leadership-friendly wording everywhere ("spend we can combine", "cash savings", "avoided costs", "what to combine"), correct procurement definitions kept in a "What these numbers mean" note + Sources tab (addressable spend = influenceable, *not* already-optimized; that's spend-under-management), savings measured bottom-up from the data with no assumed addressable-%, and **live savings-rate sliders**; 2026-08-24: **reader-chosen report scope** — a dropdown picker replaces the fire-everything workbook. Executive Summary + Sources & Methodology are always on (`ALWAYS_ON`); a lean set is pre-ticked (`DEFAULT_ANALYSES` = savings, category, Pareto, top vendors); everything else is opt-in, and infeasible analyses are never offered. **Selection controls presentation only — every analysis is still computed on every run**, so the savings headline cannot move with a checkbox; changing the picker re-renders and rebuilds the Excel instantly with no re-classification. **Every selected analysis now carries a chart** — seven new builders (`consolidation_png`, `item_consolidation_png`, `nigp_consolidation_png`, `opportunities_png`, `tail_png`, `concentration_png`, `single_multi_png`, `matrix_png`) plus generic `split_bar_png` / `stacked_barh_png`, shared by app and Excel so the two cannot drift. Contents tab appears only past `CONTENTS_THRESHOLD` tabs; 2026-08-25: **Top Consolidation Opportunities names names** — counts alone told leadership a commodity was fragmented 12 ways but not who, so the tab gains `COL_VENDOR_NAMES` / `COL_DEPT_NAMES` (spend-ordered, capped at `OPP_NAME_CAP`=15 with a "+N more" tail) beside the existing counts, plus `COL_OPP_VARIANTS` showing the wordings that rolled together — which makes a "Matched by: NIGP 3-digit class" row auditable rather than trusted. Excel gets real sort/filter dropdowns (`auto_filter`, stopping short of the TOTAL row) and the item column frozen; `_style_data_sheet` gained `wrap_widths`/`wrap_row_height` so a sheet with five text columns is navigable. Also fixed: Streamlit markdown was parsing paired `$` in the exec summary as LaTeX, rendering "$142,372 in savings ($56,949 cash)" as italic maths — `_md_money_safe` escapes them). Repo is PUBLIC.

---

## What this project does

Classifies 784,556 historical public-sector procurement rows into a three-level taxonomy on every row:

**Business Category (17) → NIGP Class (3-digit) → NIGP Item (5-digit)**

Inputs to the classifier: description text + Chicago FMPS account/object/fund codes ONLY. Vendor names and any prior-supplied NIGP codes are NOT inputs.

**Production batch metrics (post-resolver, 2026-05-14; batch not re-run since):** 100% mapped, 0 rows in review queue. Coverage breakdown (rule base was 246 curated at run time; now 280 — new rules fold in at next batch refresh):
- Tier 1 — Keyword rule (246 curated + 6,766 AI-mined at run time): 688,044 (87.7%)
- Tier 2 — Chicago FMPS account-code pattern: 2,028 (0.3%)
- Tier 3 — AI-assist resolver (saved one-time AI output, no new API call): 93,518 (11.9%)
- Tier 4 — Unclassified — No Description (no usable text in any of 4 fields): 966 (0.1%)
- Review_Flag=Yes (terminal review queue): 0 (0.0%)

---

## Key paths

**Rule files (procurement-staff editable, version-controlled):**
- `spend-analysis/data/reference/keyword_rules_DRAFT_JHK3.csv` — 336 hand-curated rules (edit these freely)
- `spend-analysis/data/reference/keyword_rules_from_ai_JHK3.csv` — 6,766 AI-mined rules (frozen — do not regenerate)
- `spend-analysis/data/reference/account_patterns_DRAFT_JHK3.csv` — 6 subgrant account patterns
- `spend-analysis/data/reference/business_categories_JHK3.csv` — canonical 138-row NIGP-class → Business-Category map
- `spend-analysis/data/reference/business_categories_summary_JHK3.csv` — 17-row summary
- `spend-analysis/data/reference/nigp_codes_{3,5,10}digit_JHK3.csv` — NIGP code reference tables

**Production scripts:**
- `spend-analysis/scripts/classifier_JHK3.py` — production classifier, dual-mode (batch + single-record), runs Tier 1 + Tier 2
- `spend-analysis/scripts/resolve_review_queue_JHK3.py` — Tier 3 AI-assist resolver (reads saved AI output, no new API call)
- `spend-analysis/scripts/spend_report_JHK3.py` — reusable, **adaptive**, now **prescriptive** spend-report engine: `profile_columns` + `plan_analyses` + vectorized `classify_series` + `compute_all` orchestrator. **Savings / cost-avoidance engine** (`consolidation_opportunities` + `savings_summary`): addressable-spend methodology, benchmark rates (`DEFAULT_SAVINGS_RATES` = +3%/vendor, +2%/dept, 15% cap, 5% tail, 40% hard-fraction), splits identified savings into **hard (cashable) + cost avoidance**, annualized + 3-year projection, per-opportunity **recommended actions**. Analyzers: spend-by-category, Pareto, top vendors, consolidation, item-level (`item_consolidation`) + same-commodity-by-NIGP-code (`nigp_item_consolidation`), spend-by-department, spend-trend, tail-spend, vendor-concentration (HHI), single-vs-multi-source, category×department matrix, spend-by-any-dimension. **Professional matplotlib charts** (`barh_png`/`pareto_png`/`line_png`/`savings_split_png`/`split_bar_png`/`stacked_barh_png` plus one builder per analysis → PNG bytes, embedded in both Excel and the app; `CHART_SERIES` is a validated 8-hue categorical palette — brand red stays reserved for alerts, and `text.parse_math` is off so `$`-pairs in labels aren't eaten as mathtext). **Report scope is reader-chosen**: `plan_analyses` keys feed a picker, `build_excel_report(..., selected=)` gates the tabs (`selected=None` → everything, which is what the CLI passes), `ALWAYS_ON`/`DEFAULT_ANALYSES`/`default_selection` define the contract. Executive summary leads with the **savings headline + ranked recommendations**. Excel adds a **Savings Opportunity Summary** tab, **Top Consolidation Opportunities** tab, and a **Sources & Methodology** tab (`SOURCES_SECTIONS`/`SOURCES_REFERENCES` — real, cited references). **General & Other is absorbed via best-fit** (`absorb_catchall`/`assign_best_fit`) so every row lands in one of the 17 categories — no visible catch-all. On uploads, **Grants (`GRANTS_CATEGORY`) fires only on strong grant signals**; memorized `exact` rules are skipped. Requires `matplotlib` (in requirements.txt).
- `spend-analysis/scripts/build_leadership_deliverables_JHK3.py` — regenerates Word/Excel summaries
- `spend-analysis/scripts/build_sop_JHK3.py` — regenerates the SOP .docx
- `spend-analysis/scripts/build_spend_report_howto_docx_JHK3.py` — regenerates the Spend Report how-to Word doc (`outputs/HOW_TO_Spend_Report_JHK3.docx`)
- `spend-analysis/scripts/audit_classifier_coverage_JHK3.py` — 63-phrase plain-English regression test
- `spend-analysis/scripts/fix_ai_rule_category_mismatches_JHK3.py` — idempotent integrity fix
- `spend-analysis/scripts/ai_topup_uncovered_JHK3.py` — targeted AI top-up (held in reserve, requires API key — not part of standard runtime)
- `spend-analysis/scripts/ai_classify_JHK3.py` — AI mining (BUILD-TIME ONLY — do not re-run, see below)

**Deliverables:**
- `outputs/NIGP_Mapping_JHK3.csv` — full 784,556-row classified file, post-resolver (~225 MB, GITIGNORED)
- `outputs/NIGP_Mapping_Review_Queue_JHK3.csv` — empty since 2026-05-14 resolver pass (0 rows); retained for schema-forward-compat (GITIGNORED)
- `outputs/NIGP_Summary_for_Leadership_JHK3.xlsx` — single-tab Excel summary
- `outputs/NIGP_Executive_Brief_JHK3.docx` — 2-page Word brief
- `outputs/NIGP_Methodology_for_Leadership_JHK3.docx` — full methodology, leadership-formatted
- `outputs/NIGP_SOP_JHK3.docx` — Standard Operating Procedure (Operations / Data Processing / Taxonomy Governance)
- `outputs/HOW_TO_USE_NIGP_Mapping_JHK3.md` — staff-facing usage guide
- `outputs/HOW_TO_Spend_Report_JHK3.docx` — leadership-formatted Spend Report how-to & methodology (Word)
- `outputs/NIGP_Web_App_Factsheet_JHK3.md` — short Layer-1 factual brief about the web app

**Documentation:**
- `spend-analysis/METHODOLOGY_JHK3.md` v1.2 — canonical methodology narrative (read before debating rationale or AI defensibility)
- `NIGP_PROJECT_INDEX.md` — repo index
- `README.md` — public-facing readme

**Web app:**
- `streamlit_app.py` (repo root) — eight-page Streamlit app: Classify → Bulk Classify → Spend Report → Spend Report Methodology → Procurement Taxonomy Logic → Methodology → Business Categories → Rule Lookup
- Deployed at `https://chicago-nigp-classifier.streamlit.app/` (subdomain rename to `nigp-sourced-category-mapper` may be in flight — verify before citing)
- Password: `chicagosbo2026` (in gitignored `.streamlit/secrets.toml`)

---

## Common commands

```bash
# Re-run full batch classification (~15 min)
cd spend-analysis && python scripts/classifier_JHK3.py --batch

# Classify a single description from terminal
python scripts/classifier_JHK3.py --describe "elevator maintenance"

# Regenerate leadership Word/Excel deliverables
python scripts/build_leadership_deliverables_JHK3.py

# Run the 63-phrase plain-English audit
python scripts/audit_classifier_coverage_JHK3.py

# Local Streamlit preview
streamlit run streamlit_app.py
```

---

## Locked architectural decisions — DO NOT relitigate

| # | Decision |
|---|---|
| 1 | Three-level taxonomy on every row: Business Category (17) → NIGP Class (3-digit) → NIGP Item (5-digit). |
| 2 | AI used ONCE during build to mine long-tail patterns. Production runtime is rules-only. No API key needed to classify. |
| 3 | Classifier inputs: description text + Chicago FMPS account/object/fund codes ONLY. Vendor and prior-supplied NIGP codes are NOT inputs. |
| 4 | Lean ~16-column output. Raw 87-column file preserved separately. |
| 5 | Dual-mode classifier: same core function for batch and single-record CLI. |
| 6 | Rule files externalized as CSV so procurement staff can edit without touching Python. |
| 7 | AI promotion thresholds: high → auto-promote; medium → promote only if row_count ≥ 5; low → never promote. |
| 8 | All 784,556 rows classified regardless of date. Time-agnostic by design. |

**Date-coverage phrasing (locked for all public summaries):** "AP activity years 2017, 2020, 2021, 2023." `METHODOLOGY_JHK3.md` lines 49-51 preserve precise PO/AP/Payment Date min-max — that's intentional, do not "reconcile" it.

**Subgrant resolution (empirical finding, locked):** 220xxx FMPS accounts (220005, 220044, 220100, 220801, 220300, 220999) are 93-100% subgrant disbursements, NOT commodity purchases. ~218K rows classify to the 17th Business Category "Grants & Pass-Through Funding" (no NIGP class). Routes: program-tag prefix keyword rules (primary), 6-account `account_patterns` (secondary).

---

## Don't-do list (hard rules)

1. **Do not re-run `ai_classify_JHK3.py`.** Cost ~$27. Output already preserved at `data/processed/ai_classified_unique_descriptions_JHK3.csv`. The 6,766 AI-mined rules in `keyword_rules_from_ai_JHK3.csv` are frozen.
2. **Do not use spend dollars to design or weight Business Categories.** Categories are MECE across *what* the agency buys, not *how much* it spends. (See `feedback_taxonomy_vs_spend.md` in user memory.)
3. **Do not reintroduce "City of Chicago," "DPS," or "Department of Procurement Services" branding into public-facing surfaces** (`streamlit_app.py`, `README.md`, `NIGP_PROJECT_INDEX.md`, `outputs/HOW_TO_USE_*`). Tool must read as independent. Internal scripts and methodology docs intentionally retain factual data-source references — that's not branding.
4. **Do not amend the `_JHK3` suffix convention.** Every deliverable, every script that produces deliverables, every rule file uses it.
5. **Do not commit `outputs/NIGP_Mapping_JHK3.csv` or `outputs/NIGP_Mapping_Review_Queue_JHK3.csv`.** They're gitignored on purpose (~225 MB and ~40 MB).
6. **Do not commit `.streamlit/secrets.toml`.** Gitignored. Contains the password.
7. **Do not change the password (`chicagosbo2026`)** without an explicit rotation request. Rotation declined and locked.
8. **Do not write spend-analysis content here.** That sub-project has its own `spend-analysis/CLAUDE.md`. This file is NIGP-only.

---

## Brand palette (locked for any leadership artifact)

- Chicago Navy `#002F6C` — h1, table headers, metric tiles
- Chicago Light Blue `#41B6E6` — h2, horizontal rules
- Chicago Red `#DA291C` — TOTAL row, callout-box labels
- Light blue tint `#D6EEF9` — table row banding, callout fills
- Light red `#E57373` — Level 3 provenance line emphasis (taxonomy visual)

---

## Open items (resume from here)

1. **Phase 2 — interactive three-level drill-down.** The original feature ask. (Note: "level" = taxonomy hierarchy Level 1 → 2 → 3; distinct from the 4-tier classification pipeline described in `METHODOLOGY_JHK3.md` §5.) Blocked on data-architecture decision because the 225 MB classified CSV is gitignored and not on Streamlit Cloud. Three options:
   - Option 1 (recommended, ~90 min): Pre-aggregate Level 1+2 into small CSVs in git; convert main file to parquet (~30 MB) + DuckDB lazy queries for Level 3.
   - Option 2 (~45 min): Aggregations + sampled Level 3 (top N per class). Loses full audit fidelity.
   - Option 3 (~60 min): Build feature, gracefully degrade Level 3 on Cloud when parquet missing.
2. **Confirm Streamlit Cloud subdomain rename** to `nigp-sourced-category-mapper.streamlit.app`. User said "i think" they did the UI rename — URL not yet confirmed loading.
3. **Polish:** `.docx` version of `outputs/HOW_TO_USE_NIGP_Mapping_JHK3.md`.
4. **Anthropic API key rotation** flagged 2026-04-30, status unconfirmed. Production is rules-only, so not blocking — but verify before any AI re-run.
5. **Optional:** scrub remaining "Chicago FMPS" references from internal scripts/methodology if user requests. Currently intentional (factual data-source description, not branding).

---

## When in doubt

Read `spend-analysis/METHODOLOGY_JHK3.md` first — it's the canonical narrative. Then ask James before relitigating any locked decision above.
