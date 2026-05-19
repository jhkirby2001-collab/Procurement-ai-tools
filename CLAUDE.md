# NIGP-Sourced Procurement Category Mapper — Project Guide

This repo is the home of the **NIGP-Sourced Procurement Category Mapper**, an independent public-sector procurement classification tool. The build is complete and in production; ongoing work is incremental polish and Phase 2 (interactive drill-down).

**Owner:** James H. Kirby III, CSCP, MS-SCM (jhkirby2001@yahoo.com)
**Output naming convention:** `_JHK3` suffix on every deliverable file
**Status:** Production. Latest commit `1cd1940` on `origin/main`. Repo is PUBLIC.

---

## What this project does

Classifies 784,556 historical public-sector procurement rows into a three-level taxonomy on every row:

**Business Category (17) → NIGP Class (3-digit) → NIGP Item (5-digit)**

Inputs to the classifier: description text + Chicago FMPS account/object/fund codes ONLY. Vendor names and any EY-supplied NIGP codes are NOT inputs.

**Production batch metrics (post-resolver, 2026-05-14):** 100% mapped, 0 rows in review queue. Coverage breakdown:
- Tier 1 — Keyword rule (246 curated + 6,766 AI-mined): 688,044 (87.7%)
- Tier 2 — Chicago FMPS account-code pattern: 2,028 (0.3%)
- Tier 3 — AI-assist resolver (saved one-time AI output, no new API call): 93,518 (11.9%)
- Tier 4 — Unclassified — No Description (no usable text in any of 4 fields): 966 (0.1%)
- Review_Flag=Yes (terminal review queue): 0 (0.0%)

---

## Key paths

**Rule files (procurement-staff editable, version-controlled):**
- `spend-analysis/data/reference/keyword_rules_DRAFT_JHK3.csv` — 246 hand-curated rules (edit these freely)
- `spend-analysis/data/reference/keyword_rules_from_ai_JHK3.csv` — 6,766 AI-mined rules (frozen — do not regenerate)
- `spend-analysis/data/reference/account_patterns_DRAFT_JHK3.csv` — 6 subgrant account patterns
- `spend-analysis/data/reference/business_categories_JHK3.csv` — canonical 138-row NIGP-class → Business-Category map
- `spend-analysis/data/reference/business_categories_summary_JHK3.csv` — 17-row summary
- `spend-analysis/data/reference/nigp_codes_{3,5,10}digit_JHK3.csv` — NIGP code reference tables

**Production scripts:**
- `spend-analysis/scripts/classifier_JHK3.py` — production classifier, dual-mode (batch + single-record), runs Tier 1 + Tier 2
- `spend-analysis/scripts/resolve_review_queue_JHK3.py` — Tier 3 AI-assist resolver (reads saved AI output, no new API call)
- `spend-analysis/scripts/build_leadership_deliverables_JHK3.py` — regenerates Word/Excel summaries
- `spend-analysis/scripts/build_sop_JHK3.py` — regenerates the SOP .docx
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
- `outputs/NIGP_Web_App_Factsheet_JHK3.md` — short Layer-1 factual brief about the web app

**Documentation:**
- `spend-analysis/METHODOLOGY_JHK3.md` v1.2 — canonical methodology narrative (read before debating rationale or AI defensibility)
- `NIGP_PROJECT_INDEX.md` — repo index
- `README.md` — public-facing readme

**Web app:**
- `streamlit_app.py` (repo root) — six-page Streamlit app: Classify → Bulk Classify → Procurement Taxonomy Logic → Methodology → Business Categories → Rule Lookup
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
| 3 | Classifier inputs: description text + Chicago FMPS account/object/fund codes ONLY. Vendor and EY-supplied NIGP codes are NOT inputs. |
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
