# City of Chicago — Procurement Taxonomy & Classification Methodology

**Prepared for:** City of Chicago Department of Procurement Services Leadership
**Author:** James H. Kirby III, CSCP, MS-SCM
**Document version:** 1.1 — finalized 2026-04-30 with production coverage results
**Project:** NIGP-Aligned Procurement Taxonomy & Classification Engine

> **Format note for the reader:** This document is structured for direct copy-paste into Microsoft Word. Section headings map to Word's Heading styles. Citations and URLs at the end can be hyperlinked in Word.

---

## Executive Summary

The City of Chicago has not historically maintained its own commodity classification system for procurement spend. The "EY raw data" file analyzed here is an EY consulting deliverable — useful as raw transactional data, but **not** an authoritative City of Chicago classification of what the City buys.

This project delivers Chicago's **first internally-owned commodity taxonomy** plus a **reusable, defensible classification engine** that the City can run on this historical file and on every future procurement extract. The deliverable is fully self-contained: no recurring vendor dependencies, no per-classification API calls, no third-party software requirements beyond standard open-source Python.

The engine classifies 784,556 historical PO and invoice line records across 23 years of City of Chicago spend (2002-2025). Each record is assigned a three-level classification — Business Category, NIGP Class, and NIGP Item — along with a confidence score, audit trail, and human-review flag.

**Headline outcome (production run, 2026-04-30):** 86.4% of all 784,556 rows were auto-classified by deterministic rule. The high-confidence (Review_Flag=No) classification covers 82.2% of rows, leaving a 17.8% review queue for procurement-staff triage. End-to-end runtime on a standard workstation was 14 minutes.

Spend dollars are **not** an input to the classifier. The classifier answers "what was bought" using description text and Chicago's own FMPS account/object/fund codes. Spend analysis is a separate downstream activity that consumes this taxonomy.

---

## 1. Project Objective and Scope

**Objective:** Build a transparent, defensible, accuracy-first classification engine that:

- Uses the **NIGP commodity code framework** as the public-standard backbone for inter-agency comparability and audit defensibility.
- Layers a custom **Business Category rollup** on top, designed for Chicago's organizational reporting needs and business-friendly leadership review.
- Operates in **two modes**: batch processing of large historical files, and single-record classification for live PO/requisition work.
- **Prioritizes accuracy over auto-coding.** Records that cannot be classified with sufficient confidence are flagged for human review rather than guessed.
- Is **fully repeatable** on future raw extracts with no code modification — the rule files are externalized as CSVs that procurement staff can edit directly.

**Scope:** The historical analysis covers the EY raw data file. The classifier itself is general-purpose and applies to any procurement file with description and account-code fields.

---

## 2. Source Data: EY Raw Data File

| Attribute | Value |
|-----------|------|
| Filename | `ey raw data.xlsx` |
| Size | 326 MB |
| Rows | 784,556 (sheet `Report 1`) |
| Columns | 87 |
| Origin | EY consulting engagement deliverable |
| Date coverage (PO Creation Date) | 2002-10-12 to 2025-05-27 (~23 years, all rows) |
| Date coverage (AP Voucher Date) | 2020-03-03 to 2025-05-28 (~5 years) |
| Date coverage (Payment Check Date) | 2023-01-03 to 2025-05-28 (~2.5 years) |

**Important caveat (driving design decisions):** The EY file already contains a populated `NIGP Code` and `NIGP Description` column on 30.4% of rows. These are EY's classification work, not Chicago's authoritative system. **The classifier does not use these EY-supplied codes as inputs.** They are preserved in the raw file for later audit cross-reference (e.g., "did our independent classifier agree with EY's prior labels?") but they do not drive any classification decision.

This decision was made deliberately: building Chicago's own taxonomy means classifying every record independently from the description text, not re-using a prior consultant's work product.

**Known data quality issues** detected during profiling:

- One Contract Start Date appears as `0203-12-31` (likely a 2003 typo)
- Maximum Contract End Date is `2072-12-31` (placeholder for open-ended contracts)
- 17.8% of `PO Description` values are uninformative ("Misc", "Per contract", etc.) — but the line-level descriptions (`PO Item Description`, `AP Invoice Line Description`) are populated on 100% of rows with substantive content, so these serve as the primary description signal.
- Only 967 of 784,556 rows (0.1%) have no usable description across any of the four description fields.

---

## 3. Three-Level Taxonomy Design

Every classified row carries three taxonomy levels, providing both business-friendly rollup for leadership reporting and audit-defensible NIGP alignment for sourcing analysis.

### 3.1 Level 1 — Business Category (Custom Chicago Rollup, 17 buckets)

Designed to be mutually exclusive, collectively exhaustive, and immediately recognizable to City leadership and audit reviewers. The 17 categories are:

| # | Business Category | What's in it |
|---|-------------------|--------------|
| 1 | Facilities Operations & Maintenance | HVAC, plumbing, lighting, doors, building hardware |
| 2 | Public Safety, Uniforms & PPE | Police/fire uniforms, badges, fire equipment, PPE |
| 3 | Construction Materials | Concrete, asphalt, steel, pipe, rebar |
| 4 | Vehicles & Fleet | Vehicles, fleet parts, repairs, fuel-related |
| 5 | Professional & Administrative Services | Management services, security guards, records storage, weather forecasting |
| 6 | Office, Print & Marketing | Office supplies, printing, advertising, signage |
| 7 | Equipment Rental & Leasing | Heavy-equipment rentals, copier leases, traffic-control rentals |
| 8 | Janitorial, Sanitation & Waste | Cleaning supplies, dumpsters, disposal, portable-toilet servicing |
| 9 | Heavy Equipment & Machinery | Loaders, sweepers, lifts, sprayers (purchased, not rented) |
| 10 | IT, Telecom & Audio/Visual | Network services, telephony, A/V, in-car video |
| 11 | Landscaping, Grounds & Irrigation | Landscape services, irrigation, pest management |
| 12 | Chemicals & Water Treatment | Industrial chemicals, water treatment, deicers |
| 13 | Medical & Health Services | Medical gases, AEDs, exam services, behavioral health |
| 14 | Animal Care & Veterinary | Veterinary supplies, animal feed, beekeeping |
| 15 | Furniture & Furnishings | Office furniture, mattresses, fitness equipment |
| 16 | Construction & Trades Services | Bridge construction, expansion-joint repair, contracted tradesmen |
| 17 | Grants & Pass-Through Funding | Subgrant disbursements to community-based organizations (DFSS, CDPH, BACP, DOH, DPD, DCASE) |

**Note on category 17 (Grants & Pass-Through Funding):** Subgrant disbursements are not commodity purchases; they are financial transfers to subrecipients that deliver social-service, public-health, business-development, and housing programs on the City's behalf. They have different procurement governance, different audit treatment, and different stakeholder visibility than commodity purchases. Lumping subgrants with professional services obscures meaningful structure, so they have their own bucket. This category does not have an NIGP class — NIGP is a commodity classification framework and does not contain a "subgrants" code.

### 3.2 Level 2 — NIGP 3-digit Class

Each Business Category maps to one or more NIGP 3-digit Classes. NIGP is a public, inter-agency-compatible commodity framework used by U.S. public procurement entities. Aligning Chicago's taxonomy to NIGP enables:

- **Inter-agency benchmarking** — peer comparisons against other municipalities.
- **Audit defensibility** — the assignment can always be traced back to a published NIGP class.
- **Future portability** — the codes will continue to make sense if Chicago later licenses a full NIGP catalog or migrates to a successor standard.

The 138 distinct 3-digit Classes appearing in the EY data have been mapped 1:1 to the 17 Business Categories. The mapping is preserved in `data/reference/business_categories_JHK3.csv` and is editable by procurement staff without touching code.

### 3.3 Level 3 — NIGP 5-digit Class-Item

When a description supports specificity, the classifier assigns a 5-digit NIGP Class-Item code. When confidence at the 5-digit level is insufficient, the row is assigned the broader 3-digit Class only and flagged accordingly via the `NIGP_Match_Level` field (`exact` / `broad` / `review`).

---

## 4. Inputs to the Classifier

The classifier accepts **only the following signals** as inputs:

1. **Transaction description text** — `PO Description`, `PO Item Description`, `AP Invoice Line Description`, `Invoice Distribution Description`. The classifier selects the first non-blank, non-uninformative description across these four fields as `Description_Best`.
2. **Chicago FMPS account/object/fund codes** — used as supplemental signal where description text alone is ambiguous.

**Inputs explicitly excluded by design:**

- **Vendor name.** The same vendor often sells in multiple commodity categories (e.g., a generic supplier may sell IT, office, and janitorial goods on one contract). Vendor-based inference introduces misclassification risk that descriptions and account codes do not.
- **EY-supplied NIGP codes and Commodity codes.** These represent EY's prior classification work, not Chicago's authoritative judgment. The City's principle is to build its own classification independently.

---

## 5. Rule Hierarchy and Order of Operations

The classifier evaluates each row through a deterministic three-step rule pipeline. The first rule that fires assigns the classification; subsequent rules do not run for that row. This produces a clean audit trail: every classified row carries the exact rule that drove the decision.

### Pass 1 — Keyword Rules on Description Text

The classifier loads two keyword-rule files and evaluates them as a single pool:

- `data/reference/keyword_rules_DRAFT_JHK3.csv` — 148 hand-curated rules drafted by the project author. These take priority within any given match-type tier.
- `data/reference/keyword_rules_from_ai_JHK3.csv` — 6,766 rules harvested from the one-time AI pattern-mining run (see §6). Each rule carries provenance metadata (model, confidence, source row count, AI reasoning) in its `notes` field.

Three match types are supported:

- **`exact`** — full description string match
- **`starts_with`** — prefix match (used for department program-code prefixes such as `DFSS-`, `BACP-`)
- **`contains`** — case-insensitive substring match (the bulk of rules)

Match priority is `exact > starts_with > contains`. The first matching rule wins.

### Pass 2 — Account-Code Patterns (Chicago FMPS supplemental signal)

`data/reference/account_patterns_DRAFT_JHK3.csv` maps Chicago FMPS account codes to (Business Category, NIGP Class). This pass only fires for rows that did not match any keyword rule.

The most consequential account-pattern rule is for the `220xxx` subgrant-disbursement account series (220005, 220044, 220100, 220801, 220300, 220999). Empirical analysis confirmed these accounts are 93-100% used for subgrant disbursements regardless of description text quality, which makes them a reliable signal for rows where a department program tag in the description would otherwise leave the row in human review.

### Pass 3 — Human Review

Any row that does not match a keyword rule or account pattern is flagged with `Classification_Method = "human_review"` and `Review_Flag = "Yes"`. These rows enter the review queue rather than being classified by guess.

---

## 6. Use of Artificial Intelligence — Limited, Transparent, One-Time

To establish defensibility upfront: **the production classification system Chicago operates is rules-only.** There is no AI in the runtime path. Any City staff member can run the production classifier from their workstation with no API key, no internet dependency, no recurring vendor cost, and no per-classification charge.

**However**, AI was used **once** during the initial taxonomy build for a specific, bounded purpose: to mine recurring patterns from the long tail of descriptions where no hand-curated rule applied. The AI's output was harvested into the keyword rules file and then frozen. The AI is not consulted at classification time; it is consulted only at rule-construction time, similar to how a procurement analyst might consult subject-matter expertise to draft a rule.

### How AI was used

- **Model:** Anthropic Claude Haiku 4.5
- **Invocation:** A single batch run over 30,342 unique long-tail descriptions (every distinct description in the EY file that the hand-curated rules did not already classify).
- **Constraints:** The AI's output was constrained by JSON schema with a closed enumeration of the 17 Business Categories and 138 NIGP 3-digit Classes. The model could not invent codes outside the catalog.
- **Input:** Each AI call sent only the description text plus a system prompt containing the controlled vocabulary. Vendor name, EY codes, Chicago account codes, and any other operational context were not passed to the AI.
- **Output:** Each unique description received a proposed Business Category, optional NIGP class, confidence rating, and short reason.
- **Confidence distribution returned by the model:** 6,419 high (21.2%), 11,645 medium (38.4%), 12,278 low (40.5%).
- **Promotion to rules:** AI proposals at high confidence were promoted automatically into the rule file. Medium-confidence proposals were promoted only if the description recurred in at least 5 source rows (frequency-of-occurrence threshold) — 1,998 of the 11,645 medium proposals met this bar; the remaining 9,647 long-tail singletons were dropped. Low-confidence proposals were never promoted; those descriptions remain in the human-review queue. Net result: **6,766 AI-mined rules promoted** out of 30,342 candidates (22.3%).
- **Audit trail:** Every AI-mined rule carries a `notes` field recording the model name, confidence level, source row count, and the model's reasoning. This record is permanent and reviewable at any time.

### Why this AI use is defensible

1. **One-time, not ongoing.** Production runs do not call AI. Once Chicago accepts the harvested rules, the AI dependency is severed.
2. **Bounded by Chicago's own taxonomy.** The AI cannot output codes outside the 17 categories and 138 classes that Chicago's procurement leadership approved.
3. **Transparent.** Every AI-mined rule is auditable. The source CSV (`keyword_rules_from_ai_JHK3.csv`) carries provenance metadata in every row.
4. **Reviewable.** Procurement staff can edit any AI-mined rule, demote it, or remove it entirely. AI-mined rules have no special status in the classifier — they are evaluated alongside hand-curated rules using the same matching logic.
5. **Replaceable.** If Chicago later decides to remove all AI-mined content, the file can be deleted and the classifier still operates on hand-curated rules alone.

---

## 7. Confidence, Match Levels, and Review Flags

Every classified row carries a confidence triple:

| Field | Values | Meaning |
|-------|--------|---------|
| `NIGP_Match_Level` | `exact` / `broad` / `review` / *empty* | `exact` = 5-digit Class-Item assignable; `broad` = 3-digit Class only; `review` = best broader category but flagged; *empty* = no match (Grants category) |
| `Classification_Confidence` | `High` / `Medium` / `Low` | Coarse-grained confidence label |
| `Confidence_Score` | 0.0 to 1.0 (float) | Fine-grained numeric confidence |

Rows with `Confidence_Score < 0.75` (the locked threshold) are auto-flagged with `Review_Flag = Yes`, regardless of which method classified them. A separate exceptions file lists all review-flagged rows for procurement-staff triage.

The `Classification_Reason` field on every row records the exact rule pattern that fired and any notes — this provides a complete audit trail end-to-end.

---

## 7.5 Production Run Results (2026-04-30)

The classifier was run end-to-end against all 784,556 rows of the EY file. Wall-clock runtime: 14 minutes 37 seconds on a standard codespace VM. Coverage results:

| Outcome | Pre-AI baseline (148 curated rules) | Final (6,914 rules: 148 curated + 6,766 AI-mined) | Δ |
|---|---:|---:|---:|
| Classified by keyword rule | 609,541 (77.7%) | 675,906 (86.2%) | +66,365 rows |
| Classified by account pattern | (folded into above) | 2,179 (0.3%) | — |
| **Total auto-classified** | **609,541 (77.7%)** | **678,085 (86.4%)** | **+68,544 rows / +8.7 pp** |
| Sent to human review | 175,015 (22.3%) | 106,471 (13.6%) | −68,544 rows |
| Review_Flag=Yes (full QA queue) | 408,037 (52.0%) | **139,868 (17.8%)** | **−268,169 rows / −34.2 pp** |

**Attribution of the AI-mined rule contribution:**

- 148 curated rules accounted for 609,889 row-hits (90.2% of all classified rows). Curated rules dominate volume because they target the highest-frequency patterns (department program-code prefixes, common rental and supply patterns).
- 6,766 AI-mined rules accounted for 66,017 row-hits (9.8% of all classified rows), averaging ~10 rows per rule — exactly the long-tail role they were designed for.
- The largest *Review_Flag* improvement came from AI-mined exact-match rules superseding broader, lower-confidence curated `contains` rules. Approximately 200,000 rows that previously hit a `review`-level rule (or no rule at all) now hit a `broad`-level AI exact-match and exit the QA queue.

**Top categories by AI-mined contribution** (rows newly classified by AI-mined rules):

| Business Category | Rows |
|---|---:|
| Grants & Pass-Through Funding (middle-of-string program tags) | 15,974 |
| IT, Telecom & Audio/Visual | 9,460 |
| Facilities Operations & Maintenance | 9,372 |
| Construction & Trades Services | 8,865 |
| Professional & Administrative Services | 3,878 |
| Office, Print & Marketing | 3,457 |
| Landscaping, Grounds & Irrigation | 3,445 |
| Chemicals & Water Treatment | 3,211 |
| Vehicles & Fleet | 2,585 |
| Other 8 categories | 5,770 |

**Final Business Category distribution** (auto-classified rows only):

| Business Category | Rows | % of classified |
|---|---:|---:|
| Grants & Pass-Through Funding | 218,126 | 32.2% |
| Equipment Rental & Leasing | 81,920 | 12.1% |
| Professional & Administrative Services | 59,643 | 8.8% |
| Office, Print & Marketing | 53,940 | 8.0% |
| Vehicles & Fleet | 44,728 | 6.6% |
| IT, Telecom & Audio/Visual | 42,802 | 6.3% |
| Facilities Operations & Maintenance | 42,490 | 6.3% |
| Construction Materials | 25,172 | 3.7% |
| Janitorial, Sanitation & Waste | 23,653 | 3.5% |
| Landscaping, Grounds & Irrigation | 23,234 | 3.4% |
| Public Safety, Uniforms & PPE | 17,384 | 2.6% |
| Construction & Trades Services | 13,908 | 2.1% |
| Medical & Health Services | 10,908 | 1.6% |
| Heavy Equipment & Machinery | 9,222 | 1.4% |
| Chemicals & Water Treatment | 4,606 | 0.7% |
| Animal Care & Veterinary | 4,161 | 0.6% |
| Furniture & Furnishings | 2,188 | 0.3% |

The remaining 106,471 rows are routed to human review — predominantly thin descriptions ("Misc supplies", "Per contract") that legitimately cannot be classified from text alone, plus department-specific codes that have not yet been encoded as rules. As described in §11, these flow into the procurement-staff triage cadence and are converted into new rules over time.

---

## 8. Output Specification

The classified dataset is delivered as a **lean ~16-column file** rather than preserving all 87 raw columns. The original 87-column raw file is preserved untouched at `data/raw/` and can be joined back if a kitchen-sink view is needed for ad-hoc analysis.

Output columns: `Transaction_ID`, `Date`, `Department`, `Vendor` (preserved for human readability only — not a classifier input), `Description_Best`, `Amount`, `Business_Category`, `NIGP_Class_3digit`, `NIGP_Item_5digit`, `NIGP_Code_Assigned`, `NIGP_Match_Level`, `Classification_Confidence`, `Confidence_Score`, `Classification_Method`, `Review_Flag`, `Classification_Reason`.

Three artifact files are produced on every batch run:

1. **`classified_full_JHK3.csv`** — every row with its full classification.
2. **`classified_review_JHK3.csv`** — review-flagged subset for staff triage.
3. **`classifier_coverage_report_JHK3.csv`** — diagnostic report of which rules fired how many times, used to identify rule gaps.

---

## 9. Single-Record Mode

The same classifier core function powers a paste-a-description CLI tool. A procurement analyst can paste a single PO or requisition description and immediately receive the proposed Business Category, NIGP code, confidence, and reason — without a batch file or a database query.

This is the foundation for future integration with Chicago's DPS PO-creation workflow: as new POs flow through the system, they can be classified at point-of-creation rather than retrospectively.

---

## 10. Key Assumptions and Limitations

**Assumptions:**

1. The four description fields in the EY data structure (`PO Description`, `PO Item Description`, `AP Invoice Line Description`, `Invoice Distribution Description`) will continue to be available in future Chicago procurement extracts.
2. The Chicago FMPS account-code structure (Fund, Account, Cost Center, etc.) will remain stable.
3. The 138 NIGP 3-digit Classes derived from the EY file are sufficient for current Chicago commodity coverage. New commodity types may require expanding the working catalog.
4. The Business Category structure approved here will be reviewed by leadership at least annually for fitness with reporting needs.

**Limitations:**

1. **Single-source dataset.** The taxonomy was derived from one consulting deliverable. New Chicago extracts may reveal commodity types not represented in the EY file.
2. **Time-agnostic by design.** The classifier does not consider transaction date when classifying. Spend changes over 23 years (commodity inflation, contract restructuring, departmental reorganizations) do not affect classification consistency.
3. **No vendor signal.** Per the locked design decision, vendor name is not used as a classification input. This is by choice — the classification should reflect what was bought, not who sold it. Procurement staff who want vendor-based views can produce those separately from the classified output.
4. **Description quality dependent.** The classifier's accuracy is bounded by description quality. Thin descriptions ("Misc supplies") legitimately cannot be classified from text alone and are correctly routed to human review rather than guessed.
5. **NIGP working set is partial.** The 138 classes derived from the EY file are a subset of the full NIGP catalog (which covers thousands of codes). Future work should consider licensing a full NIGP catalog from Periscope Holdings or sourcing a comparable public alternative.

---

## 11. Governance and Future Maintenance Recommendations

1. **Annual taxonomy review.** Procurement leadership should review the 17 Business Categories annually for continued fitness with reporting needs. Edits are made directly in `business_categories_JHK3.csv`.
2. **Rule-file ownership.** Designate a procurement analyst as owner of `keyword_rules_DRAFT_JHK3.csv`. New rules can be added as new commodity types appear; outdated rules can be retired.
3. **Review queue triage cadence.** The review-flagged subset should be triaged on a monthly or quarterly cadence. Triaged decisions are converted into new rules, shrinking the review queue over time.
4. **Single-record feedback loop.** When the planned DPS integration is built, analyst overrides of the classifier's single-record proposals become candidate rule data. Over time, the rule base grows and the classifier gets more accurate without re-running AI.
5. **NIGP catalog expansion.** Consider procuring a Periscope NIGP license or comparable subscription. This expands the working catalog from ~138 codes to the full ~9,000-code NIGP standard.
6. **Independent audit benchmark.** Once each year, sample 100 classified rows at random and have a procurement analyst independently classify them blind. Compare against the classifier's output. Track the agreement rate over time.

---

## 12. Technical Architecture (Brief)

**Repository layout:**

```
spend-analysis/
├── data/
│   ├── raw/ey raw data.xlsx                     # untouched source
│   ├── reference/                               # editable rule files
│   │   ├── nigp_codes_3digit_JHK3.csv
│   │   ├── nigp_codes_5digit_JHK3.csv
│   │   ├── nigp_codes_10digit_JHK3.csv
│   │   ├── business_categories_JHK3.csv
│   │   ├── keyword_rules_DRAFT_JHK3.csv         # hand-curated
│   │   ├── keyword_rules_from_ai_JHK3.csv       # AI-mined (auditable)
│   │   └── account_patterns_DRAFT_JHK3.csv
│   └── processed/
│       ├── classified_full_JHK3.csv             # the deliverable
│       ├── classified_review_JHK3.csv
│       └── classifier_coverage_report_JHK3.csv
└── scripts/
    ├── classifier_JHK3.py                       # production classifier
    ├── build_*.py                               # rule and reference builders
    └── ai_classify_JHK3.py                      # one-time AI mining (not run in production)
```

**Runtime dependencies:** Python 3.x, pandas, pyarrow. Standard open-source. Runs on any laptop or City workstation with no special hardware or network access.

**Single-record mode:**

```
python scripts/classifier_JHK3.py --describe "your description here"
```

**Batch mode:**

```
python scripts/classifier_JHK3.py --batch
```

---

## 13. Sources and References

1. **NIGP — Institute for Public Procurement.** Originating organization for the NIGP Commodity/Services Code framework. Public homepage: <https://www.nigp.org/>
2. **City of Chicago Department of Procurement Services.** Operating context for this project. Public homepage: <https://www.chicago.gov/city/en/depts/dps.html>
3. **City of Chicago FMPS (Financial Management & Purchasing System).** Source system for account/object/fund codes used as supplemental classification signal.
4. **Anthropic — Claude API.** Provider of the AI service used in the one-time pattern-mining phase. Public homepage: <https://www.anthropic.com/>
5. **NIGP Commodity Code public reference.** General overview of NIGP code structure available via NIGP organizational documentation. Note: a fully licensed NIGP catalog is available commercially from Periscope Holdings; this project's working catalog was derived from codes present in the source EY file rather than a licensed catalog.
6. **Pandas** (data processing library). Public homepage: <https://pandas.pydata.org/>
7. **Project repository** (internal): all builder scripts, reference files, and the classifier are committed to the project repository under `spend-analysis/`. Each file carries a docstring describing its purpose and a `_JHK3` suffix indicating authorship and version.

---

## Appendix A — Decisions Locked During Project Build

For audit traceability, the following design decisions were made during taxonomy construction and locked in writing:

| # | Decision | Rationale |
|---|----------|-----------|
| 1 | Three-level taxonomy on every row: Business Category → NIGP Class → NIGP Item | Serves both executive dashboards (Business Category) and sourcing/audit needs (NIGP codes) |
| 2 | AI assist allowed once during build, with strict guardrails | Rules-only would leave a 20-40% review pile; bounded one-time AI use shrinks it while remaining defensible |
| 3 | Classifier inputs: description text + Chicago FMPS account/object/fund codes only. Vendor and EY-supplied NIGP codes are NOT inputs | Chicago must own its own classification; vendor isn't a reliable signal; EY codes are another consultant's work product |
| 4 | Lean ~16-column output instead of preserving all 87 raw columns | Raw file preserved separately; downstream analysis doesn't need 87 columns |
| 5 | Dual-mode classifier (batch + single-record), same core function | Future-state: procurement analysts paste PO descriptions into a tool, get instant category + confidence |
| 6 | Externalized rule files (`keyword_rules.csv`, `account_patterns.csv`) | Procurement staff can update rules without touching Python code |
| 7 | Feedback loop: analyst overrides become new rule data | AI/rule usage shrinks over time as the rule base grows |
| 8 | Classify all 784,556 rows regardless of date | Classifier is time-agnostic; date filtering is a downstream concern |

---

*End of methodology document.*
