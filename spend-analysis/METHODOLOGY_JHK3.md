# City of Chicago — Procurement Taxonomy & Classification Methodology

**Prepared for:** City of Chicago Department of Procurement Services Leadership
**Author:** James H. Kirby III, CSCP, MS-SCM
**Document version:** 1.2 — updated 2026-05-14 with review-queue elimination and 4-tier coverage results (originally finalized 2026-04-30)
**Project:** NIGP-Aligned Procurement Taxonomy & Classification Engine

> **Format note for the reader:** This document is structured for direct copy-paste into Microsoft Word. Section headings map to Word's Heading styles. Citations and URLs at the end can be hyperlinked in Word.

---

## Executive Summary

The City of Chicago has not historically maintained its own commodity classification system for procurement spend. The "EY raw data" file analyzed here is an EY consulting deliverable — useful as raw transactional data, but **not** an authoritative City of Chicago classification of what the City buys.

This project delivers Chicago's **first internally-owned commodity taxonomy** plus a **reusable, defensible classification engine** that the City can run on this historical file and on every future procurement extract. The deliverable is fully self-contained: no recurring vendor dependencies, no per-classification API calls, no third-party software requirements beyond standard open-source Python.

The engine classifies 784,556 historical PO and invoice line records of City of Chicago spend (AP activity years 2017, 2020, 2021, 2023). Each record is assigned a three-level classification — Business Category, NIGP Class, and NIGP Item — along with a confidence score, audit trail, and an explicit tier label identifying which deterministic source (curated rule, AI-mined rule, account pattern, or AI-assist resolver) produced the assignment.

**Headline outcome (production run, 2026-05-14):** All 784,556 rows are mapped to a Business Category — 100% coverage, zero rows in the review queue. Classification is distributed across four deterministic tiers: 87.7% by curated/AI-mined keyword rule, 11.9% by AI-assist fallback (resolver consumes the saved 2026-04-30 AI mining output — no new API call), 0.3% by Chicago FMPS account-code pattern, and 0.1% explicitly tagged "Unclassified — No Description" (the residue of rows with no usable description text across any of four description fields). End-to-end runtime on a standard workstation: ~15 minutes for the batch + ~5 seconds for the resolver pass.

Spend dollars are **not** an input to the classifier. The classifier answers "what was bought" using description text and Chicago's own FMPS account/object/fund codes. Spend analysis is a separate downstream activity that consumes this taxonomy.

---

## 1. Project Objective and Scope

**Objective:** Build a transparent, defensible, accuracy-first classification engine that:

- Uses the **NIGP commodity code framework** as the public-standard backbone for inter-agency comparability and audit defensibility.
- Layers a custom **Business Category rollup** on top, designed for Chicago's organizational reporting needs and business-friendly leadership review.
- Operates in **two modes**: batch processing of large historical files, and single-record classification for live PO/requisition work.
- **Prioritizes transparency over a single "auto-classified" headline.** Every row carries an explicit confidence tier and tier-source label, so leadership and auditors can filter by deterministic-rule, account-pattern, or AI-assist tier rather than reading one blended accuracy number.
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

The classifier evaluates each row through a deterministic four-tier pipeline. The first tier that fires assigns the classification; subsequent tiers do not run for that row. This produces a clean audit trail: every classified row carries the exact source that drove the decision.

### Tier 1 — Keyword Rules on Description Text

The classifier loads two keyword-rule files and evaluates them as a single pool:

- `data/reference/keyword_rules_DRAFT_JHK3.csv` — 246 hand-curated rules drafted by the project author. These take priority within any given match-type tier.
- `data/reference/keyword_rules_from_ai_JHK3.csv` — 6,766 rules harvested from the one-time AI pattern-mining run (see §6). Each rule carries provenance metadata (model, confidence, source row count, AI reasoning) in its `notes` field.

Three match types are supported:

- **`exact`** — full description string match
- **`starts_with`** — prefix match (used for department program-code prefixes such as `DFSS-`, `BACP-`)
- **`contains`** — case-insensitive substring match (the bulk of rules)

Match priority is `exact > starts_with > contains`. The first matching rule wins.

### Tier 2 — Account-Code Patterns (Chicago FMPS supplemental signal)

`data/reference/account_patterns_DRAFT_JHK3.csv` maps Chicago FMPS account codes to (Business Category, NIGP Class). This tier only fires for rows that did not match any keyword rule.

The most consequential account-pattern rule is for the `220xxx` subgrant-disbursement account series (220005, 220044, 220100, 220801, 220300, 220999). Empirical analysis confirmed these accounts are 93-100% used for subgrant disbursements regardless of description text quality, which makes them a reliable signal for rows where a department program tag in the description would otherwise be ambiguous.

### Tier 3 — AI-Assist Fallback (resolver pass over saved AI mining output)

Any row that did not match a keyword rule or account pattern is resolved by a post-classifier pass against the saved AI mining output (`data/processed/ai_classified_unique_descriptions_JHK3.csv` — see §6). The resolver looks up `Description_Best` in that CSV and fills `Business_Category` / `NIGP_Class_3digit` / `Classification_Method = "ai_assist"` / `Classification_Confidence = AI-{high|medium|low}`. **No new API call is made** — the resolver reads the saved 2026-04-30 mining output. This is a second, more aggressive consumption path for the same paid-for AI output (the original promotion-to-rules path was conservative; the resolver consumes the latent medium/low AI classifications at fill time rather than dropping them).

This tier preserves the "AI used once during build, rules-only at runtime" lock: there is no live AI dependency, no API key requirement at classification time, and the resolver's source CSV is committed and reviewable.

### Tier 4 — Unclassified — No Description

Any row that reaches this tier has no usable description text in any of the four description fields (see §2 caveat on data quality). These rows are explicitly tagged `Business_Category = "Unclassified — No Description"` / `Classification_Method = "no_description"` / `Review_Flag = "No"`. They are not flagged for review because there is nothing to classify — the data itself is the limitation, not the classifier.

### Result: zero rows in review queue

After the four tiers run, every row in the source file carries a Business Category. `Review_Flag = "Yes"` is no longer used as a terminal disposition. Procurement-staff time is redirected from triaging a queue to auditing the AI-assist tier and promoting any high-quality AI-assist matches into the curated rule file when patterns warrant it (see §11).

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
- **Promotion to rules (Tier 1 consumption path):** AI proposals at high confidence were promoted automatically into the rule file. Medium-confidence proposals were promoted only if the description recurred in at least 5 source rows (frequency-of-occurrence threshold) — 1,998 of the 11,645 medium proposals met this bar; the remaining 9,647 long-tail singletons were dropped from the rule file. Low-confidence proposals were never promoted to rules. Net result: **6,766 AI-mined rules promoted** out of 30,342 candidates (22.3%).
- **Resolver fallback (Tier 3 consumption path, added 2026-05-14):** The full saved CSV (`data/processed/ai_classified_unique_descriptions_JHK3.csv`, 30,342 rows) is also consumed by the resolver as Tier 3 of the classification pipeline (see §5). The resolver matches `Description_Best` against this CSV and applies the model's Business Category and 3-digit NIGP class even for medium-low confidence rows that were not promoted to rules. This is a second, more aggressive consumption of the same paid-for output — not a new API call. The resolver applies the canonical NIGP-Class → Business-Category alignment as a final integrity check.
- **Audit trail:** Every AI-mined rule and every AI-assist resolver assignment carries a `notes` field recording the model name, confidence level, source row count, and the model's reasoning. This record is permanent and reviewable at any time. Rows resolved at Tier 3 carry `Classification_Confidence = AI-high | AI-medium | AI-low`, so leadership can immediately see which classifications rest on the strongest AI evidence and which are inferential.

### Why this AI use is defensible

1. **One-time, not ongoing.** Production runs do not call AI. The single 2026-04-30 mining run produced a CSV; every downstream consumption of that CSV is a local file read. No API key, internet access, or per-classification charge is required to run the classifier or the resolver.
2. **Bounded by Chicago's own taxonomy.** The AI cannot output codes outside the 17 categories and 138 classes that Chicago's procurement leadership approved.
3. **Transparent.** Every AI-mined rule and every AI-assist resolver assignment is auditable. Source CSVs (`keyword_rules_from_ai_JHK3.csv`, `ai_classified_unique_descriptions_JHK3.csv`) are committed and carry full provenance metadata.
4. **Reviewable.** Procurement staff can edit any AI-mined rule, demote it, or remove it entirely. AI-mined rules have no special status in the classifier — they are evaluated alongside hand-curated rules using the same matching logic. AI-assist resolver assignments can be promoted into the curated rule file at any time.
5. **Replaceable.** If Chicago later decides to remove all AI-mined content, both the rule file and the resolver source CSV can be deleted, and the classifier still operates on hand-curated rules + account patterns alone. Coverage would drop, but the system would not break.
6. **Confidence-labeled, not laundered.** AI-assist resolver assignments are tagged `AI-high`, `AI-medium`, or `AI-low` in the output file. Leadership and auditors can filter or exclude any confidence tier they choose. The AI tier is not hidden inside a generic "auto-classified" bucket.

---

## 7. Confidence, Match Levels, and Review Flags

Every classified row carries a confidence triple:

| Field | Values | Meaning |
|-------|--------|---------|
| `NIGP_Match_Level` | `exact` / `broad` / `review` / *empty* | `exact` = 5-digit Class-Item assignable; `broad` = 3-digit Class only; `review` = best broader category but flagged; *empty* = no match (Grants category) |
| `Classification_Confidence` | `High` / `Medium` / `Low` | Coarse-grained confidence label |
| `Confidence_Score` | 0.0 to 1.0 (float) | Fine-grained numeric confidence |

The `Classification_Reason` field on every row records the exact rule pattern (or resolver source) that drove the decision — this provides a complete audit trail end-to-end.

**On the `Review_Flag` field (revised 2026-05-14):** Through the original 2026-04-30 production run, `Review_Flag = "Yes"` was set on any row scoring below 0.75 confidence, producing a 17.8% review queue. After the AI-assist resolver was added (see §5 Tier 3), no rows remain unresolved — every row carries a Business Category, including the 966-row "Unclassified — No Description" residue. `Review_Flag` is preserved in the output schema for backward compatibility and defaults to `"No"`. Procurement-staff time is redirected from queue triage to confidence-tier auditing and curated-rule promotion (see §11).

---

## 7.5 Production Run Results (2026-05-14)

The classifier was run end-to-end against all 784,556 rows of the EY file, followed by the AI-assist resolver pass. Wall-clock runtime: ~15 minutes for the batch + ~5 seconds for the resolver on a standard codespace VM. Coverage results:

| Classification tier | Rows | Share |
|---|---:|---:|
| Tier 1 — Keyword rule (246 curated + 6,766 AI-mined) | 688,044 | 87.7% |
| Tier 2 — Chicago FMPS account-code pattern | 2,028 | 0.3% |
| Tier 3 — AI-assist resolver (saved AI mining output) | 93,518 | 11.9% |
| Tier 4 — Unclassified — No Description | 966 | 0.1% |
| **Total mapped to a Business Category** | **784,556** | **100.0%** |
| `Review_Flag = "Yes"` (terminal review queue) | **0** | **0.0%** |

**How this compares to the 2026-04-30 baseline:**

| Outcome | 2026-04-30 baseline | 2026-05-14 final | Δ |
|---|---:|---:|---:|
| Auto-classified (rule + account) | 678,085 (86.4%) | 690,072 (88.0%) | +11,987 / +1.6 pp |
| AI-assist resolver | — | 93,518 (11.9%) | new tier |
| Unclassified — No Description | — | 966 (0.1%) | new explicit tag |
| `Review_Flag = "Yes"` | 139,868 (17.8%) | **0 (0.0%)** | **−139,868 / −17.8 pp** |

The 17.8-point reduction in the review queue came from three sources, in order of contribution:

1. **AI-assist resolver pass (Tier 3, ~93.5K rows).** The 2026-04-30 mining run had already produced confidence-tagged classifications for most review-queue uniques, but the original rule-promotion policy only consumed high-confidence proposals and frequent-medium proposals. The resolver consumes the latent medium/low classifications at fill time, applying them as fallback rather than dropping them.
2. **Curated rule expansion (148 → 246).** 14 new top-cluster rules and 15 rule-level flips from `match_level=review` to `broad` were added 2026-05-14, mapping rows that had been review-flagged not because they were unclassifiable but because no NIGP class was assignable (category-only classification is still a valid map). Categories like Equipment Rental & Leasing, Facilities Operations & Maintenance, and Grants & Pass-Through Funding picked up the largest share.
3. **Explicit Tier 4 tag for 966 no-description rows.** These were previously routed to review by default. The resolver now tags them `Unclassified — No Description` to be honest about the data limitation — the classifier never had usable text to classify, so reviewing them would produce no signal either.

**Final Business Category distribution** (post-resolver, all 784,556 rows):

| Business Category | Rows | % of total |
|---|---:|---:|
| Grants & Pass-Through Funding | 222,109 | 28.3% |
| Professional & Administrative Services | 108,576 | 13.8% |
| Equipment Rental & Leasing | 86,156 | 11.0% |
| Office, Print & Marketing | 62,319 | 7.9% |
| Facilities Operations & Maintenance | 54,899 | 7.0% |
| IT, Telecom & Audio/Visual | 51,744 | 6.6% |
| Vehicles & Fleet | 47,858 | 6.1% |
| Construction Materials | 29,466 | 3.8% |
| Janitorial, Sanitation & Waste | 23,903 | 3.0% |
| Landscaping, Grounds & Irrigation | 23,789 | 3.0% |
| Public Safety, Uniforms & PPE | 19,257 | 2.5% |
| Construction & Trades Services | 15,136 | 1.9% |
| Medical & Health Services | 12,784 | 1.6% |
| Heavy Equipment & Machinery | 12,166 | 1.6% |
| Chemicals & Water Treatment | 6,093 | 0.8% |
| Animal Care & Veterinary | 4,847 | 0.6% |
| Furniture & Furnishings | 2,488 | 0.3% |
| Unclassified — No Description | 966 | 0.1% |
| **Total** | **784,556** | **100.0%** |

**On the trade-off implicit in this state:** The 0.0% review queue is not the same statement as 0.0% misclassification risk. Tier 3 rows carry an `AI-medium` or `AI-low` confidence tag and were classified by AI proposal rather than deterministic rule. The `Classification_Confidence` field on every row exposes this transparently, so leadership and auditors can filter by tier when accuracy concerns warrant it. The choice to fill rather than queue was an explicit leadership directive to prefer 100% mapping with confidence tagging over a queue that would never realistically be triaged at 140,000-row volume.

---

## 8. Output Specification

The classified dataset is delivered as a **lean ~16-column file** rather than preserving all 87 raw columns. The original 87-column raw file is preserved untouched at `data/raw/` and can be joined back if a kitchen-sink view is needed for ad-hoc analysis.

Output columns: `Transaction_ID`, `Date`, `Department`, `Vendor` (preserved for human readability only — not a classifier input), `Description_Best`, `Amount`, `Business_Category`, `NIGP_Class_3digit`, `NIGP_Item_5digit`, `NIGP_Code_Assigned`, `NIGP_Match_Level`, `Classification_Confidence`, `Confidence_Score`, `Classification_Method`, `Review_Flag`, `Classification_Reason`.

Two primary artifact files are produced on every batch run plus one diagnostic:

1. **`outputs/NIGP_Mapping_JHK3.csv`** — the deliverable. Every row with its full classification, post-resolver. ~225 MB; gitignored.
2. **`outputs/NIGP_Mapping_Review_Queue_JHK3.csv`** — subset of `Review_Flag = "Yes"` rows. As of 2026-05-14 this file is effectively empty (0 rows); retained in the output schema for forward compatibility if future rule changes reintroduce a review tier.
3. **`spend-analysis/data/processed/classifier_coverage_report_JHK3.csv`** — diagnostic report of which rules fired how many times, used to identify rule gaps.

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
2. **Time-agnostic by design.** The classifier does not consider transaction date when classifying. Year-over-year spend changes (commodity inflation, contract restructuring, departmental reorganizations) do not affect classification consistency.
3. **No vendor signal.** Per the locked design decision, vendor name is not used as a classification input. This is by choice — the classification should reflect what was bought, not who sold it. Procurement staff who want vendor-based views can produce those separately from the classified output.
4. **Description quality dependent.** The classifier's accuracy is bounded by description quality. Thin but non-empty descriptions ("Misc supplies", "Per contract") that recurred in the source file received an AI proposal during the 2026-04-30 mining run and are now resolved at Tier 3 with an explicit `AI-low` or `AI-medium` confidence tag — leadership and auditors can filter on that tag where accuracy concerns warrant it. Rows with no usable description across any of the four description fields are tagged `Unclassified — No Description` (Tier 4) rather than being assigned a category, because the data itself provides nothing to classify.
5. **NIGP working set is partial.** The 138 classes derived from the EY file are a subset of the full NIGP catalog (which covers thousands of codes). Future work should consider licensing a full NIGP catalog from Periscope Holdings or sourcing a comparable public alternative.

---

## 11. Governance and Future Maintenance Recommendations

1. **Annual taxonomy review.** Procurement leadership should review the 17 Business Categories annually for continued fitness with reporting needs. Edits are made directly in `business_categories_JHK3.csv`.
2. **Rule-file ownership.** Designate a procurement analyst as owner of `keyword_rules_DRAFT_JHK3.csv`. New rules can be added as new commodity types appear; outdated rules can be retired. Curated-rule count has grown from 148 (2026-04-30) to 246 (2026-05-14) on this principle.
3. **AI-assist tier audit cadence (replaces review-queue triage).** With zero rows in the terminal review queue, staff time is redirected to confidence-tier auditing. On a monthly or quarterly cadence, sample rows where `Classification_Method = "ai_assist"` and `Classification_Confidence = AI-medium | AI-low`. For descriptions that recur in volume and the AI tier got right, promote the pattern into the curated rule file (Tier 1) so the next batch run no longer depends on AI-assist. This is the new mechanism by which the rule base grows.
4. **Single-record feedback loop.** When the planned DPS integration is built, analyst overrides of the classifier's single-record proposals become candidate rule data. Over time, the rule base grows and the classifier gets more accurate without re-running AI.
5. **NIGP catalog expansion.** Consider procuring a Periscope NIGP license or comparable subscription. This expands the working catalog from ~138 codes to the full ~9,000-code NIGP standard.
6. **Independent audit benchmark.** Once each year, sample 100 classified rows at random and have a procurement analyst independently classify them blind. Compare against the classifier's output. Track the agreement rate over time. Stratify the sample across tiers (e.g., 25 keyword-rule rows, 25 account-pattern rows, 25 AI-medium, 25 AI-low) so the audit reads agreement separately by tier rather than as a single blended number.

---

## 12. Technical Architecture (Brief)

**Repository layout:**

```
Procurement-ai-tools/
├── outputs/                                         # leadership-facing deliverables
│   ├── NIGP_Mapping_JHK3.csv                        # full classified file, post-resolver (gitignored)
│   ├── NIGP_Mapping_Review_Queue_JHK3.csv           # 0 rows after 2026-05-14 (gitignored)
│   ├── NIGP_Summary_for_Leadership_JHK3.xlsx
│   ├── NIGP_Executive_Brief_JHK3.docx
│   ├── NIGP_Methodology_for_Leadership_JHK3.docx
│   └── NIGP_SOP_JHK3.docx
└── spend-analysis/
    ├── data/
    │   ├── raw/ey raw data.xlsx                     # untouched source
    │   ├── reference/                               # editable rule files
    │   │   ├── nigp_codes_3digit_JHK3.csv
    │   │   ├── nigp_codes_5digit_JHK3.csv
    │   │   ├── nigp_codes_10digit_JHK3.csv
    │   │   ├── business_categories_JHK3.csv
    │   │   ├── keyword_rules_DRAFT_JHK3.csv         # 246 hand-curated
    │   │   ├── keyword_rules_from_ai_JHK3.csv       # 6,766 AI-mined (auditable)
    │   │   └── account_patterns_DRAFT_JHK3.csv
    │   └── processed/
    │       ├── ai_classified_unique_descriptions_JHK3.csv   # 30,342-row saved AI mining output (resolver source)
    │       └── classifier_coverage_report_JHK3.csv
    └── scripts/
        ├── classifier_JHK3.py                       # production classifier (rules + account patterns)
        ├── resolve_review_queue_JHK3.py             # AI-assist resolver (Tier 3, no API call)
        ├── ai_topup_uncovered_JHK3.py               # targeted AI top-up (held in reserve, requires API key)
        ├── audit_classifier_coverage_JHK3.py        # 63-phrase plain-English regression test
        ├── fix_ai_rule_category_mismatches_JHK3.py  # idempotent integrity fix
        ├── build_leadership_deliverables_JHK3.py    # regenerates .docx / .xlsx artifacts
        ├── build_sop_JHK3.py                        # regenerates the SOP .docx
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
| 2 | AI assist allowed once during build, with strict guardrails. The single 2026-04-30 mining output is consumed in two paths at build/fill time: (a) high-confidence proposals promoted into the rule file, (b) saved CSV consulted by the post-classifier resolver for any row that no rule or account pattern caught. No runtime API call in either path. | Rules-only alone left a 17.8% review pile that would never be triaged at 140K-row volume; the resolver consumes more of the same paid-for AI output as a fallback layer. Lock preserved: production is rules-only at runtime, no API key required to classify. |
| 3 | Classifier inputs: description text + Chicago FMPS account/object/fund codes only. Vendor and EY-supplied NIGP codes are NOT inputs | Chicago must own its own classification; vendor isn't a reliable signal; EY codes are another consultant's work product |
| 4 | Lean ~16-column output instead of preserving all 87 raw columns | Raw file preserved separately; downstream analysis doesn't need 87 columns |
| 5 | Dual-mode classifier (batch + single-record), same core function | Future-state: procurement analysts paste PO descriptions into a tool, get instant category + confidence |
| 6 | Externalized rule files (`keyword_rules.csv`, `account_patterns.csv`) | Procurement staff can update rules without touching Python code |
| 7 | Feedback loop: analyst overrides become new rule data | AI/rule usage shrinks over time as the rule base grows |
| 8 | Classify all 784,556 rows regardless of date | Classifier is time-agnostic; date filtering is a downstream concern |

---

*End of methodology document.*
