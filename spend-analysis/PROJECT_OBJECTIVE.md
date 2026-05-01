# Project Objective — NIGP Procurement Taxonomy & Classification Engine

> **Author:** James H. Kirby III, CSCP, MS-SCM
> **Date locked:** 2026-04-29
> **Source file under analysis:** `spend-analysis/data/raw/ey raw data.xlsx` (City of Chicago purchasing data; EY consulting deliverable)

---

You are my coding and data analysis assistant in this Claude Code project. I am working with an Excel/CSV file called "EY raw data" that contains 4 years of City of Chicago purchasing data. Your job is to help me analyze this file, design a practical procurement taxonomy with Categories and Subcategories mapped to NIGP codes, classify the historical transactions, and build a repeatable, accuracy-first process that I can reuse on all new raw purchasing data going forward.

## Important constraints and priorities

- Future-state should prioritize **accuracy over maximum automation**.
- If a record cannot be matched with sufficient confidence, **do not force a guess.** Assign the best broader NIGP class you can justify, flag it for review, and document why.
- The process must be **transparent and defensible** so it can be explained to City leadership, auditors, or other stakeholders.

## 1. Project goals

Help me achieve all of the following from the "EY raw data" file:

- Analyze and classify the years of purchasing transactions in the file.
- Create Categories and Subcategories that align to NIGP codes.
- Use NIGP as the governing classification framework for the taxonomy.
- Build a repeatable mapping logic and Python script/workflow for future raw files.
- Prioritize classification accuracy and defensibility over aggressive auto-coding.
- Produce a leadership-ready methodology explanation (Word-document style) that describes how everything was done.
- Provide a source list with verifiable links for any external references used (e.g., NIGP structure, taxonomy best practices) so leadership can independently vet the approach.

## 2. Data context

- Primary file: "EY raw data" (Excel, 326 MB, 784,558 rows × 87 columns, sheet `Report 1`).
- This file is an **EY consulting deliverable** — it is NOT a live City of Chicago system extract. The City does not currently have its own commodity categorization system. This project will produce that.
- Likely fields include: transaction ID, date, department, vendor, description, amount, contract number, account/object code, fund, commodity/NIGP code fields (if present), and other classification fields.
- Use all available fields (especially vendor, description, and any existing codes) to maximize classification accuracy.
- The NIGP Code column populated in some rows = EY's classification work, not authoritative. Treat as a starting hypothesis to validate, not ground truth.

First, please inspect the structure of "EY raw data" and tell me:
- What columns exist and their data types.
- Any obvious data quality issues (missing values, inconsistent formats, duplicates, etc.).

## 3. NIGP-based taxonomy requirement

Use the NIGP commodity/services code structure as the main classification framework:

- NIGP is a hierarchical code system used by public entities to classify commodities and services for purchasing and spend analysis.
- Wherever possible, assign the most specific credible NIGP code (preferably Class-Item / 5-digit level).
- If an exact detailed match is not supportable, assign the best broader NIGP class (3-digit) and mark the record as needing review rather than guessing.
- Design Categories and Subcategories that are understandable to business users but still trace back to NIGP, so they support both reporting and sourcing analysis.

## 4. What I want you to produce

### 4.1 Procurement taxonomy design

Create a hierarchical taxonomy for this data with **three levels on every row**:

- **Level 1: Business Category** — custom Chicago rollup (~12-20 buckets), business-friendly
- **Level 2: NIGP Class** — NIGP 3-digit Class
- **Level 3: NIGP Item** — NIGP 5-digit Class-Item

Map each level cleanly to NIGP codes / code ranges. Keep the taxonomy as mutually exclusive and collectively exhaustive as practical. Make it realistic for a municipal procurement environment (City of Chicago context).

### 4.2 Historical transaction classification

Using "EY raw data," classify each transaction and produce a **lean output** (NOT all 87 original columns) with these fields:

- Transaction ID / PO #
- Date
- Department
- Vendor
- Best available description
- Amount
- **Business_Category** (new)
- **NIGP_Class_3digit** (new)
- **NIGP_Item_5digit** (new)
- **NIGP_Code_Assigned** (new)
- **NIGP_Match_Level** (new — exact / strong / broad-class-only)
- **Classification_Confidence** (new — High / Medium / Low)
- **Confidence_Score** (new — numeric 0-1)
- **Classification_Method** (new — existing_NIGP / vendor_rule / keyword_rule / ai_assist / human_review)
- **Review_Flag** (new — Yes/No)
- **Classification_Reason** (new — short explanation)

The original raw 87-column file remains untouched in `data/raw/` — we never overwrite source data. If a kitchen-sink version is needed later, it can be produced via a one-line join.

## 5. Accuracy-first repeatable logic

Design a reusable classification framework that can be applied to future raw data files that look like "EY raw data":

Use a combination of:
- Existing NIGP or commodity fields (validated as a starting hypothesis, not auto-trusted)
- Vendor normalization and vendor-to-category mapping
- Description keyword and phrase rules
- Pattern matching on recurring line-item language
- Account/object/fund codes as additional signals
- AI assist (constrained to the NIGP code list) ONLY on the long tail when rules fail

**Rule hierarchy / order of operations:**

1. Existing NIGP/commodity codes — validate against vendor + description; if they agree → high confidence
2. Vendor-based rules
3. Description keyword/phrase rules
4. Account / object / fund code patterns
5. AI assist (Claude API), bounded by NIGP code list, returns confidence score
6. If confidence below threshold OR no match: assign best broader NIGP class + Review_Flag = Yes

Build this in a way that is easy to maintain as new vendors and patterns appear. Mappings live in external CSV/JSON files so a procurement analyst can update rules without touching Python code.

## 6. Reusable Python script / workflow

Create runnable Python code for a repeatable pipeline that operates on "EY raw data" and future similar files. The same core classifier function serves **two modes**:

- **Batch mode:** process a whole file (e.g., quarterly extracts)
- **Single-record mode:** classify a single PO/requisition description on demand (paste a description, get a category) — for use as new POs flow through Chicago's DPS system

Functions to:
- Read the input CSV/Excel
- Normalize column names and formats
- Clean and standardize vendor names
- Normalize and preprocess descriptions
- Apply the classification rules and mapping table
- Compute Classification_Confidence and Review_Flag
- Call AI assist on the long tail (constrained, scored)

Write out:
- A categorized output file (lean, ~13 columns)
- A separate exceptions/review file (all rows with Review_Flag = Yes or low confidence)
- A rule coverage report (which rules fired most often, where the gaps are)

Structure the code so the mapping table or rules can be updated without rewriting everything (externalize mappings into a CSV/JSON).

Before writing the full code:
- Outline the pipeline step-by-step.
- Propose the data structures used for mappings (e.g., dictionaries, lookup tables).

## 7. Leadership-ready methodology explanation

Produce a Word-document-style explanation (copy/paste from Claude Code into Word) that covers:

- Project objective and scope
- Description of the "EY raw data" file and its date coverage
- Fields used for classification and why
- How the Category/Subcategory taxonomy was designed
- How NIGP alignment was determined
- The rule hierarchy and order of operations
- **How AI is used** (only after rules fail, NIGP-bounded, confidence threshold, transparent audit trail)
- How confidence and review flags work
- Key assumptions, limitations, and data quality issues
- Recommendations for governance and future maintenance

Write in clear, professional language suitable for City leadership.

## 8. Source documentation with verifiable links

For any external references used, provide:
- Source title
- Short description of what it informs
- **Direct, verifiable URL** (no fabricated links)

Organize at the end of the methodology explanation so leadership can independently review and vet the approach.

## 9. How to proceed

1. First, inspect the "EY raw data" file and summarize its structure and any issues.
2. Then propose the taxonomy and rule design for my approval.
3. After that, generate the Python pipeline code and show how to run it on "EY raw data".
4. Finally, produce the leadership-ready methodology explanation and the source list.

---

## Decisions locked during planning conversation (2026-04-29)

| # | Decision | Rationale |
|---|---|---|
| 1 | Three-level taxonomy on every row: Business Category → NIGP Class → NIGP Item (Option C) | Serves both executive dashboards (Business Category) and sourcing/audit needs (NIGP codes) |
| 2 | AI assist allowed on the long tail, with strict guardrails | Rules-only would produce a 20-40% review pile; bounded AI shrinks it while remaining defensible |
| 3 | Existing EY-supplied NIGP codes are starting hypotheses, NOT ground truth | EY's deliverable, not Chicago's authoritative system; must be validated against vendor + description |
| 4 | Lean output (~13 columns) instead of preserving all 87 | Raw file is preserved separately; downstream analysis doesn't need 87 columns and Excel chokes on them |
| 5 | EY data is raw material; the taxonomy + classifier are the actual deliverables | Categorized EY file is proof-of-concept, not the end product |
| 6 | Single source: `spend-analysis/data/raw/ey raw data.xlsx` only — no cross-dataset reconciliation in scope | Keep scope tight; other extracts can be re-run through the classifier later |
