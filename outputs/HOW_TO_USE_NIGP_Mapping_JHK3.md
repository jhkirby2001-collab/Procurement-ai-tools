# How to Use the NIGP Mapping File

**Prepared by:** James H. Kirby III, CSCP, MS-SCM
**File version:** 2026-04-30
**Project:** NIGP-Sourced Procurement Category Mapper

---

## What this file is

`NIGP_Mapping_JHK3.csv` contains every purchase order and invoice line from a 23-year public-sector procurement dataset (October 2002 through May 2025) — **784,556 transactions in total** — with a Business Category and NIGP code assigned to each one.

This is an internally-owned commodity classification of historical procurement spend. It is fully auditable and independent of any external consultant labels.

---

## What's in each column

| Column | What it tells you |
|---|---|
| `Transaction_ID` | Unique row identifier |
| `Date` | PO Creation Date |
| `Department` | Lead department code |
| `Vendor` | Vendor name (for readability only — not used to classify) |
| `Description_Best` | The transaction description the classifier used |
| `Amount` | PO dollar amount matched to AP |
| **`Business_Category`** | **The 17-bucket commodity rollup (the main answer)** |
| `NIGP_Class_3digit` | Standard NIGP 3-digit Class |
| `NIGP_Item_5digit` | Standard NIGP 5-digit Class-Item (when specific enough) |
| `NIGP_Code_Assigned` | The NIGP code we landed on (5-digit if available, else 3-digit) |
| `NIGP_Match_Level` | `exact` (high specificity) / `broad` (3-digit class only) / `review` (low confidence) |
| `Classification_Confidence` | High / Medium / Low |
| `Confidence_Score` | 0.0 to 1.0 numeric confidence |
| `Classification_Method` | `keyword_rule` / `account_pattern` / `human_review` |
| `Review_Flag` | `Yes` if the row needs human eyes; `No` if confidently classified |
| `Classification_Reason` | The exact rule that fired — full audit trail |

---

## The 17 Business Categories

| # | Category |
|---|---|
| 1 | Facilities Operations & Maintenance |
| 2 | Public Safety, Uniforms & PPE |
| 3 | Construction Materials |
| 4 | Vehicles & Fleet |
| 5 | Professional & Administrative Services |
| 6 | Office, Print & Marketing |
| 7 | Equipment Rental & Leasing |
| 8 | Janitorial, Sanitation & Waste |
| 9 | Heavy Equipment & Machinery |
| 10 | IT, Telecom & Audio/Visual |
| 11 | Landscaping, Grounds & Irrigation |
| 12 | Chemicals & Water Treatment |
| 13 | Medical & Health Services |
| 14 | Animal Care & Veterinary |
| 15 | Furniture & Furnishings |
| 16 | Construction & Trades Services |
| 17 | Grants & Pass-Through Funding |

The full mapping of which NIGP 3-digit Classes roll up to each Business Category is in `business_categories_JHK3.csv` (138 rows with judgment notes).

---

## How to use the file — three workflows

### Workflow A — "How was this categorized before?" (open in Excel, search)

The fastest answer for any description that's already happened in the last 23 years.

1. Open `NIGP_Mapping_JHK3.csv` in Excel. (File is ~225 MB and ~785K rows — give Excel 30-60 seconds to load.)
2. Click anywhere in row 1, then **Data → Filter**.
3. Click the filter arrow in the **Description_Best** column.
4. In the filter search box, type a keyword from your description (e.g., "GENERATOR" or "BULK SALT").
5. Excel shows only rows matching that keyword. Look at the **Business_Category** column — that tells you how similar items have been categorized.

This works for the majority of common purchases because the file already has 784K examples spanning 23 years.

### Workflow B — "Classify this brand-new description" (run the classifier)

For descriptions that haven't appeared before, the classifier itself can categorize them. This requires access to the project's GitHub Codespace.

1. Open the project codespace.
2. Open a terminal (`Terminal → New Terminal`).
3. Type this command, replacing the description with whatever you want to classify:

   ```bash
   python /workspaces/Procurement-ai-tools/spend-analysis/scripts/classifier_JHK3.py --describe "PASTE THE DESCRIPTION HERE"
   ```

4. The terminal prints back the Business Category, NIGP code, confidence, and the rule that fired.

If you don't have access to the codespace, send the description to James Kirby and he'll run it for you.

### Workflow C — "Process a whole new spend extract" (re-run the classifier on a new file)

When a new procurement extract is produced (next quarter, next year):

1. Save the new file to `spend-analysis/data/raw/`.
2. Update the `SRC_PARQUET` line at the top of `classifier_JHK3.py` to point at the new file.
3. Open a terminal and run:

   ```bash
   python /workspaces/Procurement-ai-tools/spend-analysis/scripts/classifier_JHK3.py --batch
   ```

4. Wait ~15 minutes. New `NIGP_Mapping_JHK3.csv` and `NIGP_Mapping_Review_Queue_JHK3.csv` files appear in `outputs/`.

---

## How to fix a misclassification permanently

If you find a description that was categorized wrong, you can teach the classifier the correct answer once and it will get every future similar description right.

1. Open `spend-analysis/data/reference/keyword_rules_DRAFT_JHK3.csv` in Excel.
2. Add a new row at the bottom with these columns:
   - `pattern` — the description text to match (e.g., `ICE MELT`)
   - `match_type` — `exact` for full string, `starts_with` for prefix, `contains` for substring
   - `business_category` — the correct Business Category (must match one of the 17 exactly)
   - `nigp_class_3digit` — the correct NIGP 3-digit class
   - `nigp_item_5digit` — leave blank if not specific enough
   - `nigp_match_level` — `broad` for normal cases, `exact` for very specific full-string rules, `review` if low confidence
   - `notes` — short justification (date, your initials, why)
3. Save the file.
4. Re-run the classifier in batch mode (Workflow C). The new rule applies to every row.

---

## When a row needs human eyes

If a row's `Review_Flag` is `Yes`, the classifier wasn't confident enough to commit. The companion file `NIGP_Mapping_Review_Queue_JHK3.csv` contains all 139,868 such rows. These are typically:

- Descriptions like "Misc supplies" or "Per contract" that don't say what was bought
- New commodity types that no rule covers yet
- Rows where the classifier matched a generic rule but the match is ambiguous

Procurement staff should triage this file periodically. Each row a staff member resolves can be converted into a new keyword rule (see "How to fix a misclassification permanently" above), shrinking the review queue over time.

---

## Files that travel with this deliverable

| File | What it is |
|---|---|
| `NIGP_Mapping_JHK3.csv` | The full classified dataset — the deliverable |
| `NIGP_Mapping_Review_Queue_JHK3.csv` | The 139,868 rows needing human review |
| `HOW_TO_USE_NIGP_Mapping_JHK3.md` | This file |
| `METHODOLOGY_JHK3.md` (in `spend-analysis/`) | Full methodology, audit narrative, locked decisions |
| `business_categories_JHK3.csv` (in `spend-analysis/data/reference/`) | The 138-row NIGP-class to Business-Category mapping with judgment notes |
| `business_categories_summary_JHK3.csv` (same folder) | The 17 Business Categories at a glance |

---

## Headline performance numbers

- **86.4%** of the 784,556 historical rows auto-classified by deterministic rule
- **17.8%** flagged for human review (Review_Flag = Yes)
- **148** hand-curated rules + **6,766** AI-mined rules + **6** account-code patterns
- **No AI is called at runtime.** The classifier is fully deterministic. AI was used once during build to mine long-tail patterns, then frozen.
- **No vendor lock-in.** All rule files are CSVs that procurement staff can edit directly.

---

## Questions

Direct to James H. Kirby III (CSCP, MS-SCM), project author. Full methodology and audit trail are in `METHODOLOGY_JHK3.md`.
