
# CLAUDE.md — Procurement Spend Analysis System

## About This Project

Procurement spend analysis tools that categorize raw transaction data, identify vendor consolidation opportunities, and generate executive deliverables (Word reports, Excel workbooks, charts) for senior leadership and CPO-level audiences.

Analyst: James H. Kirby III, Senior Buyer and Procurement Research Analyst.

## Project Structure

```
spend-analysis/
├── data/
│   ├── raw/          # Original source files — NEVER modify these
│   └── processed/    # Cleaned and enriched data files produced by scripts
├── scripts/          # Python analysis scripts — one script per task
├── outputs/          # All generated deliverables: Excel, Word, charts, reports
└── docs/             # Reference documents, instructions, methodology notes
```

## Core Rules

- **Output location:** All generated files go to `/outputs`. Never write to `data/raw`.
- **Commodity codes:** Use NIGP (National Institute of Governmental Purchasing) 5-digit class-item codes for spend categorization.
- **Currency format:** USD with comma separators and two decimal places ($1,234,567.89).
- **Reusable scripts:** Write scripts that work with any similarly structured CSV input. Use parameters or config variables at the top — avoid hardcoding file names or column names.
- **Code comments:** Write comments as if explaining to someone who understands procurement deeply but not Python.
- **Show command first, then explain** what it does and why.
- **Lead with the business insight**, then show the supporting data.

---

## SECTION 0: PROFESSIONAL ATTRIBUTION

ALL deliverables must credit James H. Kirby III.

### File Naming Convention (MANDATORY)

All output files must include initials in the filename:
- Word Report: `[Jurisdiction]_Spend_Analysis_JHK3_[YYYY-MM-DD].docx`
- Excel Workbook: `[Jurisdiction]_Spend_Analysis_Data_JHK3_[YYYY-MM-DD].xlsx`
- Charts: `chartN_description_JHK3.png`

### Word Report Attribution

**Title Page:** Include after the report title and jurisdiction:
```
Prepared by: James H. Kirby III
Procurement Research Analyst
Analysis Date: [Current Date]
```

**Footer on All Pages:**
- Left: "Prepared by James H. Kirby III"
- Center: [Section name or blank]
- Right: "Page X of Y"
- Font: Arial, 9pt, gray (#666666)

### Excel Workbook Attribution

**Executive Summary Sheet Header (Rows 1-4):**
- Row 1: "PROCUREMENT SPEND ANALYSIS"
- Row 2: "Prepared by: James H. Kirby III, Procurement Research Analyst" (Bold, Arial 11pt, blue #366092)
- Row 3: "Analysis Date: [Current Date]"
- Row 4: [Blank]

### Chart Attribution (Optional)

Small text at bottom right: "Analysis by J.H. Kirby III | [Date]" — Arial, 8pt, gray. Only add if it doesn't clutter the visualization.

---

## SECTION 1: DATA UNDERSTANDING AND CLEANUP

When provided a file or data sample:

1. Identify and describe each column (vendor name, vendor ID, description, amount, date, department, cost center, NIGP code, contract number, payment method, address, city, state, ZIP)
2. Distinguish structured vs. unstructured fields (especially free-text descriptions, vendor names, memo/notes)
3. Flag data quality issues: missing values, inconsistent vendor names, negative amounts, unusual dates, duplicates
4. Propose a cleaned "analysis-ready" schema and confirm with the user before deeper analysis

### Cleaning Standards

- **Vendor name normalization:** Group obvious variants under one standardized name (e.g., "ABC INC.", "A.B.C. Incorporated", "ABC, Inc" → "ABC Inc"). Use vendor IDs and addresses where available.
- **Description normalization:** Use descriptions plus vendor/context to infer what was purchased. Normalize spelling, remove noise, infer/refine NIGP codes when possible.
- **Department standardization:** Standardize department and cost center names. Flag unmappable records.
- **Geography tagging:** Tag vendor geography (local/regional/state/national) based on jurisdiction rules, using address fields. Explain limitations.
- **Data dictionary:** Produce a summary of final fields used for analysis.

---

## SECTION 2: CORE SPEND ANALYSES (ALL REQUIRED)

Generate ALL of the following for every analysis:

### 2.1 Spend by Vendor Summary
Total spend, transaction count, average payment size per vendor. Highlight top vendors and 80/20 patterns.

### 2.2 Spend by Category / Commodity
Group by NIGP code or higher-level taxonomy. Total spend, transaction count, share of total per category. Call out fragmented vs. concentrated categories.

### 2.3 Spend by Department / Cost Center
Break down by department. Identify top-spending departments and how their spend distributes across categories.

### 2.4 Spend by Contract Type
Group by contract vehicle (blanket PO, one-time, requirements, master agreement). Total spend, transaction count, average transaction size. Flag heavy one-time/off-contract spend.

### 2.5 Spend by Payment Method (When Data Available)
Group by payment type (P-card, check, ACH, wire). Total spend, transaction count, average size. Note optimization opportunities.

### 2.6 Geographic Spend Analysis
Analyze by vendor location (local/regional/state/national). Highlight local vendor share and categories where local participation is high or low.

Also distinguish tail spend from core/strategic spend when useful.

---

## SECTION 3: CONSOLIDATION OPPORTUNITY ANALYSES (ALL REQUIRED)

### 3.1 Cross-Department Vendor Analysis
Vendors serving 3+ departments. Total spend, department count, PO count. Highlight enterprise-wide agreement opportunities.

### 3.2 Category Fragmentation Analysis
Per category: vendor count, average spend per vendor. Flag categories with high vendor counts relative to spend.

### 3.3 Same Item / Multiple Vendor Analysis (When Data Available)
Items purchased from 2+ vendors. Total spend, vendor count, transaction count. Prioritize items with ≥$100K spend.

### 3.4 Contract Consolidation Analysis
Vendors with 5+ POs (transaction data) or 3+ contracts (contract data). Unique PO count, total spend, average transaction size. Flag where master agreements could reduce administrative burden.

### 3.5 Vendor-Department Spend Matrix
Cross-reference: which vendors serve which departments and spend amounts. Identify leverage opportunities from independent contracting with the same supplier.

---

## SECTION 4: EXCEL OUTPUT FORMATTING STANDARDS

### 4.1 Column Headers (MANDATORY — NO EXCEPTIONS)

1. **Full names only.** No abbreviations ever.
   - ✓ "Average Transaction Size" — ✗ "Avg Txn Size"
   - ✓ "Transaction Count" — ✗ "Txn Count"
   - ✓ "Purchase Order" — ✗ "PO"
   - ✓ "Department" — ✗ "Dept"

2. **Formula row in Row 2** beneath every header row. Shows calculation method in plain language.
   - NEVER start cell values with "=" (Excel interprets as formula → #NAME? errors)
   - Correct: "Formula: Total Spend ÷ Transaction Count"
   - Correct: "Sum of all contract values for this vendor"

3. **Freeze panes at Row 3** (after header and formula rows).

### 4.2 Number Formatting by Data Type (CRITICAL)

**Currency ($#,##0.00)** — ONLY for actual dollar amounts:
- Total Spend, Total Contract Value, Payment Amount
- Average Transaction Size, Average Contract Value
- Estimated Savings Potential

**Whole Number (0)** — For COUNT metrics. NO DOLLAR SIGNS:
- Transaction Count, Contract Count, Vendor Count
- Department Count, Purchase Order Count

**Decimal (#,##0.00)** — For SCORES/RATIOS/INDICES. NO DOLLAR SIGNS:
- HHI Score, Fragmentation Score, Concentration Ratio

**Percentage (0.00%)** — ONLY for proportions:
- Percentage of Total Spend, Vendor Share, Utilization Rate

**CRITICAL ERRORS TO AVOID:**
- ✗ NEVER put $ on count fields (Vendor Count = 52, not $52.00)
- ✗ NEVER put $ on score fields (HHI = 1,234.56, not $1,234.56)
- ✗ NEVER apply % unless the value is already a decimal proportion

### 4.3 Professional Styling

- Header row (Row 1): Height 30, Bold, blue background (#366092), white text, centered, wrapped
- Formula row (Row 2): Height 25, Italic, gray background (#E7E6E6), left-aligned, wrapped
- All cells: Thin borders
- Dollar amounts: $#,##0.00
- Percentages: 0.00%

### 4.4 Executive Summary Sheet (MANDATORY — FIRST SHEET)

The Excel workbook MUST include an Executive Summary as the first sheet.

**Structure:**
- Rows 1-4: Attribution header (Section 0)
- Section A — Dataset Overview: Jurisdiction, Analysis Period, Total Contracts, Total Value, Unique Vendors, Departments, Categories
- Section B — Top Performers: Top 3 Vendors, Top 3 Departments, Top 3 Categories (name, value, % of total)
- Section C — Key Findings: Vendor concentration ("X vendors control 80% of spend"), strategic vendor count, geographic spend breakdown
- Section D — Consolidation Opportunities: Cross-department vendor count and value, estimated savings (5%), multi-contract vendors, administrative savings potential

**Formatting:** Bold section headers, currency on dollar amounts, percentages on proportions, whole numbers on counts, borders between sections, 12pt section headers, one blank row between sections.

---

## SECTION 5: CHARTS AND VISUALIZATIONS (ALL REQUIRED)

### 5.1 Required Charts (ALL 10 for every analysis)

1. Top 10 Vendors by Total Spend (horizontal bar)
2. Spend by Category (horizontal bar)
3. Spend by Department (horizontal bar)
4. Spend by Contract Type (horizontal bar)
5. Vendor Concentration / Pareto Curve (line chart, 80/20)
6. Transaction Size Distribution (grouped bar) OR Geographic Distribution (for contract data)
7. Cross-Department Vendors (horizontal bar, leverage opportunities)
8. Category Fragmentation (bubble chart: vendor count vs. spend)
9. Multi-PO/Contract Vendors (horizontal bar, consolidation candidates)
10. Vendor-Department Spend Matrix (heatmap)

### 5.2 Additional Charts (when data available)

- Geographic spend by local/regional/state/national
- Time trends (monthly/quarterly)
- Spend by payment method

### 5.3 Chart Standards

- Save each as PNG: `chartN_description_JHK3.png`
- 300 DPI resolution
- Clear axis labels, titles, data labels on bars
- Consistent color scheme across all charts
- Make each downloadable via present_files tool

---

## SECTION 7: DELIVERABLES CHECKLIST

At the end of EVERY analysis, produce and make downloadable:

### Word Report
`[Jurisdiction]_Spend_Analysis_JHK3_[YYYY-MM-DD].docx`

Sections (Claude Code generates the structural framework; Claude.ai refines the narrative):
- Title page with attribution
- Executive Summary (3-5 paragraphs)
- Key Findings (bulleted)
- Consolidation Opportunities
- Visual Analytics Guide — for each chart include: Chart Name, What It Shows, How It Was Calculated, Why It Matters, Recommended Actions (5-6 items)
- Methodology section with calculated fields table
- Data limitations and assumptions
- Glossary of terms
- Footer on all pages

### Excel Workbook
`[Jurisdiction]_Spend_Analysis_Data_JHK3_[YYYY-MM-DD].xlsx`

Sheets in this order:
1. Executive Summary (first sheet, mandatory)
2. Vendor Summary
3. Category Summary
4. Department Summary
5. Contract Type Summary
6. Geographic Distribution
7. Cross-Department Vendors
8. Multi-Purchase Order Vendors
9. Multi-Vendor Items (if applicable)
10. Cleaned Data (up to 50,000 rows transaction / 1,000 rows contract)

### Chart Images
All 10 required charts as individual PNGs.

### Deliverables Summary Table

Display at end of analysis:

| File Name | Type | Description | Status |
|-----------|------|-------------|--------|
| (each file) | .docx/.xlsx/.png | (description) | ☑ Downloadable |

---

## SECTION 9: COMPLETION CHECKLIST

Before concluding ANY analysis, verify and display ALL items:

**Attribution:** ☐ Word title page ☐ Word footer ☐ Excel header ☐ _JHK3 in all filenames

**Data Prep:** ☐ Column summary ☐ Quality issues flagged ☐ Schema confirmed ☐ Data dictionary

**Core Analyses:** ☐ Vendor summary with 80/20 ☐ Category summary ☐ Department summary ☐ Contract type ☐ Transaction size or geographic distribution

**Consolidation:** ☐ Cross-department vendors ☐ Category fragmentation ☐ Same item/multiple vendor (if applicable) ☐ Multi-PO vendors ☐ Vendor-department matrix

**Charts:** ☐ All 10 required charts saved as PNG and downloadable

**Excel:** ☐ Executive Summary is first sheet ☐ All sheets in correct order ☐ Full column headers (no abbreviations) ☐ Dollar amounts = $#,##0.00 ☐ Counts = whole numbers, no $ ☐ Scores = decimals, no $ ☐ Formula rows in Row 2 ☐ Frozen panes at Row 3

**Word Report:** ☐ All sections present ☐ Visual Analytics Guide complete for all charts ☐ Methodology table ☐ Glossary ☐ Footer on all pages

**Final:** ☐ All files downloadable ☐ Deliverables summary table displayed

---

## SECTION 8: INITIAL QUESTIONS

ALWAYS ask at the start of any analysis:

1. Which jurisdiction is this data for?
2. What time period does this data cover?
3. Do you have custom category codes, NIGP rollups, or department mappings?
4. Any specific areas of focus or questions?

---

## About the User

Senior procurement professional with 28 years of experience in public sector purchasing, strategic sourcing, and vendor management.
- Assume strong domain knowledge — no need to define RFP, P-card, NIGP, consolidation, etc.
- Keep explanations focused on outcomes and decisions.
- Lead with business insight, then supporting data.
- Building a portable skillset — scripts should work at any organization, not just Chicago.