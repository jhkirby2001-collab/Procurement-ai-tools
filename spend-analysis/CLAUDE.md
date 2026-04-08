# CLAUDE.md — Procurement Intelligence System
**Version:** 4.1
**Updated:** 2026-04-08
**Change Log:**
- v4.1 (2026-04-08): Rewritten throughout to reflect leadership identity and executive framing. Updated "About" section, project description, attribution language, and portability framing. All deliverable titles and tone aligned to supply chain executive positioning.
- v4.0 (2026-04-08): Added Executive Summary narrative template (Section 4.4), Consolidation Savings Model tab template (Section 3.6), savings assumption documentation (Section 3.7), escalation/tone guidance (Section 6), portability rules (Section 10), chart gating rule (Section 5), version log. Removed hardcoded jurisdiction-specific language.
- v3.1: Absorbed execution logic from Claude.ai project instructions. Expanded from ~300 to ~2,400 words.
- v3.0: Original full instruction set.

---

## About This System

This is a portable AI-powered procurement intelligence capability built by James H. Kirby III — supply chain and procurement leader with 28 years of experience transforming how public sector organizations move from raw transaction data to executive-level decisions.

The system categorizes spend, identifies consolidation opportunities, models savings, and generates boardroom-ready deliverables for CPO and C-suite audiences. The methodology, tools, and analytical framework are the asset. Only the data changes between organizations.

**Current deployment:** Public sector procurement environment.
**Deliverable attribution:** James H. Kirby III, CSCP, MS-SCM | Senior Buyer and Procurement Research Analyst *(current official title — used on external deliverables only)*

---

## Project Structure

```
spend-analysis/
├── data/
│   ├── raw/        # Original source files — NEVER modify these
│   └── processed/  # Cleaned and enriched data files produced by scripts
├── scripts/        # Python analysis scripts — one script per task
├── outputs/        # All generated deliverables: Excel, Word, reports
└── docs/           # Reference documents, instructions, methodology notes
```

---

## Core Rules

- **Output location:** All generated files go to `/outputs`. Never write to `data/raw`.
- **Commodity codes:** Use NIGP (National Institute of Governmental Purchasing) 5-digit class-item codes for spend categorization.
- **Currency format:** USD with comma separators and two decimal places ($1,234,567.89).
- **Reusable scripts:** Write scripts that work with any similarly structured CSV input. Use parameters or config variables at the top — avoid hardcoding file names, column names, or jurisdiction-specific values.
- **Code comments:** Write comments as if explaining to someone who understands procurement deeply but not Python.
- **Show command first, then explain** what it does and why.
- **Lead with the business insight**, then show the supporting data.
- **Charts:** Do NOT produce any charts or visualizations unless explicitly instructed to do so in the current session. See Section 5.

---

## SECTION 0: PROFESSIONAL ATTRIBUTION

ALL deliverables must credit James H. Kirby III using the current official title. This title appears on external deliverables only and reflects the current organizational role — it does not define the scope or ambition of this work.

### File Naming Convention (MANDATORY)

All output files must include initials in the filename:
- Word Report: `[Jurisdiction]_Spend_Analysis_JHK3_[YYYY-MM-DD].docx`
- Excel Workbook: `[Jurisdiction]_Spend_Analysis_Data_JHK3_[YYYY-MM-DD].xlsx`
- Charts (when authorized): `chartN_description_JHK3.png`

### Word Report Attribution

**Title Page:** Include after the report title and jurisdiction:
```
Prepared by: James H. Kirby III, CSCP, MS-SCM
Senior Buyer and Procurement Research Analyst
Analysis Date: [Current Date]
```

**Footer on All Pages:**
- Left: "Prepared by James H. Kirby III"
- Center: [Section name or blank]
- Right: "Page X of Y"
- Font: Arial, 9pt, gray (#666666)

### Excel Workbook Attribution

**Executive Summary Sheet Header (Rows 1-3):**
- Row 1: "[JURISDICTION] — [DEPARTMENT NAME]" — Bold, white text, navy background (#003366), merged across all columns
- Row 2: "Prepared by: James H. Kirby III, CSCP, MS-SCM  |  Senior Buyer and Procurement Research Analyst" — Arial 11pt
- Row 3: "Analysis Date: [Current Date]"
- Row 4: [Blank]

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

Generate ALL of the following for every analysis unless the data does not support a specific analysis — flag and explain why if skipped.

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
Analyze by vendor location (local/regional/state/national). Highlight local vendor share and categories where local participation is high or low. Distinguish tail spend from core/strategic spend when useful.

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

### 3.6 Consolidation Savings Model Tab (MANDATORY FORMAT)

When producing a Consolidation Savings Model Excel tab, follow this exact template and structure. This is the standard format — do not deviate unless explicitly instructed.

#### Tab-Level Header (Rows 1-3)
- Row 1: "[JURISDICTION] — CONSOLIDATION SAVINGS MODEL — VERIFIED QUICK WIN OPPORTUNITIES" — Bold, white text, navy background (#003366), merged across all columns
- Row 2: "[Organization] | [Department]" — same navy background, white text, regular weight
- Row 3: Blank

#### Methodology Banner (Row 4)
- Row 4: "CALCULATION:   5-Year Spend  ÷  5  =  Annual Run Rate  ×  85%  =  Addressable Spend  ×  10%  =  Annual Savings  ×  3  =  3-Year Projected Value" — Bold, light gray background (#F2F2F2), full width

#### Per-Category Block Structure
Repeat the following block for each category, separated by one blank row:

**Row A — Category Header:**
- Column A: Category number (1, 2, 3...)
- Column B: Category name (e.g., "Fencing & Guardrail")
- Column G: Label "5-Year Total:"
- Column H: Total 5-year spend value (hardcoded from source data)
- Background: Medium blue (#0071BC), white bold text, full row

**Row B — Solicitation Status:**
- Column B: "Solicitation Status: [plain language description of where the procurement stands — e.g., Phase A — Specification Development. No solicitation issued.]"
- Background: Light blue (#E6EEF5), italic text

**Rows C-D — Vendor Sub-Table:**
- Row C header — Columns B–E: "VENDORS" | "5-Year Spend" | "Transactions" | "Depts Served" — Light blue (#E6EEF5), bold
- Data rows: One row per vendor, plain white background, no bold
- Vendor Subtotal row: "Vendor Subtotal" | =SUM(spend range) — Light gray (#F2F2F2), bold

**Rows E-F — Department Sub-Table:**
- Row E header — Columns B–C: "DEPARTMENTS PURCHASING" | "5-Year Spend" — Light blue (#E6EEF5), bold
- Data rows: One row per department, plain white background
- Department Subtotal row: "Department Subtotal" | =SUM(dept range) — Light gray (#F2F2F2), bold

**Savings Chain Row (immediately after Department Subtotal):**
Single row, light gray (#F2F2F2), bold. All inline across columns B–I:
- B: "Annual (÷5)" | C: =CategoryTotalCell/5
- D: "Addressable (×85%)" | E: =AnnualCell*0.85
- F: "Savings (10%)" | G: =AddressableCell*0.1
- H: "3-Year Value" | I: =AnnualSavingsCell*3

#### Combined Total Row (after all categories, one blank row gap)
- B: "COMBINED TOTAL" | C: =SUM of all category 5-year totals | D: "Annual:" | E: =SUM of all annual cells | F: "Ann Savings:" | G: =SUM of all annual savings cells | H: "3-Year:" | I: =SUM of all 3-year value cells
- Background: Navy (#003366), white bold text

#### Footer Rows
- Sources row: "Sources: [Data source description, date range, system name]" — plain text, no background
- Attribution row: "Prepared by James H. Kirby III, CSCP, MS-SCM  |  Senior Buyer and Procurement Research Analyst  |  [Jurisdiction]" — plain text, no background

---

### 3.7 Savings Assumption Documentation

The consolidation savings model uses the following assumptions. These are conservative, defensible estimates appropriate for public sector procurement. Use these defaults unless James explicitly overrides them in the current session.

**85% Addressable Spend:**
Assumes 15% of spend in any category is already optimized, under existing contract, or otherwise not actionable. This is a standard public sector analytical convention and errs on the side of caution.
- *Defense language if challenged:* "We applied an 85% addressability factor to account for transactions already under competitive contract or subject to regulatory constraints. This is a conservative assumption — actual addressable spend may be higher."

**10% Savings Rate:**
A moderate, conservative savings rate applied to addressable spend. Industry benchmarks for consolidation in public sector procurement typically range from 8%–15%. Using 10% keeps projections credible and defensible in front of budget and finance leadership.
- *Defense language if challenged:* "The 10% savings rate is based on industry benchmarks for vendor consolidation in comparable public sector environments. It is intentionally conservative. Actual savings may exceed this depending on market conditions and negotiation outcomes."

**3-Year Projection Window:**
Represents a standard contract term in public sector procurement. Long enough to be meaningful to leadership, short enough to be credible. Does not assume option years.

**When to adjust assumptions:**
- Increase savings rate to 12–15% if the category has zero existing contracts and high vendor fragmentation
- Decrease to 7–8% if the category has partial contract coverage or complex specifications
- Always note any assumption changes in the Sources/Notes row of the model

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
- Total Spend, Total Contract Value, Payment Amount, Average Transaction Size, Average Contract Value, Estimated Savings Potential

**Whole Number (0)** — For COUNT metrics. NO DOLLAR SIGNS:
- Transaction Count, Contract Count, Vendor Count, Department Count, Purchase Order Count

**Decimal (#,##0.00)** — For SCORES/RATIOS/INDICES. NO DOLLAR SIGNS:
- HHI Score, Fragmentation Score, Concentration Ratio

**Percentage (0.00%)** — ONLY for proportions:
- Percentage of Total Spend, Vendor Share, Utilization Rate

**CRITICAL ERRORS TO AVOID:**
- ✗ NEVER put $ on count fields (Vendor Count = 52, not $52.00)
- ✗ NEVER put $ on score fields (HHI = 1,234.56, not $1,234.56)
- ✗ NEVER apply % unless the value is already a decimal proportion

### 4.3 Professional Styling — Standard Data Sheets

- Header row (Row 1): Height 30, Bold, navy background (#003366), white text, centered, wrapped
- Formula row (Row 2): Height 25, Italic, gray background (#E7E6E6), left-aligned, wrapped
- All cells: Thin borders
- Dollar amounts: $#,##0.00
- Percentages: 0.00%

### 4.3a Color Palette (Use Consistently Across ALL Sheets and ALL Deliverables)

| Color | Hex | Use |
|---|---|---|
| Navy | #003366 | Primary section headers, title rows, combined total rows |
| Medium Blue | #0071BC | Category headers, sub-table headers in savings model |
| Light Blue | #E6EEF5 | Sub-headers (VENDORS, DEPARTMENTS) within category blocks |
| Light Gray | #F2F2F2 | Summary rows, subtotals, methodology banners, AT A GLANCE sections |
| Light Green | #D9EAD3 | Final totals, 3-year value highlights |
| Red | #E31837 | "THE PROBLEM" header on Executive Summary only |
| White | #FFFFFF | Standard data rows |

### 4.4 Executive Summary Sheet — Narrative Template (MANDATORY — FIRST SHEET)

The Executive Summary is not a data dump. It is a leadership brief. It must tell a story in five sections in this exact order. Every section header, color, and sequence is mandatory.

**HEADER (Rows 1-3):** Attribution block per Section 0.

**Row 4 — Blank**

**Section: THE PROBLEM**
- Header cell: "THE PROBLEM" — Bold, white text, red background (#E31837)
- Row below: One sentence. Plain language. States what is wrong with the current state. No jargon.
- Example: "5 vendors are selling the same commodity to 9 departments with no consolidated contract and no competitive solicitation on record."
- Background: white, no bold

**One blank row**

**Section: AT A GLANCE**
- Header cell: "AT A GLANCE" — Bold, light gray background (#F2F2F2)
- Column headers row: Total 5-Year Spend | Vendors | Department Instances | Contract Coverage | Annual Savings (10%) | 3-Year Projected Value — light gray background (#F2F2F2), bold
- Data row: Actual values from the analysis — light gray background (#F2F2F2)

**One blank row**

**Section: THE OPPORTUNITIES**
- Header: "THE OPPORTUNITIES" — Bold, white text, navy background (#003366)
- Column headers: # | Category | 5-Year Spend | Vendors | Departments | Annual Savings | 3-Year Savings | Solicitation Status — Medium blue (#0071BC), white text, bold
- One data row per category: alternating white and light gray (#F2F2F2) rows
- Total row: "TOTAL" label, SUM formulas for all dollar columns — Navy background (#003366), white bold text

**One blank row**

**Section: WHY THESE CATEGORIES?**
- Header: "WHY THESE CATEGORIES?" — Bold, light gray background (#F2F2F2)
- Bullet points — Column A: "•" | Column B: explanation text, one bullet per row:
  - No existing solicitations issued and no active competitive contract
  - Specifications are in early development — the window is open to shape scope
  - Commodity items with interchangeable vendors — consolidation is low-risk
  - Multiple departments buying independently — no coordination, no leverage
  - Zero political risk — no department loses a preferred vendor; the market opens to competition
  - [Add any data-specific bullets that apply to the actual findings]

**One blank row**

**Section: RECOMMENDED NEXT STEP**
- Header: "RECOMMENDED NEXT STEP" — Bold, white text, navy background (#003366)
- Row below: Single declarative sentence. One action. Who does what by when.
- Example: "Authorize the Department of Procurement Services to develop a consolidated citywide solicitation for [Category Name], targeting award within [timeframe]."

**FOOTER (last two rows):**
- "Data: [Source name and date range]  |  Analysis performed [Month Year]"
- "Prepared by James H. Kirby III, CSCP, MS-SCM  |  Senior Buyer and Procurement Research Analyst  |  [Jurisdiction]"
- Both rows: plain text, no background, Arial 9pt

---

## SECTION 5: CHARTS AND VISUALIZATIONS

### ⛔ DEFAULT RULE: DO NOT PRODUCE CHARTS

Claude Code will NOT generate any charts, graphs, or visualizations unless James explicitly says to produce them in the current session.

This is a hard default. Do not produce charts as part of a "complete analysis." Do not produce charts because Section 7 or the completion checklist references them. Charts are a separate, explicitly triggered deliverable.

**Trigger phrase required.** James must say something like "generate the charts," "produce chart 3," or "run the visualizations" before any chart is created.

### 5.1 Chart Types Available (When Authorized)

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

### 5.2 Additional Charts (When Authorized and Data Available)
- Geographic spend by local/regional/state/national
- Time trends (monthly/quarterly)
- Spend by payment method

### 5.3 Chart Standards (When Producing Charts)
- Save each as PNG: `chartN_description_JHK3.png`
- 300 DPI resolution
- Clear axis labels, titles, data labels on bars
- Consistent color scheme: navy (#003366) primary, medium blue (#0071BC) secondary
- Make each downloadable via present_files tool
- Attribution: Small text bottom right — "Analysis by J.H. Kirby III | [Date]" — Arial 8pt, gray

---

## SECTION 6: ESCALATION AND TONE GUIDANCE

All findings, recommendations, and narrative text must be written for a CPO, Deputy Commissioner, or C-suite leader who has 5 minutes to read. This system produces leadership intelligence, not analytical reports.

### When to Use "Risk" Language
Use "risk" explicitly when:
- Spend is occurring with no contract on record
- A single vendor controls >50% of spend in a critical category
- Payment method suggests off-contract purchasing (P-card for large transactions)
- Year-over-year spend is increasing in a category with no competitive solicitation

### When to Use "Opportunity" Language
Use "opportunity" when:
- Consolidation is possible but no immediate harm is occurring
- Multiple vendors exist in a category with no price benchmarking
- Administrative burden could be reduced through contract restructuring

### When to Flag for Immediate Leadership Attention
Prepend findings with "⚠ FLAG FOR LEADERSHIP:" when:
- Any vendor exceeds 80% of a single category's spend with no contract
- A department is purchasing goods or services outside any approved contract vehicle
- Duplicate payments or negative amounts appear in the data

### Tone Standard
- Every finding must answer: "So what? What should leadership do about this?"
- Use declarative sentences: "The department is paying 3 vendors for the same service." Not: "It appears that spend may be distributed across multiple vendors."
- Avoid hedging language unless the data is genuinely uncertain
- When data is uncertain, say so plainly: "Transaction descriptions are inconsistent — this finding should be verified against source purchase orders before presenting to leadership."

---

## SECTION 7: DELIVERABLES CHECKLIST

At the end of EVERY analysis, produce and make downloadable:

### Excel Workbook
`[Jurisdiction]_Spend_Analysis_Data_JHK3_[YYYY-MM-DD].xlsx`

Sheets in this order:
1. Executive Summary (first sheet, mandatory — narrative template per Section 4.4)
2. Vendor Summary
3. Category Summary
4. Department Summary
5. Contract Type Summary
6. Geographic Distribution
7. Cross-Department Vendors
8. Consolidation Savings Model (per Section 3.6 template)
9. Multi-Purchase Order Vendors
10. Multi-Vendor Items (if applicable)
11. Cleaned Data (up to 50,000 rows transaction / 1,000 rows contract)

### Word Report
`[Jurisdiction]_Spend_Analysis_JHK3_[YYYY-MM-DD].docx`

Sections:
- Title page with attribution
- Executive Summary (3-5 paragraphs — lead with business insight, not methodology)
- Key Findings (declarative statements, not observations)
- Consolidation Opportunities (reference savings model assumptions per Section 3.7)
- Methodology section with calculated fields table
- Data limitations and assumptions
- Glossary of terms
- Footer on all pages

**Note:** Claude Code generates the structural framework and data-driven content. Claude.ai (Leadership Intelligence project) refines the executive narrative and leadership framing.

### Chart Images (ONLY IF AUTHORIZED THIS SESSION)
All requested charts as individual PNGs per Section 5.3.

### Deliverables Summary Table
Display at end of analysis:

| File Name | Type | Description | Status |
|-----------|------|-------------|--------|
| (each file) | .docx / .xlsx / .png | (brief description) | ☑ Downloadable |

---

## SECTION 8: INITIAL QUESTIONS

ALWAYS ask at the start of any analysis:

1. Which jurisdiction or organization is this data for?
2. What time period does this data cover?
3. Do you have custom category codes, NIGP rollups, or department mappings?
4. Any specific areas of focus or questions leadership has raised?
5. Should charts be produced as part of this analysis? (Default: NO — charts require explicit authorization)

---

## SECTION 9: COMPLETION CHECKLIST

Before concluding ANY analysis, verify and display ALL items:

**Attribution:** ☐ Word title page ☐ Word footer ☐ Excel header rows 1-3 ☐ _JHK3 in all filenames

**Data Prep:** ☐ Column summary ☐ Quality issues flagged ☐ Schema confirmed ☐ Data dictionary

**Core Analyses:** ☐ Vendor summary with 80/20 ☐ Category summary ☐ Department summary ☐ Contract type ☐ Transaction size or geographic distribution

**Consolidation:** ☐ Cross-department vendors ☐ Category fragmentation ☐ Same item/multiple vendor (if applicable) ☐ Multi-PO vendors ☐ Vendor-department matrix ☐ Consolidation Savings Model tab (Section 3.6 format)

**Charts:** ☐ Were charts authorized this session? If NO — skip entirely. If YES — all authorized charts saved as PNG and downloadable.

**Excel:** ☐ Executive Summary is first sheet ☐ Executive Summary follows narrative template (Section 4.4) ☐ All sheets in correct order ☐ Full column headers (no abbreviations) ☐ Dollar amounts = $#,##0.00 ☐ Counts = whole numbers, no $ ☐ Scores = decimals, no $ ☐ Formula rows in Row 2 ☐ Frozen panes at Row 3 ☐ Color palette consistent per Section 4.3a

**Word Report:** ☐ All sections present ☐ Declarative findings language ☐ Methodology table ☐ Savings assumption documentation referenced ☐ Glossary ☐ Footer on all pages

**Escalation:** ☐ Any ⚠ FLAG FOR LEADERSHIP items identified and surfaced?

**Final:** ☐ All files downloadable ☐ Deliverables summary table displayed

---

## SECTION 10: PORTABILITY RULES

This system was built to travel. The methodology, scripts, templates, and analytical framework belong to James H. Kirby III and move with him to any organization. Only the data changes. This is the core strategic value of the system.

**Rules for maintaining portability:**

1. **Never hardcode jurisdiction names** in scripts. Use a variable at the top — changeable in one place.

2. **Never hardcode department names** in logic or filtering. Accept them as inputs from the data.

3. **Never reference specific contract numbers, requisition numbers, or internal system IDs** in reusable scripts. These belong only in output reports, populated from the data.

4. **Parameter block at the top of every script** — required fields:
   ```python
   # ── CONFIGURATION ──────────────────────────────────────────
   JURISDICTION    = "City of Chicago"           # Change for each organization
   ANALYSIS_PERIOD = "2021–2025"                 # Update to match data
   INPUT_FILE      = "data/raw/spend_data.csv"   # Update to match filename
   OUTPUT_DIR      = "outputs/"                  # Do not change
   ANALYST_NAME    = "James H. Kirby III, CSCP, MS-SCM"  # Do not change
   # ───────────────────────────────────────────────────────────
   ```

5. **NIGP codes are universal.** They do not change between organizations. Always use them as the backbone of category analysis.

6. **Savings assumptions (Section 3.7) are methodology, not local knowledge.** They apply at any organization and should never be modified to reflect a specific organization's internal targets without explicit documentation and justification.

---

## About James H. Kirby III

Supply chain and procurement leader with 28 years of experience driving strategic sourcing, vendor management, and procurement transformation in complex public sector environments. CSCP certified. MS in Supply Chain Management.

This system represents a portable AI-powered procurement intelligence capability — built to transform how organizations move from raw transaction data to executive-level decisions. The methodology, analytical framework, and tools are the asset. Only the data changes between organizations.

- Assume strong domain knowledge — no need to define RFP, P-card, NIGP, consolidation, HHI, tail spend, maverick spend, etc.
- Lead with business insight, then supporting data.
- Frame every output for a CPO, Deputy Commissioner, or C-suite audience.
- This is a leadership intelligence tool, not a reporting tool.
- Every deliverable should reflect the thinking of a senior supply chain executive, not a data processor.