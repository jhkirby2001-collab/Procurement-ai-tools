# CLAUDE.md — Spend Analysis Project

## About This Project

Procurement spend analysis tools that categorize raw transaction data, identify vendor consolidation opportunities, and generate executive deliverables (Word reports, Excel workbooks, PowerPoint presentations) for senior leadership and CPO-level audiences.

## Project Structure

```
spend-analysis/
├── data/
│   ├── raw/          # Original source files — never modified (CSV, Excel exports from ERP/P2P systems)
│   └── processed/    # Cleaned and enriched data files produced by scripts
├── scripts/          # Python analysis scripts — one script per task or deliverable
├── outputs/          # All generated deliverables: Excel workbooks, Word docs, charts, reports
└── docs/             # Reference documents, instructions, and methodology notes
```

## Rules

- **Output location:** All generated files go to `/outputs`. Never write to `data/raw`.
- **Commodity codes:** Use NIGP (National Institute of Governmental Purchasing) codes for spend categorization. Reference 5-digit class-item codes where possible.
- **Currency format:** Always display dollar amounts in USD with comma separators and two decimal places (e.g., $1,234,567.89).
- **Reusable scripts:** Write scripts so they work with any similarly structured CSV input — avoid hardcoding file names or column names where possible. Use parameters or config variables at the top of the script.
- **Code explanations:** Explain what code does in plain language. Avoid jargon. Write comments as if explaining to someone who understands procurement deeply but not Python.

## About the User

Senior procurement professional with 28 years of experience in public sector purchasing, strategic sourcing, and vendor management.

- **Show the command first, then explain** what it does and why.
- Assume strong domain knowledge in procurement — no need to define terms like RFP, P-card, NIGP, or consolidation.
- Keep explanations focused on outcomes and decisions, not implementation details.
- When presenting analysis results, lead with the business insight, then show the supporting data.
