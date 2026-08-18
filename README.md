# Procurement AI Tools

**A working portfolio of public-sector procurement engineering — by a 28-year procurement practitioner who also architected the analytics and AI behind it.**

These are independent, AI-assisted tools that turn slow, manual procurement work into automated, auditable, reusable capability. The work is independent and not affiliated with or endorsed by any government agency.

---

## Flagship: NIGP-Sourced Procurement Category Mapper

A reusable, rules-based classification engine that maps unstructured procurement descriptions to an industry-standard commodity taxonomy — and proves *why* every decision was made.

**Validated on a 784,556-row public-sector procurement dataset (AP activity years 2017, 2020, 2021, 2023):**

| | |
|---|---|
| **Coverage** | **100% mapped — 0 rows in a human-review queue** |
| **Taxonomy** | 17 Business Categories (MECE) → NIGP 3-digit Class → NIGP 5-digit Item |
| **Runtime** | ~15 min batch + 5 sec resolver; **no API key, no internet dependency** |
| **Rule base** | 293 hand-curated + 6,766 AI-mined rules + 6 account-code patterns |

**Coverage by method**

- 87.7% — deterministic keyword rule
- 11.9% — AI-assist fallback (saved one-time AI output, no runtime API)
- 0.3% — account-code pattern (subgrant / pass-through disbursements)
- 0.1% — explicitly tagged "Unclassified — No Description"

### The problem it solves

Most public agencies never maintain their own commodity classification of what they buy. Without one, the basic strategic questions go unanswered: where does spend actually go, which commodities are bought across departments without coordination, and can any single purchase be traced and defended in an audit?

### How it works

A four-tier pipeline classifies every row, then attaches a full audit trail to each one:

1. **Curated keyword rules** — hand-written by procurement staff, in plain CSV.
2. **Account-code patterns** — catch subgrant / pass-through disbursements that aren't commodity buys.
3. **AI-assist resolver** — consumes the *saved* output of a one-time AI mining pass; no model is ever called at runtime.
4. **Explicit no-description tagging** — nothing is silently dropped.

**A defensible AI architecture.** AI was used *exactly once*, at build time (~$27 total), to mine recurring patterns from the long tail of descriptions no hand-written rule covered. Its output was frozen into rules and the model was never called again. The production system is **100% deterministic and fully auditable** — every classified row carries the exact rule that fired, a confidence level, and a machine-readable reason. The system never invents a code.

**A live web tool.** A deployed, password-gated Streamlit app handles single-record and bulk classification, exposes the taxonomy logic, and lets a user look up the exact rule behind any decision.

### Why it's reusable capability, not a one-off

The taxonomy and the method re-point at any new spend extract or contract portfolio without re-running AI or engaging a consultant:

- **Contract administration** — commodity-code an active contract portfolio; surface overlap and consolidation candidates; keep audit-ready contract records; group renewals by commodity class for planned sourcing waves instead of fragmented one-offs.
- **Supply chain & category management** — a commodity taxonomy is the backbone of category management and the prerequisite for strategic sourcing: spend visibility by commodity, demand aggregation across departments, supplier rationalization, and reduced tail spend.

**Location:** `/spend-analysis/` (engine, rules, methodology) and `/streamlit_app.py` (web interface).

---

## Other projects

### Vendor Research Agent — In Development
**Location:** `/vendor-research/` — automated agent that investigates vendors, analyzes market data, and generates research reports.

### Team Dashboard — Planned
**Location:** `/dashboard/` — web interface for procurement teams to access analysis tools without coding knowledge.

---

## Technology stack

- **Language:** Python 3.x
- **AI Platform:** Claude (Anthropic) — used during build, **not at runtime**
- **Web framework:** Streamlit
- **Key libraries:** pandas, openpyxl, python-docx
- **Environment:** GitHub Codespaces

## Repository structure

```
procurement-ai-tools/
├── spend-analysis/          NIGP classifier engine, rules, methodology
│   ├── scripts/             Python classification engine
│   ├── data/                Reference data (rules, NIGP catalog)
│   └── outputs/             Generated reports and analysis
├── outputs/                 Final classified data and deliverables
├── streamlit_app.py         Web interface (NIGP-Sourced Procurement Category Mapper)
├── vendor-research/         In progress
├── dashboard/               Planned
└── README.md                This file
```

## Author

**James H. Kirby III, CSCP, MS-SCM**
Independent public-sector procurement practitioner — 28 years of experience in procurement, contract administration, and supply chain. Subject-matter expert who also designed and built the classification engine, rule architecture, and web application in this repository.

## License

Personal project. Provided as-is for review and educational reference.

---

*Last updated: August 2026*
