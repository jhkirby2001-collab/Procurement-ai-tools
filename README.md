# Procurement AI Tools

Independent, AI-assisted automation tools for public-sector procurement work.

## Overview

This repository contains a suite of tools that automate manual procurement processes, increase efficiency, and provide analytical insights for strategic decision-making. The work is independent and not affiliated with or endorsed by any specific government agency.

**Goal:** Transform time-consuming manual tasks into automated, consistent, and scalable solutions that free procurement analysts to focus on high-value strategic work.

## Current Projects

### NIGP-Sourced Procurement Category Mapper — Operational

**Location:** `/spend-analysis/` and `/streamlit_app.py`

A reusable rules-based classification engine that maps unstructured procurement descriptions to:

- One of 17 Business Categories (mutually exclusive, collectively exhaustive)
- A standard NIGP 3-digit Class
- A standard NIGP 5-digit Class-Item (when specific enough)

**Headline performance** (validated on a 784,556-row public-sector procurement dataset, AP activity years 2017, 2020, 2021, 2023):

- 86.9% auto-classified by deterministic rule
- 17.4% flagged for human review
- End-to-end runtime: ~15 minutes
- 220 hand-curated rules + 6,766 AI-mined rules + 6 account-code patterns
- No AI is called at runtime — the classifier is fully deterministic and auditable

A live web interface (Streamlit) is deployed for single-record and bulk classification.

### Vendor Research Agent — In Development

**Location:** `/vendor-research/`

Automated research agent that investigates vendors, analyzes market data, and generates research reports.

### Team Dashboard — Planned

**Location:** `/dashboard/`

Web-based interface for procurement teams to access AI-powered analysis tools without coding knowledge.

## Technology Stack

- **Language:** Python 3.x
- **AI Platform:** Claude (Anthropic) — used during build, not at runtime
- **Web Framework:** Streamlit
- **Key Libraries:** pandas, openpyxl, python-docx
- **Development Environment:** GitHub Codespaces

## Repository Structure

```
procurement-ai-tools/
├── spend-analysis/          NIGP classifier engine, rules, methodology
│   ├── scripts/             Python classification engine
│   ├── data/                Reference data (rules, NIGP catalog)
│   └── outputs/             Generated reports and analysis
├── outputs/                 Leadership deliverables and final classified data
├── streamlit_app.py         Web interface (NIGP-Sourced Procurement Category Mapper)
├── vendor-research/         In progress
├── dashboard/               Planned
└── README.md                This file
```

## Getting Started

Each project folder contains its own README with specific setup and usage instructions.

**Prerequisites:**

- Python 3.8+
- Anthropic API key (only required for the one-time AI rule-mining step, not for runtime classification)
- Required Python packages (see `requirements.txt`)

## Author

**James H. Kirby III, CSCP, MS-SCM**
Independent public-sector procurement practitioner — 28 years of experience.

## License

Personal project. Provided as-is for review and educational reference.

---

*Last updated: May 2026*
