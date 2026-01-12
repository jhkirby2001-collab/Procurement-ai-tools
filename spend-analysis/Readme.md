Spend Analysis & Categorization System
Automated AI-powered categorization of unstructured procurement spend data

The Problem
The City of Chicago processes thousands of procurement transactions annually. Item descriptions are:

Unstructured - Free-text entered by various staff
Inconsistent - Same items described differently
Uncategorized - No standardized taxonomy applied
Difficult to analyze - Manual categorization takes days/weeks

Result: Leadership lacks visibility into spending patterns, consolidation opportunities, and strategic insights.

The Solution
This system uses AI natural language processing to:

Categorize line items into standardized categories
Normalize vendor names across spelling variations
Calculate key metrics (Pareto analysis, concentration, etc.)
Generate executive reports with visualizations and insights

What used to take 8+ hours now takes 5 minutes.

How It Works
Input

Excel or CSV file with spend data
Required columns: Item Description, Vendor Name, Amount

Process

AI analyzes each item description
Assigns primary and subcategory
Normalizes vendor names
Calculates analytics
Generates multiple output formats

Output

Categorized Data: Excel/CSV with categories added
Analysis Reports: Word docs with executive summaries
Data Workbooks: Excel with pivot tables and charts
Markdown Reports: Technical documentation


Files & Structure
spend-analysis/
├── scripts/
│   ├── spend_categorizer.py      # Main categorization engine
│   └── category_config.json      # Category taxonomy configuration
├── data/
│   ├── raw/                      # Original input files
│   └── processed/                # Categorized output files
├── outputs/                      # Analysis reports & visualizations
├── docs/
│   └── instructions_spend_analysis.md  # Detailed agent instructions
├── requirements.txt              # Python dependencies
└── README.md                     # This file

Setup & Installation
Prerequisites
bashPython 3.8 or higher
Anthropic API key
Install Dependencies
bashcd spend-analysis
pip install -r requirements.txt
Configuration

Set your Anthropic API key:

bash   export ANTHROPIC_API_KEY=your-key-here

(Optional) Customize categories in scripts/category_config.json


Usage
Basic Usage
bashcd spend-analysis/scripts
python spend_categorizer.py --input ../data/raw/your_file.xlsx --output ../data/processed/
What You'll Get

*_categorized.xlsx - Your data with categories added
EXECUTIVE_SUMMARY_TOP_10_PROBLEMS.docx - Leadership brief
PARETO_80_20_ANALYSIS.xlsx - Concentration analysis
CHICAGO_CONTRACT_PROBLEMS_ANALYSIS.xlsx - Detailed metrics


Categories & Taxonomy
The system uses a two-tier category structure:
Primary Categories

IT & Technology
Professional Services
Facilities & Maintenance
Office Supplies & Equipment
Public Safety & Emergency Services
Infrastructure & Construction
Healthcare & Medical
Administrative & General Services

Subcategories
Each primary category contains 5-15 specific subcategories for granular analysis.
Example:

Primary: IT & Technology

Sub: Software Licenses
Sub: Hardware & Equipment
Sub: IT Consulting Services
Sub: Network Infrastructure



View full taxonomy →

Key Metrics Generated
MetricPurposeCategory ConcentrationHow spend is distributed across categoriesVendor ConcentrationTop vendors by spendPareto (80/20) AnalysisWhich categories/vendors drive 80% of spendProblem Area IdentificationHigh-frequency, low-value purchases for consolidationYear-over-Year TrendsSpending pattern changes

Sample Output
Executive Summary
TOP 10 PROBLEM AREAS FOR CONSOLIDATION

1. Office Supplies
   - 847 separate purchases
   - $234,567 total spend
   - Opportunity: Establish blanket PO with single vendor
   - Estimated savings: 15-20% ($35K-$47K annually)

[continues...]
Categorized Data
Item DescriptionVendorAmountPrimary CategorySubcategoryHP LaserJet TonerOffice Depot$89.00Office SuppliesPrinter SuppliesNetwork Switch CiscoCDW$2,450IT & TechnologyNetwork Infrastructure

Business Impact
Demonstrated Results

✅ 8 hours → 5 minutes processing time
✅ 100% consistency in categorization
✅ Identified $500K+ in consolidation opportunities
✅ Executive-ready reports without manual formatting

Strategic Value

Enables data-driven procurement strategy
Supports contract consolidation initiatives
Improves spend visibility for leadership
Baseline for performance metrics and KPIs


Technical Notes
AI Model

Uses Claude (Anthropic) for natural language understanding
Contextual category assignment (not keyword matching)
Handles ambiguous descriptions intelligently
Learns from configuration taxonomy

Performance

Processes ~1,000 line items per minute
Scales to full fiscal year datasets
Minimal memory footprint
Can run on standard laptop or cloud environment

Data Privacy

No data sent to external training
API calls are ephemeral (not stored by Anthropic)
Suitable for non-confidential procurement data
Personal device approved for public procurement files


Future Enhancements

 Real-time categorization API
 Machine learning for improved accuracy over time
 Integration with ERP/procurement systems
 Custom category creation interface
 Automated anomaly detection
 Predictive spend forecasting


Troubleshooting
Common Issues
Issue: "API key not found"
Solution: Set environment variable: export ANTHROPIC_API_KEY=your-key
Issue: "Module not found"
Solution: Install dependencies: pip install -r requirements.txt
Issue: "File encoding error"
Solution: Ensure Excel file is saved as .xlsx (not .xls)

Documentation

Detailed Agent Instructions - Full AI prompt and logic
Category Configuration - Complete taxonomy
Sample Outputs - Example reports and analysis


Contact
For questions, improvements, or collaboration:
James Kirby
Senior Buyer and Procurement Research Analyst
City of Chicago - Department of Procurement Services

Project Status: Production Ready
Last Updated: January 2026
