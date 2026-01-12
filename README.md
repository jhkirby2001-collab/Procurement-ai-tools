Procurement AI Tools
AI-powered automation tools for the City of Chicago Department of Procurement Services

Overview
This repository contains a suite of AI-powered tools designed to automate manual procurement processes, increase efficiency, and provide deeper analytical insights for strategic decision-making.
Goal: Transform time-consuming manual tasks into automated, consistent, and scalable solutions that free up procurement analysts to focus on high-value strategic work.

Current Projects
✅ Spend Analysis & Categorization
Status: Operational
Location: /spend-analysis/
Automated system for categorizing unstructured procurement spend data using AI-powered natural language processing.
Key Features:

Automatically categorizes thousands of line items from unstructured descriptions
Normalizes vendor names across inconsistent data entry
Generates executive-ready analysis reports with visualizations
Identifies problem areas and opportunities for consolidation

Business Impact:

Time Savings: Reduces 8+ hours of manual categorization to minutes
Consistency: Eliminates subjective human categorization variance
Scalability: Can process entire fiscal year spend in one run
Insights: Uncovers patterns invisible in manual review

Technical Details:

Python-based categorization engine
Configurable category taxonomy
Handles Excel and CSV inputs
Outputs categorized data + analysis reports (Excel, Word, Markdown)

View detailed documentation →

Upcoming Projects
🔨 Vendor Research Agent
Status: In Development
Location: /vendor-research/
Automated research agent that investigates vendors, analyzes market data, and generates comprehensive research reports.
Planned Features:

Web search and data extraction
Automated vendor background research
Contract analysis and comparison
Risk assessment reporting

Expected Impact:

Accelerate vendor due diligence
Ensure consistent research standards
Support better negotiation positioning


🔨 Team Dashboard
Status: Planned
Location: /dashboard/
Web-based interface for the procurement team to access AI-powered analysis tools without coding knowledge.
Planned Features:

File upload for spend data analysis
Interactive data exploration
Natural language queries about procurement data
Automated report generation

Expected Impact:

Democratize AI tools across the team
Self-service analytics for all team members
Reduced bottlenecks on analyst time


Technology Stack

Language: Python 3.x
AI Platform: Claude (Anthropic)
Development Environment: GitHub Codespaces
Key Libraries: pandas, openpyxl, python-docx
MCP Tools: Brave Search, Puppeteer (planned)


Repository Structure
procurement-ai-tools/
├── spend-analysis/          # Completed: AI-powered spend categorization
│   ├── scripts/            # Python categorization engine
│   ├── data/               # Sample data (raw and processed)
│   ├── outputs/            # Generated reports and analysis
│   └── docs/               # Project documentation
├── vendor-research/         # In Progress: Automated vendor research
├── dashboard/               # Planned: Team-facing web interface
└── README.md               # This file

Getting Started
Each project folder contains its own README with specific setup and usage instructions.
Prerequisites

Python 3.8+
Anthropic API key
Required Python packages (see requirements.txt in each project)


Innovation & Impact
This work demonstrates the City of Chicago's commitment to:
✅ Digital Transformation - Leveraging AI for operational efficiency
✅ Data-Driven Decision Making - Moving from gut instinct to evidence-based procurement
✅ Process Innovation - Reimagining workflows for the modern era
✅ Strategic Resource Allocation - Freeing analysts from manual tasks to focus on strategy

Contact
James Kirby
Senior Buyer and Procurement Research Analyst
City of Chicago - Department of Procurement Services

License
Internal use - City of Chicago Department of Procurement Services

Last Updated: January 2026
