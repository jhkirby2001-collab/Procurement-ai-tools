# Spend Data Categorizer - Google Colab Instructions

## Overview
This guide walks you through running the Spend Data Categorizer script on Google Colab. The script automatically categorizes unstructured spend/procurement data and provides interactive analysis.

---

## Step 1: Create a New Colab Notebook
Go to [colab.research.google.com](https://colab.research.google.com) → **New Notebook**

---

## Step 2: Install Dependencies
Run this cell:
```python
!pip install pandas openpyxl anthropic
```

---

## Step 3: Create the Script Files
Run this cell to create all 3 files:

```python
# Create spend_categorizer.py
spend_categorizer_code = '''#!/usr/bin/env python3
"""
Spend Data Categorizer & Analyzer
"""

import argparse
import json
import os
import pandas as pd
import re
import time
from pathlib import Path
from typing import Optional, List, Dict, Tuple
import sys
from datetime import datetime

DEFAULT_CONFIG_FILE = "category_config.json"
DEFAULT_CATEGORY = "Uncategorized"
API_BATCH_SIZE = 25

def load_config(config_path: str = None) -> dict:
    if config_path is None:
        config_path = DEFAULT_CONFIG_FILE
    if not Path(config_path).exists():
        print(f"Creating default config: {config_path}")
        create_default_config(config_path)
    with open(config_path, 'r') as f:
        return json.load(f)

def create_default_config(config_path: str) -> None:
    default_config = {
        "_instructions": "Edit categories and keywords below. Keywords are case-insensitive.",
        "default_category": "Uncategorized",
        "use_ai_fallback": False,
        "ai_category_list": [],
        "categories": {
            "Aviation & Airport Services": ["airport", "aviation", "airside", "terminal", "o'hare", "midway", "runway", "aircraft", "airline", "flight", "hangar"],
            "IT & Technology": ["software", "hardware", "computer", "database", "network", "server", "technology", "cloud", "cyber", "digital", "telecom", "internet", "programming"],
            "Professional Services": ["consulting", "legal", "accounting", "audit", "architect", "engineering", "design", "planning", "training", "staffing"],
            "Facilities & Maintenance": ["custodial", "janitorial", "cleaning", "maintenance", "repair", "hvac", "plumbing", "electrical", "building", "landscaping"],
            "Construction": ["construction", "paving", "concrete", "asphalt", "demolition", "excavation", "roofing", "masonry", "renovation"],
            "Fleet & Transportation": ["vehicle", "automotive", "truck", "fleet", "fuel", "gasoline", "diesel", "towing", "tire", "transportation"],
            "Medical & Health": ["medical", "health", "pharmaceutical", "drug", "hospital", "clinic", "surgical", "dental", "ambulance"],
            "Parts & Supplies": ["parts", "supplies", "catalog", "equipment", "tools", "materials", "commodity", "commodities", "office supplies", "grainger"],
            "Environmental Services": ["remediation", "restoration", "environmental", "hazardous", "waste", "recycling", "mold", "asbestos"],
            "Courier & Delivery": ["courier", "messenger", "delivery", "shipping", "freight"],
            "Utilities": ["electric", "water", "utility", "power", "energy", "sewer"],
            "Public Safety": ["police", "fire department", "security", "safety", "emergency"]
        }
    }
    with open(config_path, 'w') as f:
        json.dump(default_config, f, indent=2)

def categorize_by_keywords(text: str, categories: dict, default: str) -> Tuple[str, str]:
    if pd.isna(text) or not str(text).strip():
        return default, None
    text_lower = str(text).lower()
    for category, keywords in categories.items():
        for keyword in keywords:
            if keyword.lower() in text_lower:
                return category, keyword
    return default, None

def get_claude_client():
    try:
        import anthropic
        return anthropic.Anthropic()
    except ImportError:
        print("Install anthropic: pip install anthropic")
        return None
    except Exception as e:
        print(f"Anthropic client error: {e}")
        return None

def categorize_batch_ai(client, items: List[str], category_list: List[str]) -> Dict[str, str]:
    if not client:
        return {}
    cat_instruction = f"Use these categories: {', '.join(category_list)}" if category_list else "Create appropriate spend categories"
    items_text = "\\n".join([f"{i+1}. {item[:200]}" for i, item in enumerate(items)])
    prompt = f"""Categorize each procurement item below into a single category.
{cat_instruction}

Return ONLY a JSON object: {{"1": "Category", "2": "Category", ...}}

Items:
{items_text}

JSON:"""
    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}]
        )
        text = response.content[0].text.strip()
        match = re.search(r'\\{[^{}]+\\}', text)
        if match:
            result = json.loads(match.group())
            return {items[int(k)-1]: v for k, v in result.items() if int(k) <= len(items)}
    except Exception as e:
        print(f"AI error: {e}")
    return {}

def load_data(file_path: str) -> pd.DataFrame:
    path = Path(file_path)
    suffix = path.suffix.lower()
    print(f"\\nLoading: {file_path}")
    try:
        if suffix == '.csv':
            for enc in ['utf-8', 'latin-1', 'cp1252']:
                try:
                    df = pd.read_csv(file_path, encoding=enc, low_memory=False)
                    break
                except UnicodeDecodeError:
                    continue
        elif suffix in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path)
        else:
            raise ValueError(f"Unsupported format: {suffix}")
        print(f"Loaded: {len(df):,} rows, {len(df.columns)} columns")
        return df
    except Exception as e:
        print(f"Error loading file: {e}")
        sys.exit(1)

def save_data(df: pd.DataFrame, path: str) -> None:
    suffix = Path(path).suffix.lower()
    if suffix in ['.xlsx', '.xls']:
        df.to_excel(path, index=False)
    else:
        df.to_csv(path, index=False)
    print(f"Saved: {path}")

def find_description_columns(df: pd.DataFrame) -> List[str]:
    keywords = ['description', 'item', 'name', 'detail', 'memo', 'purpose']
    return [col for col in df.columns if any(kw in col.lower() for kw in keywords)]

def find_amount_columns(df: pd.DataFrame) -> List[str]:
    keywords = ['amount', 'total', 'price', 'cost', 'spend', 'value', 'billed', 'ordered']
    return [col for col in df.columns if any(kw in col.lower() for kw in keywords)]

def find_vendor_columns(df: pd.DataFrame) -> List[str]:
    keywords = ['vendor', 'supplier', 'company', 'merchant']
    return [col for col in df.columns if any(kw in col.lower() for kw in keywords)]

def clean_amount_column(df: pd.DataFrame, col: str) -> pd.Series:
    amounts = df[col].copy()
    if amounts.dtype == 'object':
        amounts = amounts.replace(r'[\\$,]', '', regex=True)
        amounts = pd.to_numeric(amounts, errors='coerce')
    return amounts

def categorize_data(df: pd.DataFrame, description_cols: List[str], config: dict, use_ai: bool = False, category_col: str = "Procurement_Category") -> pd.DataFrame:
    df = df.copy()
    categories = config.get('categories', {})
    default = config.get('default_category', DEFAULT_CATEGORY)
    ai_categories = config.get('ai_category_list', []) or list(categories.keys())
    print(f"\\nCategorizing using columns: {description_cols}")
    results = []
    uncategorized = []
    for idx, row in df.iterrows():
        text = ""
        for col in description_cols:
            if col in row.index and pd.notna(row[col]):
                text = str(row[col])
                break
        cat, _ = categorize_by_keywords(text, categories, default)
        results.append(cat)
        if cat == default and text.strip():
            uncategorized.append((idx, text))
    df[category_col] = results
    keyword_matched = len(df) - len(uncategorized)
    print(f"  Keyword matched: {keyword_matched:,} ({keyword_matched/len(df)*100:.1f}%)")
    print(f"  Unmatched: {len(uncategorized):,}")
    if use_ai and uncategorized:
        client = get_claude_client()
        if client:
            print(f"\\nRunning AI categorization on {len(uncategorized)} items...")
            texts = [t for _, t in uncategorized]
            for i in range(0, len(texts), API_BATCH_SIZE):
                batch = texts[i:i+API_BATCH_SIZE]
                indices = [idx for idx, _ in uncategorized[i:i+API_BATCH_SIZE]]
                ai_results = categorize_batch_ai(client, batch, ai_categories)
                for idx, text in zip(indices, batch):
                    if text in ai_results:
                        df.at[idx, category_col] = ai_results[text]
                print(f"  Processed {min(i+API_BATCH_SIZE, len(texts))}/{len(texts)}")
                time.sleep(0.3)
    return df

class SpendAnalyzer:
    def __init__(self, df: pd.DataFrame, description_col: str, amount_col: str = None, vendor_col: str = None, category_col: str = "Procurement_Category"):
        self.df = df
        self.description_col = description_col
        self.amount_col = amount_col
        self.vendor_col = vendor_col
        self.category_col = category_col
        self.client = get_claude_client()
        if self.amount_col and self.amount_col in self.df.columns:
            self.df['_amount_clean'] = clean_amount_column(self.df, self.amount_col)

    def summary(self) -> str:
        lines = ["\\n=== DATA SUMMARY ===\\n"]
        lines.append(f"Total Records: {len(self.df):,}")
        lines.append(f"Categories: {self.df[self.category_col].nunique()}")
        if self.vendor_col:
            lines.append(f"Unique Vendors: {self.df[self.vendor_col].nunique():,}")
        if self.amount_col:
            total = self.df['_amount_clean'].sum()
            avg = self.df['_amount_clean'].mean()
            lines.append(f"Total Spend: ${total:,.2f}")
            lines.append(f"Average Transaction: ${avg:,.2f}")
        return "\\n".join(lines)

    def spend_by_category(self) -> str:
        lines = ["\\n=== SPEND BY CATEGORY ===\\n"]
        if '_amount_clean' not in self.df.columns:
            counts = self.df[self.category_col].value_counts()
            for cat, count in counts.items():
                pct = count / len(self.df) * 100
                lines.append(f"{cat:<35} {count:>8,} records ({pct:>5.1f}%)")
        else:
            spend = self.df.groupby(self.category_col)['_amount_clean'].agg(['sum', 'count'])
            spend = spend.sort_values('sum', ascending=False)
            total = spend['sum'].sum()
            lines.append(f"{'Category':<35} {'Spend':>15} {'Count':>8} {'%':>7}")
            lines.append("-" * 68)
            for cat, row in spend.iterrows():
                pct = row['sum'] / total * 100 if total > 0 else 0
                lines.append(f"{cat:<35} ${row['sum']:>14,.0f} {int(row['count']):>8,} {pct:>6.1f}%")
            lines.append("-" * 68)
            lines.append(f"{'TOTAL':<35} ${total:>14,.0f}")
        return "\\n".join(lines)

    def top_vendors(self, n: int = 10) -> str:
        if not self.vendor_col:
            return "No vendor column detected."
        lines = [f"\\n=== TOP {n} VENDORS ===\\n"]
        if '_amount_clean' in self.df.columns:
            top = self.df.groupby(self.vendor_col)['_amount_clean'].sum()
            top = top.sort_values(ascending=False).head(n)
            for i, (vendor, amount) in enumerate(top.items(), 1):
                lines.append(f"{i:>2}. {vendor[:40]:<40} ${amount:>15,.2f}")
        else:
            top = self.df[self.vendor_col].value_counts().head(n)
            for i, (vendor, count) in enumerate(top.items(), 1):
                lines.append(f"{i:>2}. {vendor[:40]:<40} {count:>8,} records")
        return "\\n".join(lines)

    def show_uncategorized(self) -> str:
        uncat = self.df[self.df[self.category_col] == DEFAULT_CATEGORY]
        lines = [f"\\n=== UNCATEGORIZED ITEMS: {len(uncat):,} ===\\n"]
        if len(uncat) == 0:
            return "No uncategorized items!"
        sample = uncat.head(10)
        for _, row in sample.iterrows():
            desc = str(row.get(self.description_col, 'N/A'))[:60]
            lines.append(f"  - {desc}")
        if len(uncat) > 10:
            lines.append(f"\\n  ... and {len(uncat)-10:,} more")
        return "\\n".join(lines)
'''

with open('spend_categorizer.py', 'w') as f:
    f.write(spend_categorizer_code)

# Create category_config.json
config = {
    "_instructions": "Edit categories and keywords below. Keywords are case-insensitive.",
    "default_category": "Uncategorized",
    "use_ai_fallback": False,
    "ai_category_list": [],
    "categories": {
        "Aviation & Airport Services": ["airport", "aviation", "airside", "terminal", "o'hare", "midway", "runway", "aircraft", "airline", "flight", "hangar"],
        "IT & Technology": ["software", "hardware", "computer", "database", "network", "server", "technology", "cloud", "cyber", "digital", "telecom", "internet", "programming"],
        "Professional Services": ["consulting", "legal", "accounting", "audit", "architect", "engineering", "design", "planning", "training", "staffing"],
        "Facilities & Maintenance": ["custodial", "janitorial", "cleaning", "maintenance", "repair", "hvac", "plumbing", "electrical", "building", "landscaping"],
        "Construction": ["construction", "paving", "concrete", "asphalt", "demolition", "excavation", "roofing", "masonry", "renovation"],
        "Fleet & Transportation": ["vehicle", "automotive", "truck", "fleet", "fuel", "gasoline", "diesel", "towing", "tire", "transportation"],
        "Medical & Health": ["medical", "health", "pharmaceutical", "drug", "hospital", "clinic", "surgical", "dental", "ambulance"],
        "Parts & Supplies": ["parts", "supplies", "catalog", "equipment", "tools", "materials", "commodity", "commodities", "office supplies", "grainger"],
        "Environmental Services": ["remediation", "restoration", "environmental", "hazardous", "waste", "recycling", "mold", "asbestos"],
        "Courier & Delivery": ["courier", "messenger", "delivery", "shipping", "freight"],
        "Utilities": ["electric", "water", "utility", "power", "energy", "sewer"],
        "Public Safety": ["police", "fire department", "security", "safety", "emergency"]
    }
}

import json
with open('category_config.json', 'w') as f:
    json.dump(config, f, indent=2)

print("Files created successfully!")
```

---

## Step 4: Upload Your Data File
Run this cell:
```python
from google.colab import files
uploaded = files.upload()
```
Then select your CSV or Excel file from your computer.

---

## Step 5: Run the Categorizer

```python
from spend_categorizer import *

# Load your data (change filename to match your upload)
df = load_data("your_file.csv")

# Load config
config = load_config("category_config.json")

# Categorize (change column name to match your data)
df = categorize_data(df, ["Item Description"], config)

# Save output
save_data(df, "categorized_output.csv")

# Analyze
analyzer = SpendAnalyzer(df, "Item Description", " Amount Billed ", "Vendor Name")
print(analyzer.summary())
print(analyzer.spend_by_category())
print(analyzer.top_vendors(10))
```

**Important:** Update these values to match your data:
- `"your_file.csv"` → your actual filename
- `"Item Description"` → your description column name
- `" Amount Billed "` → your amount column name
- `"Vendor Name"` → your vendor column name

---

## Step 6: Download the Result
```python
files.download("categorized_output.csv")
```

---

## Optional: Enable AI Categorization

To use Claude AI for better categorization of unmatched items:

```python
import os
os.environ["ANTHROPIC_API_KEY"] = "your-api-key-here"

# Then run with use_ai=True
df = categorize_data(df, ["Item Description"], config, use_ai=True)
```

---

## Customizing Categories

Edit the `category_config.json` file to add or modify categories:

```python
# View current config
with open('category_config.json', 'r') as f:
    config = json.load(f)
    print(json.dumps(config, indent=2))

# Add a new category
config['categories']['New Category Name'] = ["keyword1", "keyword2", "keyword3"]

# Save updated config
with open('category_config.json', 'w') as f:
    json.dump(config, f, indent=2)
```

---

## Analysis Functions

After categorization, use these functions:

| Function | Description |
|----------|-------------|
| `analyzer.summary()` | Overall data summary |
| `analyzer.spend_by_category()` | Spend breakdown by category |
| `analyzer.top_vendors(n)` | Top n vendors by spend |
| `analyzer.show_uncategorized()` | View uncategorized items |

---

## Troubleshooting

**"Column not found" error:**
- Run `print(df.columns.tolist())` to see exact column names
- Column names are case-sensitive and may have extra spaces

**Low match rate:**
- Add more keywords to `category_config.json`
- Enable AI categorization with `use_ai=True`

**File upload issues:**
- Ensure file is CSV or Excel format (.csv, .xlsx, .xls)
- Check file size (Colab has limits for large files)
