#!/usr/bin/env python3
"""
Spend Data Categorizer & Analyzer
=================================
A flexible tool to categorize unstructured spend data and interactively analyze it.

Phase 1: Categorize data using keyword rules and/or AI
Phase 2: Interactive analysis - ask questions in plain English

Usage:
    python spend_categorizer.py "spen data1.csv" -d "Item Description"
    python spend_categorizer.py data.xlsx -d "Description" --use-ai
    python spend_categorizer.py data.csv --auto-detect --interactive

Set ANTHROPIC_API_KEY for AI features.
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

# =============================================================================
# CONFIGURATION
# =============================================================================

DEFAULT_CONFIG_FILE = "category_config.json"
DEFAULT_CATEGORY = "Uncategorized"
API_BATCH_SIZE = 25


# =============================================================================
# CONFIG FILE MANAGEMENT
# =============================================================================

def load_config(config_path: str = None) -> dict:
    """Load category configuration from JSON file."""
    if config_path is None:
        config_path = DEFAULT_CONFIG_FILE

    if not Path(config_path).exists():
        print(f"Creating default config: {config_path}")
        create_default_config(config_path)

    with open(config_path, 'r') as f:
        return json.load(f)


def create_default_config(config_path: str) -> None:
    """Create default configuration file."""
    default_config = {
        "_instructions": "Edit categories and keywords below. Keywords are case-insensitive.",
        "default_category": "Uncategorized",
        "use_ai_fallback": False,
        "ai_category_list": [],
        "categories": {
            "Aviation & Airport Services": [
                "airport", "aviation", "airside", "terminal", "o'hare", "midway",
                "runway", "aircraft", "airline", "flight", "hangar"
            ],
            "IT & Technology": [
                "software", "hardware", "computer", "database", "network",
                "server", "technology", "cloud", "cyber", "digital",
                "telecom", "internet", "programming"
            ],
            "Professional Services": [
                "consulting", "legal", "accounting", "audit", "architect",
                "engineering", "design", "planning", "training", "staffing"
            ],
            "Facilities & Maintenance": [
                "custodial", "janitorial", "cleaning", "maintenance", "repair",
                "hvac", "plumbing", "electrical", "building", "landscaping"
            ],
            "Construction": [
                "construction", "paving", "concrete", "asphalt", "demolition",
                "excavation", "roofing", "masonry", "renovation"
            ],
            "Fleet & Transportation": [
                "vehicle", "automotive", "truck", "fleet", "fuel", "gasoline",
                "diesel", "towing", "tire", "transportation"
            ],
            "Medical & Health": [
                "medical", "health", "pharmaceutical", "drug", "hospital",
                "clinic", "surgical", "dental", "ambulance"
            ],
            "Parts & Supplies": [
                "parts", "supplies", "catalog", "equipment", "tools",
                "materials", "commodity", "commodities", "office supplies"
            ],
            "Environmental Services": [
                "remediation", "restoration", "environmental", "hazardous",
                "waste", "recycling", "mold", "asbestos"
            ],
            "Utilities": [
                "electric", "water", "utility", "power", "energy", "sewer"
            ],
            "Public Safety": [
                "police", "fire department", "security", "safety", "emergency"
            ]
        }
    }

    with open(config_path, 'w') as f:
        json.dump(default_config, f, indent=2)


# =============================================================================
# KEYWORD CATEGORIZATION
# =============================================================================

def categorize_by_keywords(text: str, categories: dict, default: str) -> Tuple[str, str]:
    """Categorize text using keyword matching. Returns (category, matched_keyword)."""
    if pd.isna(text) or not str(text).strip():
        return default, None

    text_lower = str(text).lower()

    for category, keywords in categories.items():
        for keyword in keywords:
            if keyword.lower() in text_lower:
                return category, keyword

    return default, None


# =============================================================================
# AI CATEGORIZATION (Claude)
# =============================================================================

def get_claude_client():
    """Get Anthropic client."""
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
    """Use Claude to categorize a batch of items."""
    if not client:
        return {}

    cat_instruction = f"Use these categories: {', '.join(category_list)}" if category_list else "Create appropriate spend categories"

    items_text = "\n".join([f"{i+1}. {item[:200]}" for i, item in enumerate(items)])

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
        match = re.search(r'\{[^{}]+\}', text)
        if match:
            result = json.loads(match.group())
            return {items[int(k)-1]: v for k, v in result.items() if int(k) <= len(items)}
    except Exception as e:
        print(f"AI error: {e}")
    return {}


# =============================================================================
# DATA LOADING
# =============================================================================

def load_data(file_path: str) -> pd.DataFrame:
    """Load CSV or Excel file."""
    path = Path(file_path)
    suffix = path.suffix.lower()

    print(f"\nLoading: {file_path}")

    try:
        if suffix == '.csv':
            for enc in ['utf-8', 'latin-1', 'cp1252']:
                try:
                    df = pd.read_csv(file_path, encoding=enc)
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
    """Save DataFrame to file."""
    suffix = Path(path).suffix.lower()
    if suffix in ['.xlsx', '.xls']:
        df.to_excel(path, index=False)
    else:
        df.to_csv(path, index=False)
    print(f"Saved: {path}")


def find_description_columns(df: pd.DataFrame) -> List[str]:
    """Auto-detect description columns."""
    keywords = ['description', 'item', 'name', 'detail', 'memo', 'purpose']
    return [col for col in df.columns if any(kw in col.lower() for kw in keywords)]


def find_amount_columns(df: pd.DataFrame) -> List[str]:
    """Auto-detect amount columns."""
    keywords = ['amount', 'total', 'price', 'cost', 'spend', 'value', 'billed', 'ordered']
    return [col for col in df.columns if any(kw in col.lower() for kw in keywords)]


def find_vendor_columns(df: pd.DataFrame) -> List[str]:
    """Auto-detect vendor columns."""
    keywords = ['vendor', 'supplier', 'company', 'merchant']
    return [col for col in df.columns if any(kw in col.lower() for kw in keywords)]


def clean_amount_column(df: pd.DataFrame, col: str) -> pd.Series:
    """Clean currency formatting from amount column."""
    amounts = df[col].copy()
    if amounts.dtype == 'object':
        amounts = amounts.replace(r'[\$,]', '', regex=True)
        amounts = pd.to_numeric(amounts, errors='coerce')
    return amounts


# =============================================================================
# MAIN CATEGORIZATION
# =============================================================================

def categorize_data(
    df: pd.DataFrame,
    description_cols: List[str],
    config: dict,
    use_ai: bool = False,
    category_col: str = "Procurement_Category"
) -> pd.DataFrame:
    """Categorize DataFrame using keywords and optionally AI."""
    df = df.copy()
    categories = config.get('categories', {})
    default = config.get('default_category', DEFAULT_CATEGORY)
    ai_categories = config.get('ai_category_list', []) or list(categories.keys())

    print(f"\nCategorizing using columns: {description_cols}")

    # Keyword pass
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

    # AI pass for unmatched
    if use_ai and uncategorized:
        client = get_claude_client()
        if client:
            print(f"\nRunning AI categorization on {len(uncategorized)} items...")
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


# =============================================================================
# INTERACTIVE ANALYSIS MODE
# =============================================================================

class SpendAnalyzer:
    """Interactive spend data analyzer using Claude for natural language queries."""

    def __init__(self, df: pd.DataFrame, description_col: str, amount_col: str = None,
                 vendor_col: str = None, category_col: str = "Procurement_Category"):
        self.df = df
        self.description_col = description_col
        self.amount_col = amount_col
        self.vendor_col = vendor_col
        self.category_col = category_col
        self.client = get_claude_client()

        # Clean amount column if present
        if self.amount_col and self.amount_col in self.df.columns:
            self.df['_amount_clean'] = clean_amount_column(self.df, self.amount_col)

    def get_data_context(self) -> str:
        """Generate context about the data for Claude."""
        context = f"""DataFrame 'df' with {len(self.df):,} rows and columns: {list(self.df.columns)}

Key columns:
- Description: '{self.description_col}'
- Category: '{self.category_col}'"""

        if self.vendor_col:
            context += f"\n- Vendor: '{self.vendor_col}'"
        if self.amount_col:
            context += f"\n- Amount: '{self.amount_col}' (cleaned as '_amount_clean')"
            total = self.df['_amount_clean'].sum()
            context += f"\n- Total spend: ${total:,.2f}"

        context += f"\n\nCategories: {self.df[self.category_col].unique().tolist()}"

        return context

    def run_query(self, question: str) -> str:
        """Use Claude to interpret question and run analysis."""
        if not self.client:
            return self._run_builtin_query(question)

        context = self.get_data_context()

        prompt = f"""You are a data analyst. Given this pandas DataFrame:

{context}

User question: {question}

Write Python code to answer this question. The DataFrame is already loaded as 'df'.
Use '_amount_clean' for any amount calculations.
Return ONLY the Python code block, no explanation.
Make the output formatted nicely for terminal display.
For tables, use .to_string() with formatting.
Always print the result.

```python
"""

        try:
            response = self.client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=1024,
                messages=[{"role": "user", "content": prompt}]
            )

            code = response.content[0].text.strip()
            # Extract code from markdown if present
            if "```" in code:
                code = re.search(r'```(?:python)?\n?(.*?)```', code, re.DOTALL)
                code = code.group(1) if code else ""

            # Execute the code
            import io
            from contextlib import redirect_stdout

            output = io.StringIO()
            local_vars = {'df': self.df, 'pd': pd}

            with redirect_stdout(output):
                exec(code, local_vars)

            return output.getvalue() or "Query executed (no output)"

        except Exception as e:
            return f"Error: {e}\n\nTrying built-in analysis..."

    def _run_builtin_query(self, question: str) -> str:
        """Run built-in queries without AI."""
        q = question.lower()

        if 'top' in q and 'vendor' in q:
            return self._top_vendors(10)
        elif 'category' in q or 'breakdown' in q:
            return self._spend_by_category()
        elif 'uncategorized' in q or 'unmatched' in q:
            return self._show_uncategorized()
        elif 'summary' in q or 'overview' in q:
            return self._summary()
        else:
            return self._help()

    def _summary(self) -> str:
        """Overall data summary."""
        lines = ["\n=== DATA SUMMARY ===\n"]
        lines.append(f"Total Records: {len(self.df):,}")
        lines.append(f"Categories: {self.df[self.category_col].nunique()}")

        if self.vendor_col:
            lines.append(f"Unique Vendors: {self.df[self.vendor_col].nunique():,}")

        if self.amount_col:
            total = self.df['_amount_clean'].sum()
            avg = self.df['_amount_clean'].mean()
            lines.append(f"Total Spend: ${total:,.2f}")
            lines.append(f"Average Transaction: ${avg:,.2f}")

        return "\n".join(lines)

    def _spend_by_category(self) -> str:
        """Spend breakdown by category."""
        lines = ["\n=== SPEND BY CATEGORY ===\n"]

        if '_amount_clean' not in self.df.columns:
            # Just show counts
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

        return "\n".join(lines)

    def _top_vendors(self, n: int = 10) -> str:
        """Top vendors by spend."""
        if not self.vendor_col:
            return "No vendor column detected."

        lines = [f"\n=== TOP {n} VENDORS ===\n"]

        if '_amount_clean' in self.df.columns:
            top = self.df.groupby(self.vendor_col)['_amount_clean'].sum()
            top = top.sort_values(ascending=False).head(n)

            for i, (vendor, amount) in enumerate(top.items(), 1):
                lines.append(f"{i:>2}. {vendor[:40]:<40} ${amount:>15,.2f}")
        else:
            top = self.df[self.vendor_col].value_counts().head(n)
            for i, (vendor, count) in enumerate(top.items(), 1):
                lines.append(f"{i:>2}. {vendor[:40]:<40} {count:>8,} records")

        return "\n".join(lines)

    def _show_uncategorized(self) -> str:
        """Show uncategorized items."""
        uncat = self.df[self.df[self.category_col] == DEFAULT_CATEGORY]
        lines = [f"\n=== UNCATEGORIZED ITEMS: {len(uncat):,} ===\n"]

        if len(uncat) == 0:
            return "No uncategorized items!"

        # Show sample
        sample = uncat.head(10)
        for _, row in sample.iterrows():
            desc = str(row.get(self.description_col, 'N/A'))[:60]
            lines.append(f"  - {desc}")

        if len(uncat) > 10:
            lines.append(f"\n  ... and {len(uncat)-10:,} more")

        return "\n".join(lines)

    def _help(self) -> str:
        """Show help for built-in queries."""
        return """
Available queries (or ask anything with AI enabled):
  - "summary" / "overview"      - Overall data summary
  - "spend by category"         - Breakdown by category
  - "top vendors"               - Top vendors by spend
  - "uncategorized"             - Show unmatched items
  - "help"                      - Show this help

With AI enabled, ask natural language questions like:
  - "Show me the top 5 IT vendors"
  - "What percentage of spend is uncategorized?"
  - "List all vendors with more than $1M in spend"
  - "Trend of construction spend by quarter"
"""


def interactive_mode(analyzer: SpendAnalyzer):
    """Run interactive analysis session."""
    print("\n" + "="*60)
    print("INTERACTIVE ANALYSIS MODE")
    print("="*60)
    print("Ask questions about your spend data in plain English.")
    print("Type 'help' for examples, 'quit' to exit.\n")

    # Show initial summary
    print(analyzer._summary())
    print(analyzer._spend_by_category())

    while True:
        try:
            question = input("\n> ").strip()

            if not question:
                continue
            if question.lower() in ['quit', 'exit', 'q']:
                print("Exiting analysis mode.")
                break
            if question.lower() == 'help':
                print(analyzer._help())
                continue

            result = analyzer.run_query(question)
            print(result)

        except KeyboardInterrupt:
            print("\nExiting.")
            break
        except Exception as e:
            print(f"Error: {e}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Categorize and analyze spend data",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s data.csv -d "Item Description"
  %(prog)s data.xlsx -d "Description" --use-ai
  %(prog)s data.csv --auto-detect --interactive

Environment:
  Set ANTHROPIC_API_KEY for AI features
        """
    )

    parser.add_argument("input_file", nargs="?", help="Input CSV or Excel file")
    parser.add_argument("-d", "--description-col", action="append", dest="desc_cols")
    parser.add_argument("-a", "--amount-col", help="Amount column name")
    parser.add_argument("-v", "--vendor-col", help="Vendor column name")
    parser.add_argument("-o", "--output", help="Output file path")
    parser.add_argument("-c", "--config", default=DEFAULT_CONFIG_FILE)
    parser.add_argument("--use-ai", action="store_true", help="Use AI for categorization")
    parser.add_argument("--auto-detect", action="store_true")
    parser.add_argument("--category-col", default="Procurement_Category")
    parser.add_argument("--list-columns", action="store_true")
    parser.add_argument("--interactive", "-i", action="store_true", help="Enter interactive analysis mode")
    parser.add_argument("--create-config", action="store_true")

    args = parser.parse_args()

    # Create config only
    if args.create_config:
        create_default_config(args.config)
        print(f"Created: {args.config}")
        return

    if not args.input_file:
        parser.print_help()
        return

    # Load data
    df = load_data(args.input_file)

    # List columns
    if args.list_columns:
        print("\nColumns:")
        for i, col in enumerate(df.columns, 1):
            print(f"  {i:>2}. {col}")
        return

    # Load config
    config = load_config(args.config)

    # Detect columns
    desc_cols = args.desc_cols or (find_description_columns(df) if args.auto_detect else None)
    amount_col = args.amount_col or (find_amount_columns(df)[0] if find_amount_columns(df) else None)
    vendor_col = args.vendor_col or (find_vendor_columns(df)[0] if find_vendor_columns(df) else None)

    if not desc_cols:
        print("Specify --description-col or use --auto-detect")
        print("Columns:", list(df.columns))
        return

    print(f"\nUsing columns:")
    print(f"  Description: {desc_cols}")
    if amount_col: print(f"  Amount: {amount_col}")
    if vendor_col: print(f"  Vendor: {vendor_col}")

    # Categorize
    df = categorize_data(df, desc_cols, config, args.use_ai, args.category_col)

    # Save output
    if args.output:
        output_path = args.output
    else:
        input_path = Path(args.input_file)
        output_path = str(input_path.parent / f"{input_path.stem}_categorized.csv")

    save_data(df, output_path)

    # Interactive mode
    if args.interactive:
        analyzer = SpendAnalyzer(df, desc_cols[0], amount_col, vendor_col, args.category_col)
        interactive_mode(analyzer)
    else:
        # Show summary
        analyzer = SpendAnalyzer(df, desc_cols[0], amount_col, vendor_col, args.category_col)
        print(analyzer._summary())
        print(analyzer._spend_by_category())
        print("\nRun with --interactive (-i) for analysis mode")


if __name__ == "__main__":
    main()
