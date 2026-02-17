import pandas as pd
import numpy as np
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

# File path
file_path = "/workspaces/Spend_Analysis_and_Categorization/spend-analysis/data/raw/BottledWater raw data.xlsx"

print("=" * 100)
print("BOTTLED WATER SPEND ANALYSIS - COMPREHENSIVE REPORT")
print("=" * 100)

# Read all sheets
excel_file = pd.ExcelFile(file_path)
print(f"\n📊 FILE STRUCTURE")
print(f"Number of sheets: {len(excel_file.sheet_names)}")
print(f"Sheet names: {excel_file.sheet_names}")

# Dictionary to store all dataframes
dfs = {}

# Analyze each sheet
for sheet_name in excel_file.sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    dfs[sheet_name] = df

    print(f"\n{'=' * 100}")
    print(f"SHEET: {sheet_name}")
    print(f"{'=' * 100}")

    print(f"\n📐 DATA STRUCTURE:")
    print(f"Rows: {len(df)}")
    print(f"Columns: {len(df.columns)}")
    print(f"\nColumn names and types:")
    for col in df.columns:
        print(f"  - {col}: {df[col].dtype}")

    print(f"\n📋 FIRST 10 ROWS:")
    print(df.head(10).to_string())

    print(f"\n📊 DATA SUMMARY:")
    print(df.describe(include='all').to_string())

    print(f"\n🔍 MISSING VALUES:")
    missing = df.isnull().sum()
    if missing.sum() > 0:
        print(missing[missing > 0])
    else:
        print("No missing values")

    print(f"\n📈 UNIQUE VALUES PER COLUMN:")
    for col in df.columns:
        unique_count = df[col].nunique()
        print(f"  - {col}: {unique_count} unique values")
        if unique_count <= 20 and df[col].dtype == 'object':
            print(f"    Values: {df[col].unique().tolist()}")

print("\n" + "=" * 100)
print("SPEND ANALYSIS AND CATEGORIZATION")
print("=" * 100)

# Perform detailed spend analysis on each sheet
for sheet_name, df in dfs.items():
    print(f"\n{'=' * 100}")
    print(f"SPEND ANALYSIS - {sheet_name}")
    print(f"{'=' * 100}")

    # Identify spend/amount columns
    amount_cols = [col for col in df.columns if any(keyword in col.lower() for keyword in ['amount', 'spend', 'cost', 'price', 'total', 'value'])]

    if amount_cols:
        print(f"\n💰 SPEND COLUMNS IDENTIFIED: {amount_cols}")

        for col in amount_cols:
            # Clean and convert to numeric
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.replace('$', '').str.replace(',', '').str.strip()
            df[col] = pd.to_numeric(df[col], errors='coerce')

            total_spend = df[col].sum()
            avg_spend = df[col].mean()
            median_spend = df[col].median()
            min_spend = df[col].min()
            max_spend = df[col].max()

            print(f"\n  {col}:")
            print(f"    Total Spend: ${total_spend:,.2f}")
            print(f"    Average: ${avg_spend:,.2f}")
            print(f"    Median: ${median_spend:,.2f}")
            print(f"    Min: ${min_spend:,.2f}")
            print(f"    Max: ${max_spend:,.2f}")
            print(f"    Std Dev: ${df[col].std():,.2f}")

    # Identify categorical columns for categorization
    cat_cols = [col for col in df.columns if df[col].dtype == 'object' and df[col].nunique() < 50]

    if cat_cols and amount_cols:
        print(f"\n📂 CATEGORIZATION ANALYSIS:")

        for cat_col in cat_cols[:5]:  # Limit to top 5 categorical columns
            print(f"\n  By {cat_col}:")

            for amount_col in amount_cols:
                category_spend = df.groupby(cat_col)[amount_col].agg(['sum', 'mean', 'count']).sort_values('sum', ascending=False)
                category_spend['% of Total'] = (category_spend['sum'] / category_spend['sum'].sum() * 100)
                category_spend['Cumulative %'] = category_spend['% of Total'].cumsum()

                print(f"\n    Spend by {cat_col} ({amount_col}):")
                print(category_spend.to_string())

                # Pareto analysis (80/20 rule)
                pareto_threshold = category_spend[category_spend['Cumulative %'] <= 80]
                print(f"\n    📊 PARETO ANALYSIS (80/20 Rule):")
                print(f"    Top {len(pareto_threshold)} {cat_col}(s) account for 80% of spend")

print("\n" + "=" * 100)
print("EXECUTIVE SUMMARY")
print("=" * 100)

# Generate executive summary
total_records = sum(len(df) for df in dfs.values())
print(f"\n📊 OVERALL METRICS:")
print(f"Total Records Across All Sheets: {total_records:,}")
print(f"Number of Data Sources (Sheets): {len(dfs)}")

# Calculate total spend across all sheets
total_spend_all = 0
for sheet_name, df in dfs.items():
    amount_cols = [col for col in df.columns if any(keyword in col.lower() for keyword in ['amount', 'spend', 'cost', 'price', 'total', 'value'])]
    for col in amount_cols:
        if df[col].dtype in ['float64', 'int64']:
            total_spend_all += df[col].sum()

print(f"Total Spend (All Sheets): ${total_spend_all:,.2f}")

print("\n" + "=" * 100)
print("ANALYSIS COMPLETE")
print("=" * 100)
