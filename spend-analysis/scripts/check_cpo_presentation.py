#!/usr/bin/env python3
"""
Read the EXHIBIT_B_CPO_PRESENTATION.xlsx file and check the drinking water service detail tab
"""

import pandas as pd
from pathlib import Path

file_path = Path("/workspaces/Spend_Analysis_and_Categorization/spend-analysis/outputs/EXHIBIT_B_CPO_PRESENTATION.xlsx")

print("=" * 80)
print("READING CPO PRESENTATION WORKBOOK")
print("=" * 80)
print()

try:
    # Get all sheet names first
    xl_file = pd.ExcelFile(file_path)
    print(f"Available sheets in workbook:")
    for i, sheet_name in enumerate(xl_file.sheet_names, 1):
        print(f"  {i}. {sheet_name}")
    print()

    # Look for drinking water related sheets
    water_sheets = [s for s in xl_file.sheet_names if 'water' in s.lower() or 'drink' in s.lower()]

    if water_sheets:
        print(f"Found {len(water_sheets)} drinking water related sheet(s):")
        for sheet in water_sheets:
            print(f"\n{'='*80}")
            print(f"SHEET: {sheet}")
            print(f"{'='*80}")

            df = pd.read_excel(file_path, sheet_name=sheet)
            print(f"Shape: {df.shape[0]} rows x {df.shape[1]} columns")
            print(f"\nColumns: {list(df.columns)}")
            print(f"\nFirst 10 rows:")
            print(df.head(10))
            print()

            # Check for summary/totals
            if 'Total' in df.columns or 'TOTAL' in df.columns:
                total_col = 'Total' if 'Total' in df.columns else 'TOTAL'
                print(f"Total column sum: {df[total_col].sum():,.2f}")

            # Check for amount columns
            amount_cols = [col for col in df.columns if 'amount' in col.lower() or 'spend' in col.lower() or 'total' in col.lower()]
            if amount_cols:
                print(f"\nAmount columns found:")
                for col in amount_cols:
                    if pd.api.types.is_numeric_dtype(df[col]):
                        print(f"  {col}: ${df[col].sum():,.2f}")

            # Check for vendor columns
            vendor_cols = [col for col in df.columns if 'vendor' in col.lower()]
            if vendor_cols:
                print(f"\nUnique vendors in {vendor_cols[0]}:")
                print(df[vendor_cols[0]].value_counts())

            # Check for department columns
            dept_cols = [col for col in df.columns if 'dept' in col.lower() or 'department' in col.lower()]
            if dept_cols:
                print(f"\nUnique departments in {dept_cols[0]}:")
                print(df[dept_cols[0]].value_counts())

            print()
    else:
        print("⚠️  No sheets with 'water' or 'drink' in the name found")
        print("\nLet me check all sheets for drinking water content...")

        for sheet in xl_file.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet)
            print(f"\n{sheet}: {df.shape[0]} rows x {df.shape[1]} columns")
            print(f"  Columns: {list(df.columns)[:5]}...")  # Show first 5 columns

except Exception as e:
    print(f"Error reading file: {e}")
    import traceback
    traceback.print_exc()
