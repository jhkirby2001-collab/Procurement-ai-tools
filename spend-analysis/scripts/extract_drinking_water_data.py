#!/usr/bin/env python3
"""
Extract drinking water spend data from Exhibit B files for vendor consolidation pilot analysis
"""

import pandas as pd
import numpy as np
from pathlib import Path

# File paths
bottled_water_file = Path("/workspaces/Spend_Analysis_and_Categorization/spend-analysis/data/raw/correctedBottledWater raw data.xlsx")
exhibit_b_file = Path("/workspaces/Spend_Analysis_and_Categorization/spend-analysis/data/raw/Exhibit B Report Jan 2021- Jan 2026.xlsx")

print("=" * 80)
print("DRINKING WATER SPEND DATA EXTRACTION")
print("=" * 80)
print()

# Read bottled water data
if bottled_water_file.exists():
    print(f"Reading: {bottled_water_file.name}")
    try:
        df_water = pd.read_excel(bottled_water_file, sheet_name=0)
        print(f"✓ Loaded {len(df_water)} rows from bottled water file")
        print(f"✓ Columns: {list(df_water.columns)}")
        print()

        # Display first few rows
        print("First 5 rows:")
        print(df_water.head())
        print()

        # This file has different column structure - skip for now
        print("Note: This file has different structure (PO-based), using main Exhibit B file instead")
        print()

    except Exception as e:
        print(f"✗ Error reading bottled water file: {e}")
        import traceback
        traceback.print_exc()
else:
    print(f"✗ File not found: {bottled_water_file}")

print("\n" + "="*80)
print()

# Try to read main Exhibit B file as well
if exhibit_b_file.exists():
    print(f"Reading: {exhibit_b_file.name}")
    try:
        # Read the first sheet
        df_exhibit = pd.read_excel(exhibit_b_file, sheet_name=0)
        print(f"✓ Loaded {len(df_exhibit)} rows from Exhibit B file")
        print(f"✓ Columns: {list(df_exhibit.columns)}")
        print()

        # Filter for drinking water related vendors if possible
        if 'VENDOR NAME' in df_exhibit.columns or 'Vendor Name' in df_exhibit.columns:
            vendor_col = 'VENDOR NAME' if 'VENDOR NAME' in df_exhibit.columns else 'Vendor Name'
            dept_col = 'Department Name' if 'Department Name' in df_exhibit.columns else 'DEPARTMENT NAME'
            amount_col = 'Requisition Amount' if 'Requisition Amount' in df_exhibit.columns else 'REQUISITION AMOUNT'

            water_vendors = ['DS SERVICES', 'BLUETRITON', 'HINCKLEY', 'BREWSMART', 'WATER', 'SPRING']
            water_mask = df_exhibit[vendor_col].str.upper().str.contains('|'.join(water_vendors), na=False)
            df_water_exhibit = df_exhibit[water_mask].copy()

            if len(df_water_exhibit) > 0:
                print(f"\n{'='*80}")
                print(f"DRINKING WATER SPEND ANALYSIS (EXHIBIT B)")
                print(f"{'='*80}")
                print(f"Found {len(df_water_exhibit)} drinking water transactions")
                print()

                # Vendor summary
                print("VENDOR SUMMARY:")
                print("-" * 80)
                vendor_summary = df_water_exhibit.groupby(vendor_col).agg({
                    amount_col: ['sum', 'count', 'mean'],
                    dept_col: 'nunique'
                }).round(2)
                vendor_summary.columns = ['Total Spend', 'Transactions', 'Avg Transaction', 'Departments']
                vendor_summary = vendor_summary.sort_values('Total Spend', ascending=False)
                vendor_summary['% of Total'] = (vendor_summary['Total Spend'] / vendor_summary['Total Spend'].sum() * 100).round(1)
                print(vendor_summary)
                print()

                # Department summary
                print("DEPARTMENT SUMMARY:")
                print("-" * 80)
                dept_summary = df_water_exhibit.groupby(dept_col).agg({
                    amount_col: ['sum', 'count', 'mean'],
                    vendor_col: 'nunique'
                }).round(2)
                dept_summary.columns = ['Total Spend', 'Transactions', 'Avg Transaction', 'Vendors']
                dept_summary = dept_summary.sort_values('Total Spend', ascending=False)
                print(dept_summary)
                print()

                # Vendor-Department Matrix (SPEND)
                print("VENDOR-DEPARTMENT SPEND MATRIX:")
                print("-" * 80)
                pivot_spend = pd.pivot_table(
                    df_water_exhibit,
                    values=amount_col,
                    index=dept_col,
                    columns=vendor_col,
                    aggfunc='sum',
                    fill_value=0
                )
                pivot_spend['TOTAL'] = pivot_spend.sum(axis=1)
                pivot_spend.loc['TOTAL'] = pivot_spend.sum(axis=0)
                print(pivot_spend.round(2))
                print()

                # Vendor-Department Matrix (TRANSACTION COUNT)
                print("VENDOR-DEPARTMENT TRANSACTION COUNT:")
                print("-" * 80)
                pivot_count = pd.pivot_table(
                    df_water_exhibit,
                    values=amount_col,
                    index=dept_col,
                    columns=vendor_col,
                    aggfunc='count',
                    fill_value=0
                )
                pivot_count['TOTAL'] = pivot_count.sum(axis=1)
                pivot_count.loc['TOTAL'] = pivot_count.sum(axis=0)
                print(pivot_count.astype(int))
                print()

                # List departments using multiple vendors
                multi_vendor_depts = dept_summary[dept_summary['Vendors'] > 1]
                if len(multi_vendor_depts) > 0:
                    print("⚠️  DEPARTMENTS USING MULTIPLE VENDORS (FRAGMENTATION):")
                    print("-" * 80)
                    for dept_name in multi_vendor_depts.index:
                        dept_data = df_water_exhibit[df_water_exhibit[dept_col] == dept_name]
                        vendors = dept_data[vendor_col].unique()
                        vendor_count = len(vendors)
                        total_spend = dept_data[amount_col].sum()
                        print(f"{dept_name}:")
                        print(f"  - Using {vendor_count} vendors: {', '.join(vendors)}")
                        print(f"  - Total spend: ${total_spend:,.2f}")
                        print()

                # Overall totals
                total_spend = df_water_exhibit[amount_col].sum()
                total_transactions = len(df_water_exhibit)
                unique_vendors = df_water_exhibit[vendor_col].nunique()
                unique_depts = df_water_exhibit[dept_col].nunique()

                print(f"\n{'='*80}")
                print("OVERALL SUMMARY:")
                print(f"{'='*80}")
                print(f"Total 5-Year Spend:     ${total_spend:,.2f}")
                print(f"Annual Run Rate:        ${total_spend/5:,.2f}")
                print(f"Total Transactions:     {total_transactions}")
                print(f"Unique Vendors:         {unique_vendors}")
                print(f"Unique Departments:     {unique_depts}")
                print(f"Avg Transaction Size:   ${total_spend/total_transactions:,.2f}")
                print(f"{'='*80}")

                # Save detailed data for workbook creation
                output_file = Path("/workspaces/Spend_Analysis_and_Categorization/spend-analysis/data/processed/drinking_water_detailed_data.csv")
                df_water_exhibit.to_csv(output_file, index=False)
                print(f"\n✓ Saved detailed data to: {output_file}")

                # Save pivot tables
                pivot_spend.to_csv(Path("/workspaces/Spend_Analysis_and_Categorization/spend-analysis/data/processed/drinking_water_vendor_dept_matrix.csv"))
                print(f"✓ Saved vendor-department matrix")

                vendor_summary.to_csv(Path("/workspaces/Spend_Analysis_and_Categorization/spend-analysis/data/processed/drinking_water_vendor_summary.csv"))
                dept_summary.to_csv(Path("/workspaces/Spend_Analysis_and_Categorization/spend-analysis/data/processed/drinking_water_department_summary.csv"))
                print(f"✓ Saved summary tables")

    except Exception as e:
        print(f"Note: Could not read main Exhibit B file: {e}")

print("\nData extraction complete.")
