#!/usr/bin/env python3
"""
Extract complete vendor-department matrix from the CPO presentation
"""

import pandas as pd
from pathlib import Path

file_path = Path("/workspaces/Spend_Analysis_and_Categorization/spend-analysis/outputs/EXHIBIT_B_CPO_PRESENTATION.xlsx")

print("=" * 80)
print("EXTRACTING VENDOR-DEPARTMENT MATRIX FOR PILOT BUSINESS CASE")
print("=" * 80)
print()

# Read the main Exhibit B data with the correct filter
exhibit_b = pd.read_excel(
    Path("/workspaces/Spend_Analysis_and_Categorization/spend-analysis/data/raw/Exhibit B Report Jan 2021- Jan 2026.xlsx"),
    sheet_name=0
)

# Filter for the 4 verified drinking water vendors only
water_vendors = ['DS Services of America, Inc.', 'BLUETRITON BRANDS INC', 'HINCKLEY SPRINGS', 'BREWSMART BEVERAGE']
vendor_filter = exhibit_b['Vendor Name'].str.upper().isin([v.upper() for v in water_vendors])
df_water = exhibit_b[vendor_filter].copy()

print(f"Total transactions: {len(df_water)}")
print(f"Total spend: ${df_water['Requisition Amount'].sum():,.2f}")
print(f"Unique vendors: {df_water['Vendor Name'].nunique()}")
print(f"Unique departments: {df_water['Department Name'].nunique()}")
print()

# Create vendor-department spend matrix
print("=" * 80)
print("VENDOR-DEPARTMENT SPEND MATRIX")
print("=" * 80)
pivot_spend = pd.pivot_table(
    df_water,
    values='Requisition Amount',
    index='Department Name',
    columns='Vendor Name',
    aggfunc='sum',
    fill_value=0
)
pivot_spend = pivot_spend.sort_values('Department Name')
pivot_spend['TOTAL'] = pivot_spend.sum(axis=1)
pivot_spend = pivot_spend.sort_values('TOTAL', ascending=False)

print(pivot_spend.round(2))
print()

# Create vendor-department transaction count matrix
print("=" * 80)
print("VENDOR-DEPARTMENT TRANSACTION COUNT")
print("=" * 80)
pivot_count = pd.pivot_table(
    df_water,
    values='Requisition Amount',
    index='Department Name',
    columns='Vendor Name',
    aggfunc='count',
    fill_value=0
)
pivot_count = pivot_count.loc[pivot_spend.index]  # Same order as spend matrix
pivot_count['TOTAL'] = pivot_count.sum(axis=1)

print(pivot_count.astype(int))
print()

# Vendor summary
print("=" * 80)
print("VENDOR SUMMARY")
print("=" * 80)
vendor_summary = df_water.groupby('Vendor Name').agg({
    'Requisition Amount': ['sum', 'count'],
    'Department Name': 'nunique'
}).round(2)
vendor_summary.columns = ['Total Spend', 'Transactions', 'Departments Served']
vendor_summary['% of Total'] = (vendor_summary['Total Spend'] / vendor_summary['Total Spend'].sum() * 100).round(1)
vendor_summary = vendor_summary.sort_values('Total Spend', ascending=False)
vendor_summary['Total Spend'] = vendor_summary['Total Spend'].apply(lambda x: f"${x:,.2f}")
print(vendor_summary)
print()

# Department summary
print("=" * 80)
print("DEPARTMENT SUMMARY")
print("=" * 80)
dept_summary = df_water.groupby('Department Name').agg({
    'Requisition Amount': ['sum', 'count'],
    'Vendor Name': 'nunique'
}).round(2)
dept_summary.columns = ['Total Spend', 'Transactions', 'Vendors Used']
dept_summary = dept_summary.sort_values('Total Spend', ascending=False)

print(dept_summary)
print()

# Departments using multiple vendors (fragmentation alert)
multi_vendor = dept_summary[dept_summary['Vendors Used'] > 1].copy()
print("=" * 80)
print("⚠️  DEPARTMENTS USING MULTIPLE VENDORS (FRAGMENTATION)")
print("=" * 80)
for dept in multi_vendor.index:
    dept_data = df_water[df_water['Department Name'] == dept]
    vendors = dept_data['Vendor Name'].unique()
    print(f"\n{dept}:")
    print(f"  Vendors Used: {len(vendors)} - {', '.join(vendors)}")
    print(f"  Total Spend: ${dept_data['Requisition Amount'].sum():,.2f}")
    print(f"  Transactions: {len(dept_data)}")

    # Show spend per vendor
    for vendor in vendors:
        vendor_spend = dept_data[dept_data['Vendor Name'] == vendor]['Requisition Amount'].sum()
        vendor_txns = len(dept_data[dept_data['Vendor Name'] == vendor])
        print(f"    - {vendor}: ${vendor_spend:,.2f} ({vendor_txns} transactions)")

print()

# Save the matrices as CSV for easy reference
pivot_spend.to_csv('/workspaces/Spend_Analysis_and_Categorization/spend-analysis/data/processed/pilot_vendor_dept_spend_matrix.csv')
pivot_count.to_csv('/workspaces/Spend_Analysis_and_Categorization/spend-analysis/data/processed/pilot_vendor_dept_count_matrix.csv')
vendor_summary.to_csv('/workspaces/Spend_Analysis_and_Categorization/spend-analysis/data/processed/pilot_vendor_summary.csv')
dept_summary.to_csv('/workspaces/Spend_Analysis_and_Categorization/spend-analysis/data/processed/pilot_dept_summary.csv')

print("\n✓ All matrices saved to data/processed/")
