"""
Identify active landscaping contracts and calculate current annual spend
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# Load data
input_path = "/workspaces/Spend_Analysis_and_Categorization/spend-analysis/data/processed/landscaping_structured.xlsx"
df = pd.read_excel(input_path)

print("\n" + "="*100)
print("ANALYZING CONTRACT DATA STRUCTURE")
print("="*100)

print("\nAvailable Columns:")
for col in df.columns:
    print(f"  - {col}")

print("\n" + "="*100)
print("SAMPLE CONTRACTS (First 10)")
print("="*100)

# Show sample data
sample_cols = ['Contract (PO) #', 'Award_Date', 'Award_Amount_Numeric', 'Vendor name', 'Service_Category', 'Notes']
available_cols = [col for col in sample_cols if col in df.columns]
print("\n", df[available_cols].head(10).to_string())

print("\n" + "="*100)
print("CONTRACT DATE ANALYSIS")
print("="*100)

# Convert Award_Date to datetime
df['Award_Date'] = pd.to_datetime(df['Award_Date'])
df['Award_Year'] = df['Award_Date'].dt.year

# Current date
current_date = datetime.now()
print(f"\nCurrent Date: {current_date.strftime('%Y-%m-%d')}")

# Check if there are any end date or term columns
print("\nChecking for contract term indicators...")
term_cols = [col for col in df.columns if any(keyword in col.lower() for keyword in ['end', 'expir', 'term', 'duration', 'year'])]
if term_cols:
    print(f"Found potential term columns: {term_cols}")
    for col in term_cols:
        print(f"\n{col} sample values:")
        print(df[col].value_counts().head(10))
else:
    print("No explicit term/end date columns found")

print("\n" + "="*100)
print("IDENTIFYING ACTIVE CONTRACTS")
print("="*100)

# Strategy 1: Contracts awarded in last 5 years (typical multi-year contract length)
five_years_ago = current_date - timedelta(days=5*365)
recent_contracts = df[df['Award_Date'] >= five_years_ago].copy()

print(f"\nContracts Awarded in Last 5 Years (since {five_years_ago.strftime('%Y-%m-%d')}):")
print(f"  Count: {len(recent_contracts)}")
print(f"  Total Award Amount: ${recent_contracts['Award_Amount_Numeric'].sum():,.2f}")

# Group by vendor for recent contracts
recent_by_vendor = recent_contracts.groupby('Vendor name').agg({
    'Award_Amount_Numeric': 'sum',
    'Contract (PO) #': 'count',
    'Award_Date': ['min', 'max']
}).reset_index()
recent_by_vendor.columns = ['Vendor', 'Total Award', 'Contract Count', 'First Award', 'Last Award']
recent_by_vendor = recent_by_vendor.sort_values('Total Award', ascending=False)

print("\nTop 10 Vendors (Last 5 Years):")
for idx, row in recent_by_vendor.head(10).iterrows():
    print(f"  {row['Vendor'][:50]:<50} ${row['Total Award']:>15,.2f}  ({int(row['Contract Count'])} contracts)")

print("\n" + "="*100)
print("ESTIMATING ANNUAL SPEND FROM ACTIVE CONTRACTS")
print("="*100)

# For multi-year contracts, we need to estimate the annual spend
# Typical landscaping contracts are 3-5 years
print("\nApproach 1: Divide recent contract awards by assumed contract terms")
print("\nAssuming average 3-year contract terms:")
assumed_term_years = 3
estimated_annual_from_3yr = recent_contracts['Award_Amount_Numeric'].sum() / assumed_term_years
print(f"  Estimated Annual Spend: ${estimated_annual_from_3yr:,.2f}")

print("\nAssuming average 5-year contract terms:")
assumed_term_years = 5
estimated_annual_from_5yr = recent_contracts['Award_Amount_Numeric'].sum() / assumed_term_years
print(f"  Estimated Annual Spend: ${estimated_annual_from_5yr:,.2f}")

print("\n" + "="*100)
print("APPROACH 2: CONTRACTS AWARDED IN LAST 3 YEARS")
print("="*100)

three_years_ago = current_date - timedelta(days=3*365)
very_recent_contracts = df[df['Award_Date'] >= three_years_ago].copy()

print(f"\nContracts Awarded in Last 3 Years (since {three_years_ago.strftime('%Y-%m-%d')}):")
print(f"  Count: {len(very_recent_contracts)}")
print(f"  Total Award Amount: ${very_recent_contracts['Award_Amount_Numeric'].sum():,.2f}")

# If these are 3-year contracts, annual spend would be:
if len(very_recent_contracts) > 0:
    annual_from_recent_3yr = very_recent_contracts['Award_Amount_Numeric'].sum() / 3
    print(f"\n  If 3-year contracts, Annual Spend: ${annual_from_recent_3yr:,.2f}")
    print(f"  If 5-year contracts, Annual Spend: ${very_recent_contracts['Award_Amount_Numeric'].sum() / 5:,.2f}")

print("\n" + "="*100)
print("APPROACH 3: ANALYZE CONTRACT NOTES/SPECIFICATIONS")
print("="*100)

# Check if Notes or other fields contain contract term information
if 'Notes' in df.columns:
    print("\nSample Notes (looking for term information):")
    notes_with_content = df[df['Notes'].notna()]['Notes'].head(20)
    for i, note in enumerate(notes_with_content, 1):
        if 'year' in str(note).lower() or 'term' in str(note).lower():
            print(f"  {i}. {note}")

# Check Specification field
if 'Specification #' in df.columns or 'Embedded_Spec' in df.columns:
    spec_col = 'Embedded_Spec' if 'Embedded_Spec' in df.columns else 'Specification #'
    print(f"\nSample {spec_col} (looking for term information):")
    specs_with_content = df[df[spec_col].notna()][spec_col].head(20)
    for i, spec in enumerate(specs_with_content, 1):
        spec_str = str(spec).lower()
        if any(keyword in spec_str for keyword in ['year', 'term', 'annual', 'duration']):
            print(f"  {i}. {spec}")

print("\n" + "="*100)
print("RECOMMENDATION")
print("="*100)

# Provide best estimate
total_5yr_awards = recent_contracts['Award_Amount_Numeric'].sum()
estimated_annual_conservative = total_5yr_awards / 5
estimated_annual_moderate = total_5yr_awards / 4
estimated_annual_aggressive = total_5yr_awards / 3

print(f"\nBased on contracts awarded in last 5 years (${total_5yr_awards:,.2f}):")
print(f"\n  Conservative Estimate (5-year avg contract term): ${estimated_annual_conservative:,.2f}/year")
print(f"  Moderate Estimate (4-year avg contract term):     ${estimated_annual_moderate:,.2f}/year")
print(f"  Aggressive Estimate (3-year avg contract term):   ${estimated_annual_aggressive:,.2f}/year")

print(f"\n{'='*100}")
print("RECOMMENDED ANNUAL SPEND FOR ANALYSIS")
print(f"{'='*100}")
print(f"\nUse MODERATE ESTIMATE: ${estimated_annual_moderate:,.2f} per year")
print("\nRationale:")
print("  - Typical government landscaping contracts are 3-5 years")
print("  - 4-year average is reasonable middle ground")
print("  - This represents current active contract portfolio")
print(f"\n{'='*100}")

# Show details of recent contracts
print("\n" + "="*100)
print("DETAILED LIST: CONTRACTS AWARDED IN LAST 5 YEARS")
print("="*100)

recent_detailed = recent_contracts[['Award_Date', 'Vendor name', 'Award_Amount_Numeric', 'Service_Category', 'Contract (PO) #']].copy()
recent_detailed = recent_detailed.sort_values('Award_Date', ascending=False)
recent_detailed['Award_Date_Str'] = recent_detailed['Award_Date'].dt.strftime('%Y-%m-%d')

print("\n")
for idx, row in recent_detailed.iterrows():
    print(f"{row['Award_Date_Str']}  {row['Vendor name'][:35]:<35}  ${row['Award_Amount_Numeric']:>15,.2f}  {row['Service_Category'][:30]}")
