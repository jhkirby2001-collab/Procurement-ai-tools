"""
Calculate the actual annual landscaping spend from the historical contract data
"""

import pandas as pd
import numpy as np

# Load data
input_path = "/workspaces/Spend_Analysis_and_Categorization/spend-analysis/data/processed/landscaping_structured.xlsx"
df = pd.read_excel(input_path)

print("\n" + "="*100)
print("CALCULATING ANNUAL LANDSCAPING SPEND")
print("="*100)

# Total historical spend
total_historical = df['Award_Amount_Numeric'].sum()
print(f"\nTotal Historical Spend (1993-2025): ${total_historical:,.2f}")
print(f"Number of Contracts: {len(df)}")
print(f"Date Range: {df['Award_Date'].min()} to {df['Award_Date'].max()}")

# Calculate spend by year
df['Award_Year'] = pd.to_datetime(df['Award_Date']).dt.year
yearly_spend = df.groupby('Award_Year').agg({
    'Award_Amount_Numeric': 'sum',
    'Contract (PO) #': 'count'
}).reset_index()
yearly_spend.columns = ['Year', 'Total Spend', 'Contract Count']
yearly_spend = yearly_spend.sort_values('Year')

print("\n" + "="*100)
print("SPEND BY YEAR")
print("="*100)
for _, row in yearly_spend.iterrows():
    print(f"{int(row['Year'])}: ${row['Total Spend']:>15,.2f}  ({int(row['Contract Count'])} contracts)")

# Calculate recent annual average (last 5 complete years)
recent_years = yearly_spend[yearly_spend['Year'].between(2019, 2024)]
if len(recent_years) > 0:
    avg_annual_recent = recent_years['Total Spend'].mean()
    print(f"\n{'='*100}")
    print(f"RECENT ANNUAL AVERAGE (2019-2024): ${avg_annual_recent:,.2f}")
    print(f"{'='*100}")
else:
    avg_annual_recent = yearly_spend['Total Spend'].mean()
    print(f"\n{'='*100}")
    print(f"OVERALL ANNUAL AVERAGE: ${avg_annual_recent:,.2f}")
    print(f"{'='*100}")

# Most recent complete year
most_recent_complete_year = yearly_spend[yearly_spend['Year'] == 2024]
if len(most_recent_complete_year) > 0:
    most_recent_spend = most_recent_complete_year['Total Spend'].values[0]
    print(f"\nMost Recent Complete Year (2024): ${most_recent_spend:,.2f}")

# Calculate median annual spend
median_annual = yearly_spend['Total Spend'].median()
print(f"Median Annual Spend: ${median_annual:,.2f}")

# Recommendation
print("\n" + "="*100)
print("RECOMMENDATION FOR ANALYSIS")
print("="*100)
print(f"\nUse Recent Annual Average as baseline: ${avg_annual_recent:,.2f}")
print("\nThis represents the current annual spending level for landscaping services")
print("and should be used for consolidation savings and cost avoidance calculations.")
print("="*100)
