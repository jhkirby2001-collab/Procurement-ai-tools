"""
Check longer timeframe for active contracts (7-8 years)
"""

import pandas as pd
from datetime import datetime, timedelta

# Load data
input_path = "/workspaces/Spend_Analysis_and_Categorization/spend-analysis/data/processed/landscaping_structured.xlsx"
df = pd.read_excel(input_path)

df['Award_Date'] = pd.to_datetime(df['Award_Date'])
current_date = datetime.now()

print("\n" + "="*100)
print("CHECKING DIFFERENT TIMEFRAMES FOR ACTIVE CONTRACTS")
print("="*100)

# Check different lookback periods
timeframes = [
    (3, "3 years - typical contract minimum"),
    (5, "5 years - standard contract term"),
    (7, "7 years - extended contract term"),
    (10, "10 years - maximum reasonable lookback")
]

for years, description in timeframes:
    cutoff_date = current_date - timedelta(days=years*365)
    contracts_in_period = df[df['Award_Date'] >= cutoff_date]
    total_awarded = contracts_in_period['Award_Amount_Numeric'].sum()

    print(f"\n{description.upper()}")
    print(f"Cutoff Date: {cutoff_date.strftime('%Y-%m-%d')}")
    print(f"Contracts: {len(contracts_in_period)}")
    print(f"Total Awarded: ${total_awarded:,.2f}")
    print(f"Estimated Annual Spend (assuming {years}-year terms): ${total_awarded/years:,.2f}")

print("\n" + "="*100)
print("MAJOR CONTRACT YEARS (2018-2019)")
print("="*100)

# Look specifically at 2018-2019 which had large awards
major_years = df[df['Award_Date'].dt.year.isin([2018, 2019])].copy()
major_years = major_years.sort_values('Award_Amount_Numeric', ascending=False)

print(f"\nContracts Awarded in 2018-2019:")
print(f"Total Count: {len(major_years)}")
print(f"Total Awarded: ${major_years['Award_Amount_Numeric'].sum():,.2f}")

print("\nTop Contracts:")
for idx, row in major_years.head(10).iterrows():
    print(f"  {row['Award_Date'].strftime('%Y-%m-%d')}  {row['Vendor name'][:40]:<40}  ${row['Award_Amount_Numeric']:>15,.2f}  {row['Service_Category']}")

# If these are 5-7 year contracts, they'd still be active through 2023-2025
print("\n" + "="*100)
print("ACTIVE CONTRACT ESTIMATION")
print("="*100)

total_2018_2019 = major_years['Award_Amount_Numeric'].sum()
print(f"\nTotal Awarded 2018-2019: ${total_2018_2019:,.2f}")
print(f"\nIf these were 5-year contracts (2018-2023, 2019-2024):")
print(f"  Annual Spend from these contracts: ${total_2018_2019/5:,.2f}")

print(f"\nIf these were 7-year contracts (2018-2025, 2019-2026):")
print(f"  Annual Spend from these contracts: ${total_2018_2019/7:,.2f}")

# Plus the recent smaller contracts
recent_5yr = df[df['Award_Date'] >= (current_date - timedelta(days=5*365))]
recent_total = recent_5yr['Award_Amount_Numeric'].sum()

print(f"\n{'='*100}")
print("TOTAL CURRENT ANNUAL SPEND ESTIMATE")
print(f"{'='*100}")

# Scenario 1: 2018-2019 contracts were 7-year (still active) + new contracts
estimate_7yr = (total_2018_2019/7) + (recent_total/4)
print(f"\nScenario 1 (2018-2019 contracts = 7-year + new contracts = 4-year):")
print(f"  2018-2019 annual: ${total_2018_2019/7:,.2f}")
print(f"  Recent annual: ${recent_total/4:,.2f}")
print(f"  TOTAL ANNUAL SPEND: ${estimate_7yr:,.2f}")

# Scenario 2: 2018-2019 contracts were 5-year (mostly expired) + new contracts
estimate_5yr = (recent_total/4)
print(f"\nScenario 2 (2018-2019 contracts expired, only recent contracts active):")
print(f"  Recent contracts annual: ${recent_total/4:,.2f}")
print(f"  TOTAL ANNUAL SPEND: ${estimate_5yr:,.2f}")

print(f"\n{'='*100}")
print("RECOMMENDATION FOR ANALYSIS")
print(f"{'='*100}")

# Best estimate considering both scenarios
best_estimate = (estimate_7yr + estimate_5yr) / 2

print(f"\nBest Estimate (average of both scenarios): ${best_estimate:,.2f}/year")
print(f"\nOR use the 7-year lookback estimate: ${df[df['Award_Date'] >= (current_date - timedelta(days=7*365))]['Award_Amount_Numeric'].sum()/7:,.2f}/year")

print(f"\n{'='*100}")
