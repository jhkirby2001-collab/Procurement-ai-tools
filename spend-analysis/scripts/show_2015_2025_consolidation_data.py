"""
Show consolidation scenario data for contracts awarded 2015-2025
"""

import pandas as pd
import numpy as np

# Load data
input_path = "/workspaces/Spend_Analysis_and_Categorization/spend-analysis/data/processed/landscaping_structured.xlsx"
df = pd.read_excel(input_path)

# Filter to 2015-2025
df['Award_Date'] = pd.to_datetime(df['Award_Date'])
df['Award_Year'] = df['Award_Date'].dt.year
df_filtered = df[(df['Award_Year'] >= 2015) & (df['Award_Year'] <= 2025)].copy()

print("\n" + "="*100)
print("LANDSCAPING CONSOLIDATION ANALYSIS: 2015-2025 CONTRACTS")
print("="*100)

# Calculate baseline metrics
total_spend = df_filtered['Award_Amount_Numeric'].sum()
total_vendors = df_filtered['Vendor name'].nunique()
total_contracts = len(df_filtered)
avg_contract_value = df_filtered['Award_Amount_Numeric'].mean()
median_contract_value = df_filtered['Award_Amount_Numeric'].median()

print("\n" + "="*100)
print("BASELINE METRICS (2015-2025)")
print("="*100)
print(f"\nTotal Contract Awards (2015-2025):  ${total_spend:,.2f}")
print(f"Number of Unique Vendors:            {total_vendors}")
print(f"Number of Contracts:                 {total_contracts}")
print(f"Average Contract Value:              ${avg_contract_value:,.2f}")
print(f"Median Contract Value:               ${median_contract_value:,.2f}")

# Vendor concentration
vendor_spend = df_filtered.groupby('Vendor name').agg({
    'Award_Amount_Numeric': 'sum',
    'Contract (PO) #': 'count'
}).reset_index()
vendor_spend.columns = ['Vendor', 'Total Spend', 'Contract Count']
vendor_spend = vendor_spend.sort_values('Total Spend', ascending=False)
vendor_spend['% of Total'] = (vendor_spend['Total Spend'] / total_spend * 100)
vendor_spend['Cumulative %'] = vendor_spend['% of Total'].cumsum()

print("\n" + "="*100)
print("TOP 10 VENDORS (2015-2025)")
print("="*100)
for idx, row in vendor_spend.head(10).iterrows():
    print(f"{row['Vendor'][:45]:<45}  ${row['Total Spend']:>15,.2f}  ({row['% of Total']:>5.1f}%)  {int(row['Contract Count'])} contracts")

top5_pct = vendor_spend.head(5)['% of Total'].sum()
print(f"\nTop 5 vendors control: {top5_pct:.1f}% of spend")

small_vendors = vendor_spend[vendor_spend['% of Total'] < 1]
small_vendors_count = len(small_vendors)
small_vendors_spend = small_vendors['Total Spend'].sum()
print(f"Small vendors (<1% each): {small_vendors_count} vendors = ${small_vendors_spend:,.2f} ({small_vendors_spend/total_spend*100:.1f}%)")

# Service categories
service_spend = df_filtered.groupby('Service_Category').agg({
    'Award_Amount_Numeric': 'sum',
    'Vendor name': 'nunique',
    'Contract (PO) #': 'count'
}).reset_index()
service_spend.columns = ['Service Category', 'Total Spend', 'Vendor Count', 'Contract Count']
service_spend = service_spend.sort_values('Total Spend', ascending=False)
service_spend['% of Total'] = (service_spend['Total Spend'] / total_spend * 100)

print("\n" + "="*100)
print("SERVICE CATEGORIES (2015-2025)")
print("="*100)
for idx, row in service_spend.iterrows():
    print(f"{row['Service Category'][:40]:<40}  ${row['Total Spend']:>15,.2f}  ({row['% of Total']:>5.1f}%)  {int(row['Vendor Count'])} vendors, {int(row['Contract Count'])} contracts")

# Year by year breakdown
yearly_spend = df_filtered.groupby('Award_Year').agg({
    'Award_Amount_Numeric': 'sum',
    'Contract (PO) #': 'count',
    'Vendor name': 'nunique'
}).reset_index()
yearly_spend.columns = ['Year', 'Total Awarded', 'Contract Count', 'Unique Vendors']

print("\n" + "="*100)
print("YEAR-BY-YEAR BREAKDOWN (2015-2025)")
print("="*100)
for idx, row in yearly_spend.iterrows():
    print(f"{int(row['Year'])}:  ${row['Total Awarded']:>15,.2f}  ({int(row['Contract Count'])} contracts, {int(row['Unique Vendors'])} vendors)")

print("\n" + "="*100)
print("CONSOLIDATION SCENARIOS")
print("="*100)

# Industry benchmarks
ADMIN_COST_PER_VENDOR = 50000
INFLATION_RATE = 0.045

# Calculate scenarios
scenarios = []

# Conservative: 26% reduction
conservative_target = int(total_vendors * 0.74)  # Keep 74%
conservative_eliminated = total_vendors - conservative_target
conservative_volume_savings = total_spend * 0.04
conservative_admin_savings = conservative_eliminated * ADMIN_COST_PER_VENDOR
conservative_total = conservative_volume_savings + conservative_admin_savings

scenarios.append({
    'Scenario': 'Conservative',
    'Current Vendors': total_vendors,
    'Target Vendors': conservative_target,
    'Vendors Eliminated': conservative_eliminated,
    'Reduction %': 26,
    'Volume Savings': conservative_volume_savings,
    'Admin Savings': conservative_admin_savings,
    'Total Savings': conservative_total
})

# Moderate: 45% reduction
moderate_target = int(total_vendors * 0.55)  # Keep 55%
moderate_eliminated = total_vendors - moderate_target
moderate_volume_savings = total_spend * 0.08
moderate_admin_savings = moderate_eliminated * ADMIN_COST_PER_VENDOR
moderate_total = moderate_volume_savings + moderate_admin_savings

scenarios.append({
    'Scenario': 'Moderate (RECOMMENDED)',
    'Current Vendors': total_vendors,
    'Target Vendors': moderate_target,
    'Vendors Eliminated': moderate_eliminated,
    'Reduction %': 45,
    'Volume Savings': moderate_volume_savings,
    'Admin Savings': moderate_admin_savings,
    'Total Savings': moderate_total
})

# Aggressive: 65% reduction
aggressive_target = int(total_vendors * 0.35)  # Keep 35%
aggressive_eliminated = total_vendors - aggressive_target
aggressive_volume_savings = total_spend * 0.13
aggressive_admin_savings = aggressive_eliminated * ADMIN_COST_PER_VENDOR
aggressive_total = aggressive_volume_savings + aggressive_admin_savings

scenarios.append({
    'Scenario': 'Aggressive',
    'Current Vendors': total_vendors,
    'Target Vendors': aggressive_target,
    'Vendors Eliminated': aggressive_eliminated,
    'Reduction %': 65,
    'Volume Savings': aggressive_volume_savings,
    'Admin Savings': aggressive_admin_savings,
    'Total Savings': aggressive_total
})

print("\n{:<25} {:<12} {:<12} {:<12} {:<12} {:<18} {:<18} {:<20}".format(
    'Scenario', 'Current', 'Target', 'Eliminated', 'Reduction %', 'Volume Savings', 'Admin Savings', 'Total Savings'
))
print("-" * 140)

for s in scenarios:
    print("{:<25} {:<12} {:<12} {:<12} {:<12} ${:<17,.2f} ${:<17,.2f} ${:<19,.2f}".format(
        s['Scenario'],
        s['Current Vendors'],
        s['Target Vendors'],
        s['Vendors Eliminated'],
        f"{s['Reduction %']}%",
        s['Volume Savings'],
        s['Admin Savings'],
        s['Total Savings']
    ))

print("\n" + "="*100)
print("COST AVOIDANCE THROUGH STRATEGIC SOURCING (3-YEAR)")
print("="*100)

# Calculate cost avoidance
cost_avoidance_data = []
for year in range(1, 4):
    inflated_cost = total_spend * ((1 + INFLATION_RATE) ** year)
    cost_with_lock = total_spend
    avoidance = inflated_cost - cost_with_lock
    cost_avoidance_data.append({
        'Year': year,
        'Without Strategic Sourcing': inflated_cost,
        'With Rate Lock': cost_with_lock,
        'Cost Avoidance': avoidance
    })

cumulative_avoidance = 0
print("\n{:<10} {:<25} {:<25} {:<25} {:<25}".format(
    'Year', 'Without Sourcing', 'With Rate Lock', 'Annual Avoidance', 'Cumulative Avoidance'
))
print("-" * 110)

for data in cost_avoidance_data:
    cumulative_avoidance += data['Cost Avoidance']
    print("{:<10} ${:<24,.2f} ${:<24,.2f} ${:<24,.2f} ${:<24,.2f}".format(
        f"Year {data['Year']}",
        data['Without Strategic Sourcing'],
        data['With Rate Lock'],
        data['Cost Avoidance'],
        cumulative_avoidance
    ))

total_3yr_avoidance = cumulative_avoidance

print("\n" + "="*100)
print("TOTAL FINANCIAL IMPACT (MODERATE SCENARIO + 3-YEAR COST AVOIDANCE)")
print("="*100)

moderate_annual = moderate_total
three_year_savings = moderate_annual * 3
combined_benefit = three_year_savings + total_3yr_avoidance

print(f"\nModerate Scenario Annual Savings:           ${moderate_annual:,.2f}")
print(f"Moderate Scenario 3-Year Savings:           ${three_year_savings:,.2f}")
print(f"Strategic Sourcing 3-Year Cost Avoidance:   ${total_3yr_avoidance:,.2f}")
print(f"{'='*100}")
print(f"TOTAL 3-YEAR FINANCIAL BENEFIT:             ${combined_benefit:,.2f}")
print(f"{'='*100}")

print("\n" + "="*100)
print("KEY INSIGHTS")
print("="*100)
print(f"\n• This analysis covers {total_contracts} contracts awarded from 2015-2025")
print(f"• Total contract value over this period: ${total_spend:,.2f}")
print(f"• Current vendor base: {total_vendors} vendors")
print(f"• Top 5 vendors control {top5_pct:.1f}% of contract awards")
print(f"• {small_vendors_count} small vendors (<1% each) represent consolidation opportunity")
print(f"• Moderate consolidation could save ${moderate_annual:,.2f} annually")
print(f"• Combined with strategic sourcing: ${combined_benefit:,.2f} over 3 years")

print("\n" + "="*100)
print("DOES THIS LOOK CORRECT?")
print("="*100)
print("\nPlease review these numbers before I proceed with regenerating:")
print("  - Excel workbook with 8 sheets")
print("  - 7 professional visualizations")
print("  - Word document leadership guide")
print("\nAll will be based on 2015-2025 contract data shown above.")
print("="*100 + "\n")
