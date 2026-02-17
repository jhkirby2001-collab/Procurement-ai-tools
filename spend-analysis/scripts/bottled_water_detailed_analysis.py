import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

# Set style
sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (14, 8)

# File paths
file_path = "/workspaces/Spend_Analysis_and_Categorization/spend-analysis/data/raw/BottledWater raw data.xlsx"
output_dir = Path("/workspaces/Spend_Analysis_and_Categorization/spend-analysis/outputs/bottled_water_analysis")
output_dir.mkdir(parents=True, exist_ok=True)

print("="*100)
print("FOCUSED BOTTLED WATER CONTRACT ANALYSIS")
print("Sheets 1 & 2: SPC49756A-PO14954 and SPC82197E-PO30043")
print("="*100)

# ============================================================================
# LOAD AND CLEAN DATA
# ============================================================================
print("\n📥 Loading Data...")

# Sheet 1: SPC49756A-PO14954 (2007-2013)
df1_raw = pd.read_excel(file_path, sheet_name='SPC49756A-PO14954')
print(f"\nSheet 1 - SPC49756A-PO14954 (Ice Mountain):")
print(f"  Raw rows: {len(df1_raw)}")

# Sheet 2: SPC82197E-PO30043 (2014-2022)
df2_raw = pd.read_excel(file_path, sheet_name='SPC82197E-PO30043')
print(f"\nSheet 2 - SPC82197E-PO30043 (Bluetriton Brands):")
print(f"  Raw rows: {len(df2_raw)}")

# ============================================================================
# CLEAN SHEET 1
# ============================================================================
print("\n\n🧹 Cleaning Sheet 1 Data...")

# Get contract header info
contract1_info = df1_raw.iloc[0].to_dict()
print(f"\nContract 1 Information:")
print(f"  PO Number: {contract1_info['PO Number']}")
print(f"  Supplier: {contract1_info['Supplier']}")
print(f"  Description: {contract1_info['Description']}")
print(f"  Contract Limit: {contract1_info['Contract Limit']}")
print(f"  Matched Amount: {contract1_info['Matched Amount']}")
print(f"  Contract Period: {contract1_info['Begins On']} to {contract1_info['Expires On']}")

# Extract transaction details (skip headers and summary rows)
df1 = df1_raw[
    (df1_raw['Supplier'].notna()) &
    (df1_raw['Supplier'] != 'Supplier') &
    (df1_raw['Supplier'] != 'Req Item or Category') &
    (~df1_raw['Supplier'].astype(str).str.contains('Number of entries', na=False))
].copy()

# Clean amount columns
for col in ['Amount', 'Amount Agreed', 'Contract Limit']:
    if col in df1.columns:
        df1[col] = df1[col].astype(str).str.replace('$', '').str.replace(',', '').str.strip()
        df1[col] = pd.to_numeric(df1[col], errors='coerce')

# Filter to actual product rows
df1_products = df1[
    (df1['Description'].notna()) &
    (df1['Description'] != 'Description') &
    (df1['Description'].str.contains('WATER', na=False))
].copy()

print(f"\nCleaned data:")
print(f"  Product line items: {len(df1_products)}")
print(f"\nProduct categories found:")
for desc in df1_products['Description'].unique():
    print(f"  - {desc}")

# ============================================================================
# CLEAN SHEET 2
# ============================================================================
print("\n\n🧹 Cleaning Sheet 2 Data...")

# Get contract header info
contract2_info = df2_raw.iloc[0].to_dict()
print(f"\nContract 2 Information:")
print(f"  PO Number: {contract2_info['PO Number']}")
print(f"  Supplier: {contract2_info['Supplier']}")
print(f"  Description: {contract2_info['Description']}")
print(f"  Contract Limit: {contract2_info['Contract Limit']}")
print(f"  Matched Amount: {contract2_info['Matched Amount']}")
print(f"  Contract Period: {contract2_info['Begins On']} to {contract2_info['Expires On']}")

# Extract transaction details
df2 = df2_raw[
    (df2_raw['Supplier'].notna()) &
    (df2_raw['Supplier'] != 'Supplier') &
    (df2_raw['Supplier'] != 'Req Item or Category') &
    (~df2_raw['Supplier'].astype(str).str.contains('Number of entries', na=False))
].copy()

# Clean amount columns
for col in ['Amount', 'Amount Agreed', 'Contract Limit']:
    if col in df2.columns:
        df2[col] = df2[col].astype(str).str.replace('$', '').str.replace(',', '').str.strip()
        df2[col] = pd.to_numeric(df2[col], errors='coerce')

# Filter to actual product rows
df2_products = df2[
    (df2['Description'].notna()) &
    (df2['Description'] != 'Description') &
    (df2['Description'].str.contains('WATER', na=False))
].copy()

print(f"\nCleaned data:")
print(f"  Product line items: {len(df2_products)}")
print(f"\nProduct categories found:")
for desc in df2_products['Description'].unique():
    print(f"  - {desc}")

# ============================================================================
# DETAILED ANALYSIS - CONTRACT 1
# ============================================================================
print("\n\n" + "="*100)
print("CONTRACT 1 DETAILED ANALYSIS (2007-2013)")
print("="*100)

# Product spend breakdown
product_spend_1 = df1_products.groupby('Description').agg({
    'Amount Agreed': ['sum', 'count', 'mean', 'min', 'max']
}).round(2)
product_spend_1.columns = ['Total_Spend', 'Count', 'Avg_Order', 'Min_Order', 'Max_Order']
product_spend_1['Pct_of_Total'] = (product_spend_1['Total_Spend'] / product_spend_1['Total_Spend'].sum() * 100).round(2)
product_spend_1 = product_spend_1.sort_values('Total_Spend', ascending=False)

print("\n📊 Product Spend Breakdown:")
print(product_spend_1.to_string())

# Department analysis
dept_spend_1 = df1_products.groupby('Amount').agg({
    'Amount Agreed': ['sum', 'count']
})
dept_spend_1.columns = ['Total_Spend', 'Count']
dept_spend_1 = dept_spend_1.sort_values('Total_Spend', ascending=False)

print(f"\n📍 Department Distribution:")
print(f"  Unique departments: {df1_products['Amount'].nunique()}")
print(f"  Total orders: {len(df1_products)}")

# Summary metrics
total_spend_1 = product_spend_1['Total_Spend'].sum()
avg_order_1 = df1_products['Amount Agreed'].mean()
median_order_1 = df1_products['Amount Agreed'].median()

print(f"\n💰 Financial Summary:")
print(f"  Total Contract Value: ${total_spend_1:,.2f}")
print(f"  Average Order Size: ${avg_order_1:,.2f}")
print(f"  Median Order Size: ${median_order_1:,.2f}")
print(f"  Contract Limit: ${float(contract1_info['Contract Limit']):,.2f}")
print(f"  Matched/Spent: ${float(contract1_info['Matched Amount']):,.2f}")
print(f"  Utilization Rate: {float(contract1_info['Matched Amount'])/float(contract1_info['Contract Limit'])*100:.2f}%")

# ============================================================================
# DETAILED ANALYSIS - CONTRACT 2
# ============================================================================
print("\n\n" + "="*100)
print("CONTRACT 2 DETAILED ANALYSIS (2014-2022)")
print("="*100)

# Product spend breakdown
product_spend_2 = df2_products.groupby('Description').agg({
    'Amount Agreed': ['sum', 'count', 'mean', 'min', 'max']
}).round(2)
product_spend_2.columns = ['Total_Spend', 'Count', 'Avg_Order', 'Min_Order', 'Max_Order']
product_spend_2['Pct_of_Total'] = (product_spend_2['Total_Spend'] / product_spend_2['Total_Spend'].sum() * 100).round(2)
product_spend_2 = product_spend_2.sort_values('Total_Spend', ascending=False)

print("\n📊 Product Spend Breakdown:")
print(product_spend_2.to_string())

# Department analysis
dept_spend_2 = df2_products.groupby('Amount').agg({
    'Amount Agreed': ['sum', 'count']
})
dept_spend_2.columns = ['Total_Spend', 'Count']
dept_spend_2 = dept_spend_2.sort_values('Total_Spend', ascending=False)

print(f"\n📍 Department Distribution:")
print(f"  Unique departments: {df2_products['Amount'].nunique()}")
print(f"  Total orders: {len(df2_products)}")

# Summary metrics
total_spend_2 = product_spend_2['Total_Spend'].sum()
avg_order_2 = df2_products['Amount Agreed'].mean()
median_order_2 = df2_products['Amount Agreed'].median()

print(f"\n💰 Financial Summary:")
print(f"  Total Contract Value: ${total_spend_2:,.2f}")
print(f"  Average Order Size: ${avg_order_2:,.2f}")
print(f"  Median Order Size: ${median_order_2:,.2f}")
print(f"  Contract Limit: ${float(contract2_info['Contract Limit']):,.2f}")
print(f"  Matched/Spent: ${float(contract2_info['Matched Amount']):,.2f}")
print(f"  Utilization Rate: {float(contract2_info['Matched Amount'])/float(contract2_info['Contract Limit'])*100:.2f}%")

# ============================================================================
# COMPARATIVE ANALYSIS
# ============================================================================
print("\n\n" + "="*100)
print("COMPARATIVE ANALYSIS: CONTRACT 1 vs CONTRACT 2")
print("="*100)

print(f"\n📈 Contract Growth:")
print(f"  Contract Limit Growth: ${float(contract2_info['Contract Limit']) - float(contract1_info['Contract Limit']):,.2f} ({(float(contract2_info['Contract Limit'])/float(contract1_info['Contract Limit'])-1)*100:+.1f}%)")
print(f"  Actual Spend Growth: ${float(contract2_info['Matched Amount']) - float(contract1_info['Matched Amount']):,.2f} ({(float(contract2_info['Matched Amount'])/float(contract1_info['Matched Amount'])-1)*100:+.1f}%)")
print(f"  Duration: Contract 1 = 6.5 years, Contract 2 = 8.5 years")

print(f"\n📊 Product Mix Changes:")
# Compare common products
common_products = set(product_spend_1.index) & set(product_spend_2.index)
for product in common_products:
    spend1 = product_spend_1.loc[product, 'Total_Spend']
    spend2 = product_spend_2.loc[product, 'Total_Spend']
    pct1 = product_spend_1.loc[product, 'Pct_of_Total']
    pct2 = product_spend_2.loc[product, 'Pct_of_Total']
    print(f"\n  {product[:60]}...")
    print(f"    Contract 1: ${spend1:,.2f} ({pct1:.1f}%)")
    print(f"    Contract 2: ${spend2:,.2f} ({pct2:.1f}%)")
    print(f"    Change: ${spend2-spend1:,.2f} ({(spend2/spend1-1)*100:+.1f}%)")

# ============================================================================
# CREATE DETAILED VISUALIZATIONS
# ============================================================================
print("\n\n" + "="*100)
print("GENERATING BOTTLED WATER VISUALIZATIONS")
print("="*100)

# Chart 1: Product Mix Comparison - Detailed
print("\n1. Creating detailed product mix comparison...")
fig, axes = plt.subplots(2, 2, figsize=(18, 14))

# Contract 1 - Spend
ax1 = axes[0, 0]
colors1 = sns.color_palette("Blues_r", len(product_spend_1))
bars1 = ax1.barh(range(len(product_spend_1)), product_spend_1['Total_Spend'], color=colors1, edgecolor='black')
ax1.set_yticks(range(len(product_spend_1)))
ax1.set_yticklabels([desc[:50] for desc in product_spend_1.index], fontsize=9)
ax1.set_xlabel('Total Spend ($)', fontweight='bold')
ax1.set_title('Contract 1 (2007-2013) - Product Spend', fontweight='bold', fontsize=12)
ax1.invert_yaxis()
for i, (bar, value, pct) in enumerate(zip(bars1, product_spend_1['Total_Spend'], product_spend_1['Pct_of_Total'])):
    ax1.text(value, bar.get_y() + bar.get_height()/2, f' ${value:,.0f} ({pct:.1f}%)',
            va='center', fontsize=9, fontweight='bold')

# Contract 2 - Spend
ax2 = axes[0, 1]
colors2 = sns.color_palette("Greens_r", len(product_spend_2))
bars2 = ax2.barh(range(len(product_spend_2)), product_spend_2['Total_Spend'], color=colors2, edgecolor='black')
ax2.set_yticks(range(len(product_spend_2)))
ax2.set_yticklabels([desc[:50] for desc in product_spend_2.index], fontsize=9)
ax2.set_xlabel('Total Spend ($)', fontweight='bold')
ax2.set_title('Contract 2 (2014-2022) - Product Spend', fontweight='bold', fontsize=12)
ax2.invert_yaxis()
for i, (bar, value, pct) in enumerate(zip(bars2, product_spend_2['Total_Spend'], product_spend_2['Pct_of_Total'])):
    ax2.text(value, bar.get_y() + bar.get_height()/2, f' ${value:,.0f} ({pct:.1f}%)',
            va='center', fontsize=9, fontweight='bold')

# Contract 1 - Order Frequency
ax3 = axes[1, 0]
bars3 = ax3.barh(range(len(product_spend_1)), product_spend_1['Count'], color=colors1, edgecolor='black', alpha=0.7)
ax3.set_yticks(range(len(product_spend_1)))
ax3.set_yticklabels([desc[:50] for desc in product_spend_1.index], fontsize=9)
ax3.set_xlabel('Number of Orders', fontweight='bold')
ax3.set_title('Contract 1 - Order Frequency', fontweight='bold', fontsize=12)
ax3.invert_yaxis()
for i, (bar, count) in enumerate(zip(bars3, product_spend_1['Count'])):
    ax3.text(count, bar.get_y() + bar.get_height()/2, f' {int(count)} orders',
            va='center', fontsize=9, fontweight='bold')

# Contract 2 - Order Frequency
ax4 = axes[1, 1]
bars4 = ax4.barh(range(len(product_spend_2)), product_spend_2['Count'], color=colors2, edgecolor='black', alpha=0.7)
ax4.set_yticks(range(len(product_spend_2)))
ax4.set_yticklabels([desc[:50] for desc in product_spend_2.index], fontsize=9)
ax4.set_xlabel('Number of Orders', fontweight='bold')
ax4.set_title('Contract 2 - Order Frequency', fontweight='bold', fontsize=12)
ax4.invert_yaxis()
for i, (bar, count) in enumerate(zip(bars4, product_spend_2['Count'])):
    ax4.text(count, bar.get_y() + bar.get_height()/2, f' {int(count)} orders',
            va='center', fontsize=9, fontweight='bold')

plt.suptitle('Bottled Water Product Mix: Detailed Comparison', fontsize=16, fontweight='bold', y=0.995)
plt.tight_layout()
plt.savefig(output_dir / 'BW_01_product_mix_detailed.png', dpi=300, bbox_inches='tight')
plt.close()
print("   Saved: BW_01_product_mix_detailed.png")

# Chart 2: Contract Performance Comparison
print("\n2. Creating contract performance comparison...")
fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(16, 12))

# Contract values comparison
contracts = ['Contract 1\n(2007-2013)', 'Contract 2\n(2014-2022)']
limits = [float(contract1_info['Contract Limit']), float(contract2_info['Contract Limit'])]
matched = [float(contract1_info['Matched Amount']), float(contract2_info['Matched Amount'])]
agreed = [total_spend_1, total_spend_2]

x = np.arange(len(contracts))
width = 0.25

bars1 = ax1.bar(x - width, [l/1000 for l in limits], width, label='Contract Limit', color='lightblue', edgecolor='black')
bars2 = ax1.bar(x, [m/1000 for m in matched], width, label='Matched/Actual', color='blue', edgecolor='black')
bars3 = ax1.bar(x + width, [a/1000 for a in agreed], width, label='Amount Agreed', color='darkblue', edgecolor='black')

ax1.set_ylabel('Amount ($K)', fontweight='bold')
ax1.set_title('Contract Financial Comparison', fontweight='bold', fontsize=12)
ax1.set_xticks(x)
ax1.set_xticklabels(contracts)
ax1.legend()
ax1.grid(axis='y', alpha=0.3)

for bars in [bars1, bars2, bars3]:
    for bar in bars:
        height = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width()/2., height, f'${height:.0f}K',
                ha='center', va='bottom', fontsize=9, fontweight='bold')

# Utilization rates
utilization = [(m/l*100) for l, m in zip(limits, matched)]
colors_util = ['green' if u >= 85 else 'orange' if u >= 70 else 'red' for u in utilization]
bars = ax2.bar(x, utilization, color=colors_util, edgecolor='black', linewidth=2, alpha=0.7)
ax2.set_ylabel('Utilization Rate (%)', fontweight='bold')
ax2.set_title('Contract Utilization Efficiency', fontweight='bold', fontsize=12)
ax2.set_xticks(x)
ax2.set_xticklabels(contracts)
ax2.set_ylim([0, 100])
ax2.axhline(y=85, color='green', linestyle='--', linewidth=1.5, alpha=0.5, label='Target: 85%')
ax2.axhline(y=70, color='orange', linestyle='--', linewidth=1.5, alpha=0.5, label='Warning: 70%')
ax2.legend()
ax2.grid(axis='y', alpha=0.3)

for bar, util in zip(bars, utilization):
    height = bar.get_height()
    ax2.text(bar.get_x() + bar.get_width()/2., height, f'{util:.1f}%',
            ha='center', va='bottom', fontsize=11, fontweight='bold')

# Order patterns
order_counts = [len(df1_products), len(df2_products)]
avg_orders = [avg_order_1, avg_order_2]

bars = ax3.bar(x, order_counts, color=['steelblue', 'seagreen'], edgecolor='black', linewidth=2, alpha=0.7)
ax3.set_ylabel('Number of Orders', fontweight='bold')
ax3.set_title('Total Order Volume', fontweight='bold', fontsize=12)
ax3.set_xticks(x)
ax3.set_xticklabels(contracts)
ax3.grid(axis='y', alpha=0.3)

for bar, count in zip(bars, order_counts):
    height = bar.get_height()
    ax3.text(bar.get_x() + bar.get_width()/2., height, f'{int(count)}\norders',
            ha='center', va='bottom', fontsize=10, fontweight='bold')

# Average order size
bars = ax4.bar(x, [a/1000 for a in avg_orders], color=['steelblue', 'seagreen'], edgecolor='black', linewidth=2, alpha=0.7)
ax4.set_ylabel('Average Order Size ($K)', fontweight='bold')
ax4.set_title('Average Order Size Comparison', fontweight='bold', fontsize=12)
ax4.set_xticks(x)
ax4.set_xticklabels(contracts)
ax4.grid(axis='y', alpha=0.3)

for bar, avg in zip(bars, avg_orders):
    height = bar.get_height()
    ax4.text(bar.get_x() + bar.get_width()/2., height, f'${height:.2f}K',
            ha='center', va='bottom', fontsize=10, fontweight='bold')

plt.suptitle('Bottled Water Contract Performance Analysis', fontsize=16, fontweight='bold', y=0.995)
plt.tight_layout()
plt.savefig(output_dir / 'BW_02_contract_performance.png', dpi=300, bbox_inches='tight')
plt.close()
print("   Saved: BW_02_contract_performance.png")

# Chart 3: Cost per Unit Analysis
print("\n3. Creating cost efficiency analysis...")

# Create simplified product categories for comparison
product_map = {
    'WATER, BOTTLED - DRINKING WATER IN 5 GAL. BOTTLES.': '5-Gal Drinking',
    'WATER, BOTTLED - 16.9 OZ. BOTTLES 24 BOTTLES/CASE': '16.9oz Case',
    'WATER, BOTTLED - DISTILLED WATER IN 5 GAL. BOTTLES.': '5-Gal Distilled',
    'WATER, BOTTLED - DISTILLED WATER IN 1 GAL. BOTTLES.': '1-Gal Distilled',
    'RENTAL OF WATER COOLERS. - ELECTRIC HOT AND COLD COOLER W/O REFRIGERATOR.  MONTHLY RENTAL CHARGE': 'Cooler Rental (No Fridge)',
    'RENTAL OF WATER COOLERS. - ELECTRIC HOT AND COLD COOLER W/REFRIGERATOR. - MONTHLY RENTAL CHARGE': 'Cooler Rental (w/Fridge)'
}

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 8))

# Contract 1 average order size by product
avg_1 = product_spend_1['Avg_Order'].sort_values(ascending=True)
colors = sns.color_palette("viridis", len(avg_1))
bars = ax1.barh(range(len(avg_1)), avg_1, color=colors, edgecolor='black', alpha=0.8)
ax1.set_yticks(range(len(avg_1)))
ax1.set_yticklabels([product_map.get(p, p[:40]) for p in avg_1.index], fontsize=9)
ax1.set_xlabel('Average Order Size ($)', fontweight='bold')
ax1.set_title('Contract 1: Avg Cost per Order by Product', fontweight='bold', fontsize=12)
ax1.invert_yaxis()

for bar, value in zip(bars, avg_1):
    ax1.text(value, bar.get_y() + bar.get_height()/2, f' ${value:,.0f}',
            va='center', fontsize=9, fontweight='bold')

# Contract 2 average order size by product
avg_2 = product_spend_2['Avg_Order'].sort_values(ascending=True)
bars = ax2.barh(range(len(avg_2)), avg_2, color=colors[:len(avg_2)], edgecolor='black', alpha=0.8)
ax2.set_yticks(range(len(avg_2)))
ax2.set_yticklabels([product_map.get(p, p[:40]) for p in avg_2.index], fontsize=9)
ax2.set_xlabel('Average Order Size ($)', fontweight='bold')
ax2.set_title('Contract 2: Avg Cost per Order by Product', fontweight='bold', fontsize=12)
ax2.invert_yaxis()

for bar, value in zip(bars, avg_2):
    ax2.text(value, bar.get_y() + bar.get_height()/2, f' ${value:,.0f}',
            va='center', fontsize=9, fontweight='bold')

plt.suptitle('Bottled Water Cost Efficiency Analysis', fontsize=16, fontweight='bold', y=0.98)
plt.tight_layout()
plt.savefig(output_dir / 'BW_03_cost_efficiency.png', dpi=300, bbox_inches='tight')
plt.close()
print("   Saved: BW_03_cost_efficiency.png")

print("\n" + "="*100)
print("BOTTLED WATER ANALYSIS COMPLETE")
print("="*100)
print(f"\nOutput directory: {output_dir}")
print("\nGenerated files:")
print("  - BW_01_product_mix_detailed.png")
print("  - BW_02_contract_performance.png")
print("  - BW_03_cost_efficiency.png")
print("="*100)
