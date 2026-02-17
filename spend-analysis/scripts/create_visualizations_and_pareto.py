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
plt.rcParams['font.size'] = 10

# File paths
file_path = "/workspaces/Spend_Analysis_and_Categorization/spend-analysis/data/raw/BottledWater raw data.xlsx"
output_dir = Path("/workspaces/Spend_Analysis_and_Categorization/spend-analysis/outputs/charts")
output_dir.mkdir(parents=True, exist_ok=True)

print("="*100)
print("BOTTLED WATER SPEND ANALYSIS - VISUALIZATIONS & PARETO ANALYSIS")
print("="*100)

# Read all sheets
excel_file = pd.ExcelFile(file_path)
dfs = {}

for sheet_name in excel_file.sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    dfs[sheet_name] = df

# ============================================================================
# CHART 1: EXHIBIT B - DEPARTMENT SPEND PARETO ANALYSIS
# ============================================================================
print("\n1. Creating Department Spend Pareto Analysis...")

df_exhibit = dfs['EXHIBIT B'].copy()
dept_spend = df_exhibit.groupby('Department Description')['Requisition Amount'].sum().sort_values(ascending=False)
dept_spend_pct = (dept_spend / dept_spend.sum() * 100)
dept_cumulative = dept_spend_pct.cumsum()

fig, ax1 = plt.subplots(figsize=(16, 10))

# Bar chart
x_pos = np.arange(len(dept_spend))
bars = ax1.bar(x_pos, dept_spend / 1000000, color='steelblue', alpha=0.7, edgecolor='black')
ax1.set_xlabel('Department', fontsize=12, fontweight='bold')
ax1.set_ylabel('Spend ($ Millions)', fontsize=12, fontweight='bold', color='steelblue')
ax1.tick_params(axis='y', labelcolor='steelblue')
ax1.set_xticks(x_pos)
ax1.set_xticklabels(dept_spend.index, rotation=45, ha='right', fontsize=9)

# Add value labels on bars
for i, (bar, value) in enumerate(zip(bars, dept_spend)):
    height = bar.get_height()
    ax1.text(bar.get_x() + bar.get_width()/2., height,
            f'${value/1000000:.1f}M',
            ha='center', va='bottom', fontsize=8, fontweight='bold')

# Cumulative line
ax2 = ax1.twinx()
line = ax2.plot(x_pos, dept_cumulative, color='red', marker='o', linewidth=2, markersize=6, label='Cumulative %')
ax2.set_ylabel('Cumulative Percentage (%)', fontsize=12, fontweight='bold', color='red')
ax2.tick_params(axis='y', labelcolor='red')
ax2.set_ylim([0, 105])
ax2.axhline(y=80, color='darkred', linestyle='--', linewidth=2, label='80% Threshold')

# Add grid
ax1.grid(axis='y', alpha=0.3)

plt.title('PARETO ANALYSIS: Department Spend Distribution\n(80/20 Rule)',
          fontsize=16, fontweight='bold', pad=20)

# Add Pareto annotation
pareto_count = len(dept_cumulative[dept_cumulative <= 80])
total_depts = len(dept_spend)
pareto_text = f'Top {pareto_count} departments ({pareto_count/total_depts*100:.1f}%) account for 80% of spend'
plt.text(0.5, 0.95, pareto_text, transform=fig.transFigure,
         ha='center', fontsize=12, bbox=dict(boxstyle='round', facecolor='yellow', alpha=0.7))

plt.tight_layout()
plt.savefig(output_dir / '01_department_pareto_analysis.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"   Saved: 01_department_pareto_analysis.png")

# ============================================================================
# CHART 2: TOP 10 DEPARTMENTS - PIE CHART
# ============================================================================
print("\n2. Creating Top 10 Departments Pie Chart...")

top10_depts = dept_spend.head(10)
other_spend = dept_spend[10:].sum()
pie_data = pd.concat([top10_depts, pd.Series({'Other Departments': other_spend})])

fig, ax = plt.subplots(figsize=(14, 10))
colors = sns.color_palette("Set3", len(pie_data))
wedges, texts, autotexts = ax.pie(pie_data, labels=pie_data.index, autopct='%1.1f%%',
                                    startangle=90, colors=colors, textprops={'fontsize': 10})

# Enhance text
for autotext in autotexts:
    autotext.set_color('white')
    autotext.set_fontweight('bold')
    autotext.set_fontsize(11)

plt.title('Top 10 Departments by Spend\n(Total: $273.69M)',
          fontsize=16, fontweight='bold', pad=20)
plt.tight_layout()
plt.savefig(output_dir / '02_top10_departments_pie.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"   Saved: 02_top10_departments_pie.png")

# ============================================================================
# CHART 3: JUSTIFICATION CATEGORY PARETO ANALYSIS
# ============================================================================
print("\n3. Creating Justification Category Pareto Analysis...")

just_spend = df_exhibit.groupby('Justification')['Requisition Amount'].agg(['sum', 'count']).sort_values('sum', ascending=False)
just_spend = just_spend[just_spend.index.notna()]
just_spend_pct = (just_spend['sum'] / just_spend['sum'].sum() * 100)
just_cumulative = just_spend_pct.cumsum()

fig, ax1 = plt.subplots(figsize=(16, 10))

# Bar chart
x_pos = np.arange(len(just_spend))
bars = ax1.bar(x_pos, just_spend['sum'] / 1000000, color='coral', alpha=0.7, edgecolor='black')
ax1.set_xlabel('Justification Category', fontsize=12, fontweight='bold')
ax1.set_ylabel('Spend ($ Millions)', fontsize=12, fontweight='bold', color='coral')
ax1.tick_params(axis='y', labelcolor='coral')

# Wrap long labels
wrapped_labels = []
for label in just_spend.index:
    if len(str(label)) > 50:
        words = str(label).split()
        wrapped = '\n'.join([' '.join(words[i:i+6]) for i in range(0, len(words), 6)])
        wrapped_labels.append(wrapped)
    else:
        wrapped_labels.append(str(label))

ax1.set_xticks(x_pos)
ax1.set_xticklabels(wrapped_labels, rotation=30, ha='right', fontsize=9)

# Add value labels
for i, (bar, value, count) in enumerate(zip(bars, just_spend['sum'], just_spend['count'])):
    height = bar.get_height()
    ax1.text(bar.get_x() + bar.get_width()/2., height,
            f'${value/1000000:.1f}M\n({count:,} req)',
            ha='center', va='bottom', fontsize=9, fontweight='bold')

# Cumulative line
ax2 = ax1.twinx()
line = ax2.plot(x_pos, just_cumulative, color='darkgreen', marker='D', linewidth=2.5, markersize=8)
ax2.set_ylabel('Cumulative Percentage (%)', fontsize=12, fontweight='bold', color='darkgreen')
ax2.tick_params(axis='y', labelcolor='darkgreen')
ax2.set_ylim([0, 105])
ax2.axhline(y=80, color='red', linestyle='--', linewidth=2, label='80% Threshold')

plt.title('PARETO ANALYSIS: Justification Categories\n(Problem: 67% Expired Contracts)',
          fontsize=16, fontweight='bold', pad=20)
plt.tight_layout()
plt.savefig(output_dir / '03_justification_pareto_analysis.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"   Saved: 03_justification_pareto_analysis.png")

# ============================================================================
# CHART 4: APPROVAL vs DENIAL ANALYSIS
# ============================================================================
print("\n4. Creating Approval vs Denial Analysis...")

approval_data = df_exhibit[df_exhibit['Approved or Denied'].notna()]
approval_spend = approval_data.groupby('Approved or Denied')['Requisition Amount'].agg(['sum', 'count'])

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))

# Spend breakdown
colors_approval = {'APPROVED': 'green', 'DENIED': 'red'}
bars1 = ax1.bar(approval_spend.index, approval_spend['sum'] / 1000000,
                color=[colors_approval[x] for x in approval_spend.index],
                alpha=0.7, edgecolor='black', linewidth=2)
ax1.set_ylabel('Total Spend ($ Millions)', fontsize=12, fontweight='bold')
ax1.set_xlabel('Decision Status', fontsize=12, fontweight='bold')
ax1.set_title('Spend by Approval Status', fontsize=14, fontweight='bold')
ax1.grid(axis='y', alpha=0.3)

for bar, value, count in zip(bars1, approval_spend['sum'], approval_spend['count']):
    height = bar.get_height()
    ax1.text(bar.get_x() + bar.get_width()/2., height,
            f'${value/1000000:.1f}M\n({count:,} req)\n{value/approval_spend["sum"].sum()*100:.1f}%',
            ha='center', va='bottom', fontsize=11, fontweight='bold')

# Count breakdown
bars2 = ax2.bar(approval_spend.index, approval_spend['count'],
                color=[colors_approval[x] for x in approval_spend.index],
                alpha=0.7, edgecolor='black', linewidth=2)
ax2.set_ylabel('Number of Requisitions', fontsize=12, fontweight='bold')
ax2.set_xlabel('Decision Status', fontsize=12, fontweight='bold')
ax2.set_title('Requisition Count by Approval Status', fontsize=14, fontweight='bold')
ax2.grid(axis='y', alpha=0.3)

for bar, count in zip(bars2, approval_spend['count']):
    height = bar.get_height()
    ax2.text(bar.get_x() + bar.get_width()/2., height,
            f'{count:,}\n{count/approval_spend["count"].sum()*100:.1f}%',
            ha='center', va='bottom', fontsize=11, fontweight='bold')

plt.suptitle('Approval vs Denial Analysis (Exhibit B)', fontsize=16, fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig(output_dir / '04_approval_denial_analysis.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"   Saved: 04_approval_denial_analysis.png")

# ============================================================================
# CHART 5: TOP 20 VENDORS - PARETO ANALYSIS
# ============================================================================
print("\n5. Creating Top 20 Vendors Pareto Analysis...")

vendor_spend = df_exhibit.groupby('Vendor Name')['Requisition Amount'].agg(['sum', 'count']).sort_values('sum', ascending=False).head(20)
vendor_spend_pct = (vendor_spend['sum'] / df_exhibit['Requisition Amount'].sum() * 100)
vendor_cumulative = vendor_spend_pct.cumsum()

fig, ax1 = plt.subplots(figsize=(16, 10))

x_pos = np.arange(len(vendor_spend))
bars = ax1.barh(x_pos, vendor_spend['sum'] / 1000000, color='purple', alpha=0.7, edgecolor='black')
ax1.set_ylabel('Vendor Name', fontsize=12, fontweight='bold')
ax1.set_xlabel('Spend ($ Millions)', fontsize=12, fontweight='bold', color='purple')
ax1.tick_params(axis='x', labelcolor='purple')
ax1.set_yticks(x_pos)
ax1.set_yticklabels(vendor_spend.index, fontsize=9)
ax1.invert_yaxis()

# Add value labels
for i, (bar, value, count) in enumerate(zip(bars, vendor_spend['sum'], vendor_spend['count'])):
    width = bar.get_width()
    ax1.text(width, bar.get_y() + bar.get_height()/2.,
            f' ${value/1000000:.1f}M ({count} req)',
            ha='left', va='center', fontsize=9, fontweight='bold')

# Cumulative line
ax2 = ax1.twiny()
line = ax2.plot(vendor_cumulative, x_pos, color='orange', marker='s', linewidth=2.5, markersize=7)
ax2.set_xlabel('Cumulative Percentage (%)', fontsize=12, fontweight='bold', color='orange')
ax2.tick_params(axis='x', labelcolor='orange')
ax2.set_xlim([0, 105])
ax2.axvline(x=80, color='red', linestyle='--', linewidth=2)

plt.title('PARETO ANALYSIS: Top 20 Vendors by Spend',
          fontsize=16, fontweight='bold', pad=20)
plt.tight_layout()
plt.savefig(output_dir / '05_top20_vendors_pareto.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"   Saved: 05_top20_vendors_pareto.png")

# ============================================================================
# CHART 6: CONTRACT COMPARISON (Sheets 1 & 2)
# ============================================================================
print("\n6. Creating Contract Comparison Analysis...")

# Process Sheet 1
df1 = dfs['SPC49756A-PO14954'].copy()
df1_clean = df1[df1['Description'].notna() & (df1['Description'] != 'Description')]
df1_clean = df1_clean[df1_clean['Description'].str.contains('WATER', na=False)]

# Convert amounts
for col in ['Amount Agreed']:
    if col in df1_clean.columns:
        df1_clean[col] = pd.to_numeric(df1_clean[col].astype(str).str.replace('$', '').str.replace(',', ''), errors='coerce')

product_spend_1 = df1_clean.groupby('Description')['Amount Agreed'].sum().sort_values(ascending=False)

# Process Sheet 2
df2 = dfs['SPC82197E-PO30043'].copy()
df2_clean = df2[df2['Description'].notna() & (df2['Description'] != 'Description')]
df2_clean = df2_clean[df2_clean['Description'].str.contains('WATER', na=False)]

for col in ['Amount Agreed']:
    if col in df2_clean.columns:
        df2_clean[col] = pd.to_numeric(df2_clean[col].astype(str).str.replace('$', '').str.replace(',', ''), errors='coerce')

product_spend_2 = df2_clean.groupby('Description')['Amount Agreed'].sum().sort_values(ascending=False)

# Create comparison chart
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 10))

# Contract 1 (2007-2013)
if len(product_spend_1) > 0:
    colors1 = sns.color_palette("Blues_r", len(product_spend_1))
    wedges1, texts1, autotexts1 = ax1.pie(product_spend_1, labels=product_spend_1.index, autopct='%1.1f%%',
                                            startangle=90, colors=colors1, textprops={'fontsize': 8})
    for autotext in autotexts1:
        autotext.set_color('white')
        autotext.set_fontweight('bold')
    ax1.set_title('Contract 1: SPC49756A-PO14954\n(2007-2013 | Ice Mountain)\nTotal: ${:.1f}K'.format(product_spend_1.sum()/1000),
                  fontsize=12, fontweight='bold')

# Contract 2 (2014-2022)
if len(product_spend_2) > 0:
    colors2 = sns.color_palette("Greens_r", len(product_spend_2))
    wedges2, texts2, autotexts2 = ax2.pie(product_spend_2, labels=product_spend_2.index, autopct='%1.1f%%',
                                            startangle=90, colors=colors2, textprops={'fontsize': 8})
    for autotext in autotexts2:
        autotext.set_color('white')
        autotext.set_fontweight('bold')
    ax2.set_title('Contract 2: SPC82197E-PO30043\n(2014-2022 | Bluetriton Brands)\nTotal: ${:.1f}K'.format(product_spend_2.sum()/1000),
                  fontsize=12, fontweight='bold')

plt.suptitle('Product Mix Comparison: Contract 1 vs Contract 2', fontsize=16, fontweight='bold', y=0.98)
plt.tight_layout()
plt.savefig(output_dir / '06_contract_comparison.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"   Saved: 06_contract_comparison.png")

# ============================================================================
# CHART 7: REQUISITION AMOUNT DISTRIBUTION
# ============================================================================
print("\n7. Creating Requisition Amount Distribution...")

fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(16, 12))

# Histogram
amounts = df_exhibit['Requisition Amount'].dropna()
ax1.hist(amounts[amounts < 100000] / 1000, bins=50, color='skyblue', edgecolor='black', alpha=0.7)
ax1.set_xlabel('Requisition Amount ($K)', fontsize=11, fontweight='bold')
ax1.set_ylabel('Frequency', fontsize=11, fontweight='bold')
ax1.set_title('Distribution of Requisition Amounts\n(Amounts < $100K)', fontsize=12, fontweight='bold')
ax1.grid(axis='y', alpha=0.3)

# Box plot
ax2.boxplot([amounts / 1000], vert=True, patch_artist=True,
            boxprops=dict(facecolor='lightgreen', alpha=0.7),
            medianprops=dict(color='red', linewidth=2),
            flierprops=dict(marker='o', markerfacecolor='red', markersize=3, alpha=0.5))
ax2.set_ylabel('Requisition Amount ($K)', fontsize=11, fontweight='bold')
ax2.set_title('Box Plot: Requisition Amount Distribution', fontsize=12, fontweight='bold')
ax2.grid(axis='y', alpha=0.3)

# Add statistics
stats_text = f'Mean: ${amounts.mean():,.0f}\nMedian: ${amounts.median():,.0f}\nStd Dev: ${amounts.std():,.0f}\nMax: ${amounts.max():,.0f}'
ax2.text(1.15, amounts.max()/1000 * 0.5, stats_text, fontsize=10,
         bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.7))

# Cumulative distribution
sorted_amounts = np.sort(amounts)
cumulative = np.arange(1, len(sorted_amounts) + 1) / len(sorted_amounts) * 100
ax3.plot(sorted_amounts / 1000, cumulative, color='purple', linewidth=2)
ax3.set_xlabel('Requisition Amount ($K)', fontsize=11, fontweight='bold')
ax3.set_ylabel('Cumulative Percentage (%)', fontsize=11, fontweight='bold')
ax3.set_title('Cumulative Distribution Function', fontsize=12, fontweight='bold')
ax3.grid(True, alpha=0.3)
ax3.set_xlim([0, 500])

# Log scale distribution
ax4.hist(amounts / 1000, bins=50, color='coral', edgecolor='black', alpha=0.7)
ax4.set_xlabel('Requisition Amount ($K)', fontsize=11, fontweight='bold')
ax4.set_ylabel('Frequency', fontsize=11, fontweight='bold')
ax4.set_title('Distribution (Log Scale)', fontsize=12, fontweight='bold')
ax4.set_yscale('log')
ax4.grid(True, alpha=0.3, which='both')

plt.suptitle('Requisition Amount Statistical Analysis', fontsize=16, fontweight='bold', y=0.995)
plt.tight_layout()
plt.savefig(output_dir / '07_requisition_distribution.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"   Saved: 07_requisition_distribution.png")

# ============================================================================
# CHART 8: TIME SERIES ANALYSIS (Monthly Trends)
# ============================================================================
print("\n8. Creating Time Series Analysis...")

df_exhibit['Month'] = pd.to_datetime(df_exhibit['Date Received-DPS']).dt.to_period('M')
monthly_data = df_exhibit.groupby('Month').agg({
    'Requisition Amount': ['sum', 'count', 'mean']
}).reset_index()
monthly_data.columns = ['Month', 'Total_Spend', 'Count', 'Avg_Amount']
monthly_data['Month'] = monthly_data['Month'].astype(str)

fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(16, 14))

# Monthly spend
x_pos = np.arange(len(monthly_data))
ax1.plot(x_pos, monthly_data['Total_Spend'] / 1000000, marker='o', linewidth=2, color='blue', markersize=4)
ax1.fill_between(x_pos, monthly_data['Total_Spend'] / 1000000, alpha=0.3, color='blue')
ax1.set_ylabel('Total Spend ($ Millions)', fontsize=11, fontweight='bold')
ax1.set_title('Monthly Spend Trend', fontsize=12, fontweight='bold')
ax1.grid(True, alpha=0.3)
ax1.set_xticks(x_pos[::3])
ax1.set_xticklabels(monthly_data['Month'][::3], rotation=45, ha='right')

# Monthly count
ax2.bar(x_pos, monthly_data['Count'], color='green', alpha=0.7, edgecolor='black')
ax2.set_ylabel('Number of Requisitions', fontsize=11, fontweight='bold')
ax2.set_title('Monthly Requisition Volume', fontsize=12, fontweight='bold')
ax2.grid(axis='y', alpha=0.3)
ax2.set_xticks(x_pos[::3])
ax2.set_xticklabels(monthly_data['Month'][::3], rotation=45, ha='right')

# Average amount
ax3.plot(x_pos, monthly_data['Avg_Amount'] / 1000, marker='s', linewidth=2, color='red', markersize=4)
ax3.fill_between(x_pos, monthly_data['Avg_Amount'] / 1000, alpha=0.3, color='red')
ax3.set_xlabel('Month', fontsize=11, fontweight='bold')
ax3.set_ylabel('Average Requisition ($K)', fontsize=11, fontweight='bold')
ax3.set_title('Average Requisition Amount Trend', fontsize=12, fontweight='bold')
ax3.grid(True, alpha=0.3)
ax3.set_xticks(x_pos[::3])
ax3.set_xticklabels(monthly_data['Month'][::3], rotation=45, ha='right')

plt.suptitle('Time Series Analysis: Monthly Trends (2022-2026)', fontsize=16, fontweight='bold', y=0.995)
plt.tight_layout()
plt.savefig(output_dir / '08_time_series_analysis.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"   Saved: 08_time_series_analysis.png")

# ============================================================================
# CHART 9: PROCUREMENT PHASE ANALYSIS
# ============================================================================
print("\n9. Creating Procurement Phase Analysis...")

phase_data = df_exhibit[df_exhibit['Procurement Phase'].notna()]
phase_spend = phase_data.groupby('Procurement Phase').agg({
    'Requisition Amount': ['sum', 'count', 'mean']
}).reset_index()
phase_spend.columns = ['Phase', 'Total_Spend', 'Count', 'Avg_Amount']
phase_spend = phase_spend.sort_values('Total_Spend', ascending=False)

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 8))

# Spend by phase
colors = sns.color_palette("Set2", len(phase_spend))
bars1 = ax1.barh(np.arange(len(phase_spend)), phase_spend['Total_Spend'] / 1000000,
                 color=colors, alpha=0.8, edgecolor='black')
ax1.set_yticks(np.arange(len(phase_spend)))
ax1.set_yticklabels(phase_spend['Phase'], fontsize=9)
ax1.set_xlabel('Total Spend ($ Millions)', fontsize=11, fontweight='bold')
ax1.set_title('Spend by Procurement Phase', fontsize=12, fontweight='bold')
ax1.invert_yaxis()

for i, (bar, value, count) in enumerate(zip(bars1, phase_spend['Total_Spend'], phase_spend['Count'])):
    width = bar.get_width()
    ax1.text(width, bar.get_y() + bar.get_height()/2.,
            f' ${value/1000000:.1f}M ({count:,} req)',
            ha='left', va='center', fontsize=9, fontweight='bold')

# Count by phase
bars2 = ax2.barh(np.arange(len(phase_spend)), phase_spend['Count'],
                 color=colors, alpha=0.8, edgecolor='black')
ax2.set_yticks(np.arange(len(phase_spend)))
ax2.set_yticklabels(phase_spend['Phase'], fontsize=9)
ax2.set_xlabel('Number of Requisitions', fontsize=11, fontweight='bold')
ax2.set_title('Requisition Count by Procurement Phase', fontsize=12, fontweight='bold')
ax2.invert_yaxis()

for i, (bar, count) in enumerate(zip(bars2, phase_spend['Count'])):
    width = bar.get_width()
    ax2.text(width, bar.get_y() + bar.get_height()/2.,
            f' {count:,}',
            ha='left', va='center', fontsize=9, fontweight='bold')

plt.suptitle('Procurement Phase Analysis', fontsize=16, fontweight='bold', y=0.98)
plt.tight_layout()
plt.savefig(output_dir / '09_procurement_phase_analysis.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"   Saved: 09_procurement_phase_analysis.png")

# ============================================================================
# CHART 10: CONTRACT UTILIZATION ANALYSIS
# ============================================================================
print("\n10. Creating Contract Utilization Analysis...")

contracts = [
    {'Name': 'SPC49756A-PO14954\n(2007-2013)', 'Limit': 211523.70, 'Matched': 195246.71, 'Vendor': 'Ice Mountain'},
    {'Name': 'SPC82197E-PO30043\n(2014-2022)', 'Limit': 925006.00, 'Matched': 689659.82, 'Vendor': 'Bluetriton'}
]

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8))

# Utilization comparison
names = [c['Name'] for c in contracts]
limits = [c['Limit'] for c in contracts]
matched = [c['Matched'] for c in contracts]
unused = [l - m for l, m in zip(limits, matched)]
utilization_pct = [(m/l*100) for l, m in zip(limits, matched)]

x = np.arange(len(contracts))
width = 0.35

bars1 = ax1.bar(x - width/2, [l/1000 for l in limits], width, label='Contract Limit',
                color='lightblue', edgecolor='black', linewidth=2)
bars2 = ax1.bar(x + width/2, [m/1000 for m in matched], width, label='Actual Spend',
                color='darkblue', edgecolor='black', linewidth=2)

ax1.set_ylabel('Amount ($K)', fontsize=11, fontweight='bold')
ax1.set_title('Contract Limit vs Actual Spend', fontsize=12, fontweight='bold')
ax1.set_xticks(x)
ax1.set_xticklabels(names, fontsize=10)
ax1.legend(fontsize=10)
ax1.grid(axis='y', alpha=0.3)

# Add value labels
for bar in bars1:
    height = bar.get_height()
    ax1.text(bar.get_x() + bar.get_width()/2., height,
            f'${height:.0f}K',
            ha='center', va='bottom', fontsize=10, fontweight='bold')

for bar in bars2:
    height = bar.get_height()
    ax1.text(bar.get_x() + bar.get_width()/2., height,
            f'${height:.0f}K',
            ha='center', va='bottom', fontsize=10, fontweight='bold')

# Utilization percentage
colors_util = ['green' if u >= 85 else 'orange' if u >= 70 else 'red' for u in utilization_pct]
bars3 = ax2.bar(x, utilization_pct, color=colors_util, alpha=0.7, edgecolor='black', linewidth=2)
ax2.set_ylabel('Utilization Rate (%)', fontsize=11, fontweight='bold')
ax2.set_title('Contract Utilization Rate', fontsize=12, fontweight='bold')
ax2.set_xticks(x)
ax2.set_xticklabels(names, fontsize=10)
ax2.set_ylim([0, 100])
ax2.axhline(y=85, color='green', linestyle='--', linewidth=1, alpha=0.5, label='Target: 85%')
ax2.axhline(y=70, color='orange', linestyle='--', linewidth=1, alpha=0.5, label='Warning: 70%')
ax2.legend(fontsize=9)
ax2.grid(axis='y', alpha=0.3)

for bar, util in zip(bars3, utilization_pct):
    height = bar.get_height()
    ax2.text(bar.get_x() + bar.get_width()/2., height,
            f'{util:.1f}%',
            ha='center', va='bottom', fontsize=11, fontweight='bold')

plt.suptitle('Bottled Water Contract Utilization Analysis', fontsize=16, fontweight='bold', y=0.98)
plt.tight_layout()
plt.savefig(output_dir / '10_contract_utilization.png', dpi=300, bbox_inches='tight')
plt.close()
print(f"   Saved: 10_contract_utilization.png")

# ============================================================================
# SUMMARY PARETO TABLE
# ============================================================================
print("\n" + "="*100)
print("PARETO ANALYSIS SUMMARY (80/20 RULE)")
print("="*100)

print("\n1. DEPARTMENTS:")
pareto_depts = dept_cumulative[dept_cumulative <= 80]
print(f"   - Top {len(pareto_depts)} departments account for 80% of spend")
print(f"   - These represent {len(pareto_depts)/len(dept_spend)*100:.1f}% of all departments")
print(f"   - Departments:", ', '.join(pareto_depts.index.tolist()))

print("\n2. JUSTIFICATIONS:")
just_pareto = just_cumulative[just_cumulative <= 80]
print(f"   - Top {len(just_pareto)} justification(s) account for 80% of spend")
print(f"   - Primary issue: '{just_pareto.index[0]}'")
print(f"   - This category alone: {just_spend_pct.iloc[0]:.1f}% of total spend")

print("\n3. VENDORS:")
vendor_cumul_all = (df_exhibit.groupby('Vendor Name')['Requisition Amount'].sum().sort_values(ascending=False) / df_exhibit['Requisition Amount'].sum() * 100).cumsum()
vendor_pareto = vendor_cumul_all[vendor_cumul_all <= 80]
print(f"   - Top {len(vendor_pareto)} vendors account for 80% of spend")
print(f"   - These represent {len(vendor_pareto)/df_exhibit['Vendor Name'].nunique()*100:.1f}% of all vendors ({len(vendor_pareto)}/{df_exhibit['Vendor Name'].nunique()})")

print("\n4. APPROVAL STATUS:")
total_approved_spend = approval_spend.loc['APPROVED', 'sum']
total_spend_decisions = approval_spend['sum'].sum()
print(f"   - Approved requisitions: {total_approved_spend/total_spend_decisions*100:.1f}% of spend")
print(f"   - Denied requisitions: {approval_spend.loc['DENIED', 'sum']/total_spend_decisions*100:.1f}% of spend")
print(f"   - Approval rate: {approval_spend.loc['APPROVED', 'count']/(approval_spend['count'].sum())*100:.1f}%")

print("\n" + "="*100)
print("VISUALIZATION GENERATION COMPLETE!")
print("="*100)
print(f"\nAll charts saved to: {output_dir}")
print("\nGenerated Charts:")
print("  1. Department Pareto Analysis")
print("  2. Top 10 Departments Pie Chart")
print("  3. Justification Category Pareto Analysis")
print("  4. Approval vs Denial Analysis")
print("  5. Top 20 Vendors Pareto Analysis")
print("  6. Contract Comparison (Product Mix)")
print("  7. Requisition Amount Distribution")
print("  8. Time Series Analysis (Monthly Trends)")
print("  9. Procurement Phase Analysis")
print(" 10. Contract Utilization Analysis")
print("="*100)
