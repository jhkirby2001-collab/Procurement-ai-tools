#!/usr/bin/env python3
"""
Read the Drinking Water Services Detail tab more carefully
"""

import pandas as pd
from pathlib import Path

file_path = Path("/workspaces/Spend_Analysis_and_Categorization/spend-analysis/outputs/EXHIBIT_B_CPO_PRESENTATION.xlsx")

print("=" * 80)
print("DRINKING WATER SERVICES DETAIL TAB")
print("=" * 80)
print()

try:
    # Read without headers to see raw structure
    df = pd.read_excel(file_path, sheet_name='Drinking Water Services Detail', header=None)

    print(f"Shape: {df.shape[0]} rows x {df.shape[1]} columns")
    print()

    # Print all rows to see the structure
    print("Full content:")
    print("-" * 80)
    for idx, row in df.iterrows():
        # Print non-null values in the row
        row_values = [str(val) for val in row if pd.notna(val)]
        if row_values:
            print(f"Row {idx}: {' | '.join(row_values)}")

except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
