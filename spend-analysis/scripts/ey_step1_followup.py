"""
Quick follow-up checks on Step 1 findings:
1. PO Contract Number — encoding of 'no contract' (not nulls → check placeholders)
2. Favorite Healthcare Staffing — is the $48B concentration real or an artifact?
3. Date distribution — is the bulk of spend in the 2021-2025 window or spread across 2002-2025?
4. AP Invoice Amount zero-count vs Current PO Dist Amount zero-count (may explain 72% variance)
"""

import pandas as pd
import numpy as np

IN = "/workspaces/Procurement-ai-tools/spend-analysis/data/raw/ey raw data.xlsx"
OUT = "/workspaces/Procurement-ai-tools/spend-analysis/outputs/ey_benchmark_step1_followup_JHK3.txt"

print("[load] reading xlsx …")
df = pd.read_excel(IN, sheet_name="Report 1", engine="openpyxl")
df.columns = [c.strip() for c in df.columns]
print(f"[load] {len(df):,} rows")

log = []

def add(s=""):
    print(s)
    log.append(s)

# ── 1. PO Contract Number encoding ──────────────────────────────
add("="*78); add("1. PO CONTRACT NUMBER ENCODING"); add("="*78)
col = "PO Contract Number"
s = df[col]
add(f"dtype                    : {s.dtype}")
add(f"null count               : {s.isna().sum():,}")
# Show the 15 most common values to understand encoding
vc = s.fillna("(NULL)").astype(str).str.strip().value_counts().head(15)
add(f"\nTop 15 values by frequency:")
for k, v in vc.items():
    add(f"  {str(k)[:40]:<40s} : {v:>8,}")
# Is there a "0"/"none"/"n/a" convention?
blanks_like = s.astype(str).str.strip().str.lower().isin(
    ["", "nan", "none", "n/a", "na", "null", "0"])
add(f"\nValues matching blank-like patterns (''/nan/none/n-a/na/null/0) : "
    f"{int(blanks_like.sum()):,} rows")
# Spend under blank-like
spend = pd.to_numeric(df["Current PO Dist Amount"], errors="coerce")
add(f"Spend under blank-like contract values (PO Dist)              : "
    f"${float(spend[blanks_like].sum()):,.2f}")

# ── 2. Favorite Healthcare Staffing drill ───────────────────────
add(""); add("="*78); add("2. FAVORITE HEALTHCARE STAFFING — ANOMALY CHECK"); add("="*78)
fhs = df[df["Vendor Name"].astype(str).str.contains("Favorite Healthcare", case=False, na=False)]
add(f"Rows             : {len(fhs):,}")
if "Current PO Dist Amount" in df.columns:
    s_cur = pd.to_numeric(fhs["Current PO Dist Amount"], errors="coerce")
    s_ap  = pd.to_numeric(fhs["AP Invoice Amount"], errors="coerce")
    s_pc  = pd.to_numeric(fhs["Pay Check Amount"], errors="coerce")
    add(f"Current PO Dist  : sum=${float(s_cur.sum()):>20,.2f}  "
        f"max=${float(s_cur.max()):,.2f}  "
        f"median=${float(s_cur.median()):,.2f}")
    add(f"AP Invoice       : sum=${float(s_ap.sum()):>20,.2f}  "
        f"max=${float(s_ap.max()):,.2f}  "
        f"median=${float(s_ap.median()):,.2f}")
    add(f"Pay Check        : sum=${float(s_pc.sum()):>20,.2f}  "
        f"max=${float(s_pc.max()):,.2f}  "
        f"median=${float(s_pc.median()):,.2f}")
    # Top rows
    top10 = fhs.nlargest(5, "Current PO Dist Amount")[
        ["PO Creation Date","Purchase Order","PO Contract Number",
         "Current PO Dist Amount","AP Invoice Amount","Pay Check Amount"]
        ]
    add("\nTop 5 largest rows by Current PO Dist Amount:")
    for _, r in top10.iterrows():
        add(f"  {str(r.get('PO Creation Date',''))[:10]}  "
            f"PO={str(r.get('Purchase Order',''))[:12]:<12s}  "
            f"Cur=${float(r.get('Current PO Dist Amount',0) or 0):>18,.2f}  "
            f"AP=${float(r.get('AP Invoice Amount',0) or 0):>18,.2f}  "
            f"PC=${float(r.get('Pay Check Amount',0) or 0):>18,.2f}")

# ── 3. Date distribution ────────────────────────────────────────
add(""); add("="*78); add("3. SPEND BY YEAR (PO Creation Date)"); add("="*78)
dates = pd.to_datetime(df["PO Creation Date"], errors="coerce")
spend = pd.to_numeric(df["Current PO Dist Amount"], errors="coerce")
df_y = pd.DataFrame({"year": dates.dt.year, "spend": spend, "ap": pd.to_numeric(df["AP Invoice Amount"], errors="coerce")})
by_year = df_y.groupby("year", dropna=True).agg(
    rows=("spend","size"),
    current_po_dist=("spend","sum"),
    ap_invoice=("ap","sum")
).reset_index().sort_values("year")
for _, r in by_year.iterrows():
    add(f"  {int(r['year'])}  rows={int(r['rows']):>8,}  "
        f"PO Dist=${float(r['current_po_dist']):>18,.2f}  "
        f"AP=${float(r['ap_invoice']):>18,.2f}")

# ── 4. Zero-amount counts ───────────────────────────────────────
add(""); add("="*78); add("4. ZERO-VALUE TRANSACTION COUNTS"); add("="*78)
for c in ("Current PO Dist Amount","AP Invoice Amount","Pay Check Amount"):
    s = pd.to_numeric(df[c], errors="coerce").fillna(0)
    add(f"{c:<30s} rows==0: {int((s==0).sum()):>8,}  "
        f"rows<0: {int((s<0).sum()):>6,}  "
        f"rows>0: {int((s>0).sum()):>8,}")

# ── 5. Is there line-item duplication driving AP < PO? ──────────
add(""); add("="*78); add("5. LINE-ITEM / DUPLICATION CHECK"); add("="*78)
if "Purchase Order" in df.columns:
    per_po = df.groupby("Purchase Order").size()
    add(f"Unique Purchase Orders          : {per_po.index.nunique():,}")
    add(f"Avg rows per Purchase Order     : {per_po.mean():.2f}")
    add(f"Max rows per Purchase Order     : {int(per_po.max()):,}")
    add(f"POs with >1 row                 : {int((per_po>1).sum()):,}")

with open(OUT, "w") as f:
    f.write("\n".join(log))
print(f"\n[done] {OUT}")
