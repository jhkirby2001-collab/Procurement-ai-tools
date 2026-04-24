"""
Step 1 — Data Validation Profile
================================
Portable validation pass for a procurement transaction extract.
Produces a leadership-ready data-quality profile before any savings modeling.

Per CLAUDE.md v5.1 / user instructions — this script must run on any
organization's procurement extract by changing ONLY the CONFIG block below.
"""

import os
import sys
import json
import warnings
from pathlib import Path

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# ── CONFIG ──────────────────────────────────────────────────────────
JURISDICTION_NAME           = "City of Chicago DPS"
ANALYSIS_PERIOD             = "2021-2026 (EY benchmark extract)"
INPUT_FILE                  = "/workspaces/Procurement-ai-tools/spend-analysis/data/raw/ey raw data.xlsx"
SHEET_NAME                  = "Report 1"
OUTPUT_DIR                  = "/workspaces/Procurement-ai-tools/spend-analysis/outputs"
CACHE_FILE                  = "/workspaces/Procurement-ai-tools/spend-analysis/data/processed/ey_raw_cache.parquet"
VALIDATION_REPORT           = "ey_benchmark_step1_validation_JHK3.txt"

ASSUMPTION_ADDRESSABILITY   = 0.85
ASSUMPTION_SAVINGS_RATE     = 0.10
ASSUMPTION_HORIZON_YEARS    = 3
# ────────────────────────────────────────────────────────────────────


def money(x):
    try:
        if x is None or (isinstance(x, float) and np.isnan(x)):
            return "$0.00"
        return f"${x:,.2f}"
    except Exception:
        return str(x)


def pct(part, whole):
    if whole in (0, None) or (isinstance(whole, float) and np.isnan(whole)):
        return "n/a"
    return f"{(part / whole) * 100:.2f}%"


def load_data():
    Path(os.path.dirname(CACHE_FILE)).mkdir(parents=True, exist_ok=True)
    if os.path.exists(CACHE_FILE):
        mtime_raw = os.path.getmtime(INPUT_FILE)
        mtime_cache = os.path.getmtime(CACHE_FILE)
        if mtime_cache > mtime_raw:
            print(f"[cache] Loading {CACHE_FILE}")
            return pd.read_parquet(CACHE_FILE)

    print(f"[read] Loading {INPUT_FILE} sheet='{SHEET_NAME}' (may take several minutes)")
    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME, engine="openpyxl")
    print(f"[read] Loaded {len(df):,} rows × {len(df.columns)} columns")

    # Cache to parquet for fast reuse by Step 2
    try:
        df_cache = df.copy()
        # Convert any object columns that are problematic for parquet to string
        for c in df_cache.columns:
            if df_cache[c].dtype == object:
                df_cache[c] = df_cache[c].astype(str).where(df_cache[c].notna(), None)
        df_cache.to_parquet(CACHE_FILE, index=False)
        print(f"[cache] Saved {CACHE_FILE}")
    except Exception as e:
        print(f"[cache] Skipped parquet cache: {e}")
    return df


def section(title, out):
    line = "=" * 78
    out.append("")
    out.append(line)
    out.append(title)
    out.append(line)


def sub(title, out):
    out.append("")
    out.append(f"— {title} " + "-" * max(4, 72 - len(title)))


def main():
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
    out = []

    section(f"{JURISDICTION_NAME} — STEP 1 DATA VALIDATION PROFILE", out)
    out.append(f"Source file         : {INPUT_FILE}")
    out.append(f"Sheet               : {SHEET_NAME}")
    out.append(f"Analysis period     : {ANALYSIS_PERIOD}")
    out.append(f"Defense assumptions : Addressability={ASSUMPTION_ADDRESSABILITY:.0%} | "
               f"Savings rate={ASSUMPTION_SAVINGS_RATE:.0%} | "
               f"Horizon={ASSUMPTION_HORIZON_YEARS} yrs")

    df = load_data()

    # Normalize column names: strip whitespace
    df.columns = [str(c).strip() for c in df.columns]

    flags = []  # items to surface at top under ⚠ FLAG FOR LEADERSHIP

    # ── 1. Shape and date range ─────────────────────────────────────
    section("1. SHAPE AND DATE RANGE", out)
    out.append(f"Row count    : {len(df):,}")
    out.append(f"Column count : {len(df.columns)}")

    po_date_col = "PO Creation Date"
    if po_date_col in df.columns:
        dates = pd.to_datetime(df[po_date_col], errors="coerce")
        out.append(f"{po_date_col} min : {dates.min()}")
        out.append(f"{po_date_col} max : {dates.max()}")
        out.append(f"{po_date_col} null count : {dates.isna().sum():,} "
                   f"({pct(dates.isna().sum(), len(df))})")
    else:
        flags.append(f"Missing expected column: {po_date_col}")

    # ── 2. Total spend — three ways ─────────────────────────────────
    section("2. TOTAL SPEND — THREE RECONCILIATION PATHS", out)
    spend_candidates = [
        ("Current PO Dist Amount", "Sum of Current PO Dist Amount"),
        ("AP Invoice Amount",      "Sum of AP Invoice Amount"),
        ("Pay Check Amount",       "Sum of Pay Check Amount"),
    ]
    spend_totals = {}
    for col, label in spend_candidates:
        if col in df.columns:
            s = pd.to_numeric(df[col], errors="coerce")
            total = float(s.sum(skipna=True))
            spend_totals[col] = total
            out.append(f"{label:<42s} : {money(total):>22s}   "
                       f"(n={s.notna().sum():>8,}, null={s.isna().sum():>7,})")
        else:
            out.append(f"{label:<42s} : COLUMN NOT FOUND")
            flags.append(f"Missing expected spend column: {col}")

    if len(spend_totals) >= 2:
        vals = list(spend_totals.values())
        max_v, min_v = max(vals), min(vals)
        variance = max_v - min_v
        variance_pct = variance / max_v if max_v else 0
        out.append("")
        out.append(f"Max - Min variance : {money(variance)} ({variance_pct:.2%} of max)")
        if variance_pct > 0.05:
            flags.append(f"Material variance ({variance_pct:.1%}) across three "
                         f"total-spend paths — Current PO Dist vs AP Invoice "
                         f"vs Pay Check Amount.")

    # ── 3. Unique counts ────────────────────────────────────────────
    section("3. UNIQUE ENTITY COUNTS", out)
    uniques = {
        "Vendor Number"      : "Unique vendors (by Vendor Number)",
        "Vendor Name"        : "Unique Vendor Names (raw, pre-normalization)",
        "NIGP Code"          : "Unique NIGP codes",
        "NIGP Description"   : "Unique NIGP Descriptions",
        "PO Lead Dept"       : "Unique PO Lead Departments",
        "Purchase Order"     : "Unique Purchase Order numbers",
        "PO Number"          : "Unique PO Numbers",
        "PO Contract Number" : "Unique PO Contract Numbers",
    }
    for col, label in uniques.items():
        if col in df.columns:
            n_unique = df[col].nunique(dropna=True)
            out.append(f"{label:<55s} : {n_unique:>10,}")
    # Vendor integrity check: names per vendor-number
    if "Vendor Number" in df.columns and "Vendor Name" in df.columns:
        per_num = df.groupby("Vendor Number")["Vendor Name"].nunique()
        multi_name = (per_num > 1).sum()
        if multi_name:
            flags.append(f"{multi_name:,} Vendor Numbers have >1 Vendor Name variant — "
                         f"name normalization likely required before vendor analytics.")

    # ── 4. Null / blank analysis ────────────────────────────────────
    section("4. NULL / BLANK ANALYSIS ON CRITICAL FIELDS", out)
    critical_cols = [
        "NIGP Code", "NIGP Description",
        "PO Contract Number", "Vendor Number",
        "PO Lead Dept", "Current PO Dist Amount",
    ]
    # Build a stable spend series for attribution where needed
    spend_col = None
    for c in ("Current PO Dist Amount", "AP Invoice Amount", "Pay Check Amount"):
        if c in df.columns:
            spend_col = c
            break
    total_spend = float(pd.to_numeric(df[spend_col], errors="coerce").sum()) if spend_col else 0.0

    for col in critical_cols:
        if col not in df.columns:
            out.append(f"{col:<30s} : COLUMN NOT FOUND")
            continue
        s = df[col]
        # Treat NaN, empty string, 'nan', 'None' as blank
        is_blank = s.isna() | (s.astype(str).str.strip().isin(["", "nan", "None", "NaN"]))
        n_blank = int(is_blank.sum())
        row_pct = pct(n_blank, len(df))
        # Spend attributed to blank rows
        if spend_col:
            blank_spend = float(pd.to_numeric(df.loc[is_blank, spend_col], errors="coerce").sum())
            spend_str = f"  spend={money(blank_spend):>18s} ({pct(blank_spend, total_spend)})"
        else:
            spend_str = ""
        out.append(f"{col:<30s} : blank rows={n_blank:>8,} ({row_pct}){spend_str}")

        # Flags
        if col in ("NIGP Code", "NIGP Description") and spend_col and total_spend:
            if blank_spend / total_spend > 0.05:
                flags.append(f"{pct(blank_spend, total_spend)} of spend has no {col} — "
                             f"category analytics will under-report by that amount.")
        if col == "PO Contract Number" and spend_col and total_spend:
            # This is the maverick-spend indicator; not a flag per se, but a headline
            out.append(f"  → MAVERICK INDICATOR: {pct(blank_spend, total_spend)} of spend "
                       f"has no PO Contract Number (off-contract candidate pool).")

    # ── 5. DPS vs Non-DPS ───────────────────────────────────────────
    section("5. DPS vs NON-DPS OVERSIGHT", out)
    dps_col = None
    for candidate in ("DPS / Non-DPS Flag", "DPS / Non-DPS", "DPS/Non-DPS Flag"):
        if candidate in df.columns:
            dps_col = candidate
            break
    if dps_col and spend_col:
        g = df.groupby(df[dps_col].fillna("(blank)"), dropna=False).agg(
            txn_count=(spend_col, "size"),
            total_spend=(spend_col, lambda s: pd.to_numeric(s, errors="coerce").sum()),
        ).reset_index()
        g.columns = [dps_col, "Transactions", "Total Spend"]
        g = g.sort_values("Total Spend", ascending=False)
        for _, r in g.iterrows():
            out.append(f"{str(r[dps_col]):<25s} txns={int(r['Transactions']):>8,}   "
                       f"spend={money(r['Total Spend']):>20s} "
                       f"({pct(r['Total Spend'], total_spend)})")
    else:
        out.append(f"DPS flag column not found (checked several variants).")
        flags.append("Cannot compute DPS vs Non-DPS breakdown — flag column missing.")

    # ── 6. Approval status ──────────────────────────────────────────
    section("6. APPROVAL STATUS BREAKDOWN", out)
    status_col = None
    for c in ("PO Approval Status", "Approval Status"):
        if c in df.columns:
            status_col = c
            break
    if status_col and spend_col:
        g = df.groupby(df[status_col].fillna("(blank)"), dropna=False).agg(
            txn_count=(spend_col, "size"),
            total_spend=(spend_col, lambda s: pd.to_numeric(s, errors="coerce").sum()),
        ).reset_index()
        g.columns = [status_col, "Transactions", "Total Spend"]
        g = g.sort_values("Total Spend", ascending=False)
        cancelled_spend = 0.0
        for _, r in g.iterrows():
            out.append(f"{str(r[status_col]):<30s} txns={int(r['Transactions']):>8,}   "
                       f"spend={money(r['Total Spend']):>20s} "
                       f"({pct(r['Total Spend'], total_spend)})")
            s_lower = str(r[status_col]).lower()
            if any(k in s_lower for k in ("cancel", "reject", "void")):
                cancelled_spend += float(r["Total Spend"])
        if cancelled_spend > 0:
            flags.append(f"{money(cancelled_spend)} in spend is on Cancelled/Rejected/Voided POs — "
                         f"exclude from run-rate, report as cost avoidance indicator.")
    else:
        out.append("Approval status column not found.")

    # AP payment status pass-through
    if "AP Invoice Payment Status" in df.columns:
        sub("AP Invoice Payment Status", out)
        pay = df["AP Invoice Payment Status"].fillna("(blank)")
        vc = pay.value_counts(dropna=False).head(15)
        for k, v in vc.items():
            out.append(f"{str(k):<35s} : {int(v):>10,}")

    # ── 7. Top 20 leaderboards ─────────────────────────────────────
    section("7. TOP 20 LEADERBOARDS (raw — no analysis yet)", out)
    if spend_col and "Vendor Name" in df.columns:
        sub("Top 20 Vendors by Spend", out)
        s = pd.to_numeric(df[spend_col], errors="coerce")
        top_v = (df.assign(_spend=s)
                   .groupby(["Vendor Number", "Vendor Name"], dropna=False)["_spend"]
                   .sum().sort_values(ascending=False).head(20))
        for (vnum, vname), amt in top_v.items():
            out.append(f"{str(vnum):<15s} {str(vname)[:55]:<55s} {money(amt):>20s}")

    if spend_col and "NIGP Code" in df.columns and "NIGP Description" in df.columns:
        sub("Top 20 NIGP Categories by Spend", out)
        s = pd.to_numeric(df[spend_col], errors="coerce")
        top_n = (df.assign(_spend=s)
                   .groupby(["NIGP Code", "NIGP Description"], dropna=False)["_spend"]
                   .sum().sort_values(ascending=False).head(20))
        for (ncode, ndesc), amt in top_n.items():
            out.append(f"{str(ncode):<12s} {str(ndesc)[:58]:<58s} {money(amt):>20s}")

    # ── Leadership flag summary (top of report) ─────────────────────
    flag_block = ["", "=" * 78, "⚠ FLAG FOR LEADERSHIP", "=" * 78]
    if flags:
        for f in flags:
            flag_block.append(f"• {f}")
    else:
        flag_block.append("No material data-quality flags surfaced during validation.")
    flag_block.append("")

    # Write report (flags at top, details after)
    full = flag_block + out
    report_path = os.path.join(OUTPUT_DIR, VALIDATION_REPORT)
    with open(report_path, "w") as f:
        f.write("\n".join(full))

    # Also print to stdout
    print("\n".join(full))
    print(f"\n[done] Validation report written to: {report_path}")


if __name__ == "__main__":
    main()
