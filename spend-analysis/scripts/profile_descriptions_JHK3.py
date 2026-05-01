"""
Profile description text columns in the EY raw file.

Purpose: tell us where keyword rules can carry the load (high-repetition,
informative descriptions) and where AI assist will be needed (low-repetition,
sparse, or uninformative descriptions).

Inputs:  spend-analysis/data/processed/ey_raw_cache.parquet
Outputs:
  spend-analysis/data/processed/description_column_profile_JHK3.csv
  spend-analysis/data/processed/top_description_patterns_JHK3.csv
  spend-analysis/data/processed/lowinfo_description_examples_JHK3.csv
"""

from pathlib import Path
import re
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "data" / "processed" / "ey_raw_cache.parquet"
OUT_DIR = ROOT / "data" / "processed"

# Description columns we care about. EXCLUDES EY-supplied taxonomy fields
# (NIGP Description, Commodity Description) per locked decision: classifier
# uses transaction descriptions only, not EY's classification work.
DESC_COLS = [
    "PO Description",
    "PO Item Description",
    "AP Invoice Line Description",
    "Invoice Distribution Description",
]

LOWINFO_PATTERNS = [
    r"^\s*$",
    r"^(misc|miscellaneous|various|other|n/?a|none|unknown|tbd)\s*$",
    r"^per\s+(contract|agreement|spec|specs|specification|po|order|invoice)\s*$",
    r"^as\s+(needed|specified|per\s+spec)\s*$",
    r"^(see|refer\s+to)\s+(attached|attachment|invoice|po|contract).*$",
    r"^(item|line|charge)\s*[#:\-]?\s*\d+\s*$",
    r"^[\W_\d]+$",  # only punctuation/digits
]
LOWINFO_RE = re.compile("|".join(LOWINFO_PATTERNS), flags=re.IGNORECASE)


def is_lowinfo(s) -> bool:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return True
    s = str(s).strip()
    if not s:
        return True
    if len(s) <= 2:
        return True
    return bool(LOWINFO_RE.match(s))


def main() -> None:
    print(f"Loading {SRC} ...")
    df = pd.read_parquet(SRC, columns=DESC_COLS)
    n = len(df)
    print(f"Rows: {n:,}\n")

    # Per-column profile
    rows = []
    for col in DESC_COLS:
        s = df[col]
        nonnull = s.notna()
        nonblank = nonnull & (s.astype(str).str.strip() != "")
        text = s[nonblank].astype(str).str.strip()
        lengths = text.str.len()
        lowinfo_mask = s.apply(is_lowinfo)
        rows.append({
            "column": col,
            "fill_rate_pct": round(nonblank.sum() / n * 100, 1),
            "n_distinct": text.nunique(),
            "median_len": int(lengths.median()) if not lengths.empty else 0,
            "p90_len": int(lengths.quantile(0.90)) if not lengths.empty else 0,
            "p99_len": int(lengths.quantile(0.99)) if not lengths.empty else 0,
            "max_len": int(lengths.max()) if not lengths.empty else 0,
            "n_lowinfo_or_missing": int(lowinfo_mask.sum()),
            "lowinfo_or_missing_pct": round(lowinfo_mask.sum() / n * 100, 1),
        })
    profile = pd.DataFrame(rows)
    out_profile = OUT_DIR / "description_column_profile_JHK3.csv"
    profile.to_csv(out_profile, index=False)

    print("=== Description column profile ===")
    print(profile.to_string(index=False))
    print(f"\nWritten: {out_profile}")

    # Build a Description_Best column = first non-blank/non-lowinfo across the
    # four columns, in priority order. Then profile its top patterns.
    print("\nBuilding Description_Best (first usable across columns) ...")
    best = pd.Series([""] * n, index=df.index, dtype=object)
    for col in DESC_COLS:
        empty = best == ""
        candidate = df[col].where(~df[col].apply(is_lowinfo), "")
        best = best.where(~empty, candidate.fillna(""))
    best = best.astype(str).str.strip()

    n_best_filled = (best != "").sum()
    n_best_lowinfo = n - n_best_filled
    print(f"Description_Best populated:           {n_best_filled:,} ({n_best_filled/n*100:.1f}%)")
    print(f"No usable description across all 4:   {n_best_lowinfo:,} ({n_best_lowinfo/n*100:.1f}%)")

    # Top recurring patterns in Description_Best (rule candidates)
    counts = best[best != ""].value_counts()
    cum_pct = (counts.cumsum() / n_best_filled * 100).round(2)
    top = pd.DataFrame({
        "description": counts.index,
        "row_count": counts.values,
        "pct_of_filled": (counts.values / n_best_filled * 100).round(3),
        "cumulative_pct_of_filled": cum_pct.values,
    })

    out_top = OUT_DIR / "top_description_patterns_JHK3.csv"
    top.head(2000).to_csv(out_top, index=False)
    print(f"\nTop 2,000 unique Description_Best strings written: {out_top}")

    # Concentration stats
    for k in [10, 50, 100, 500, 1000, 5000]:
        if len(top) >= k:
            covered = top["row_count"].head(k).sum()
            print(f"  Top {k:>5,} unique descriptions cover {covered:>9,} rows ({covered/n_best_filled*100:5.1f}% of filled)")
    print(f"  Total unique non-lowinfo descriptions: {len(top):,}")

    # Sample of low-info / unusable rows
    lowinfo_examples = df[best == ""][DESC_COLS].head(50)
    out_low = OUT_DIR / "lowinfo_description_examples_JHK3.csv"
    lowinfo_examples.to_csv(out_low, index=False)
    print(f"\n50 sample rows with no usable description: {out_low}")

    # Print top 30 patterns inline
    print("\n=== Top 30 most-repeated descriptions (rule candidates) ===")
    pd.set_option("display.max_colwidth", 90)
    pd.set_option("display.width", 200)
    print(top.head(30).to_string(index=False))


if __name__ == "__main__":
    main()
