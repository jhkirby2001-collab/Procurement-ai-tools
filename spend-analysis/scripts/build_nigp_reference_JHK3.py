"""
Build NIGP reference catalog from the EY raw data file.

Inputs:  spend-analysis/data/processed/ey_raw_cache.parquet
Outputs: spend-analysis/data/reference/nigp_codes_10digit_JHK3.csv
         spend-analysis/data/reference/nigp_codes_5digit_JHK3.csv

Canonical description rule: mode (most-frequent description per code),
with longest string as tiebreaker. Approved 2026-04-29 by JHK3.
"""

from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "data" / "processed" / "ey_raw_cache.parquet"
OUT_DIR = ROOT / "data" / "reference"
OUT_DIR.mkdir(parents=True, exist_ok=True)

SOURCE_LABEL = "EY raw data file (City of Chicago purchasing extract, 2002-2025)"


def canonical_description(series: pd.Series) -> str:
    """Most-frequent non-null description; ties broken by longest string."""
    s = series.dropna().astype(str).str.strip()
    s = s[s != ""]
    if s.empty:
        return ""
    counts = s.value_counts()
    top_freq = counts.iloc[0]
    candidates = counts[counts == top_freq].index.tolist()
    return max(candidates, key=len)


def main() -> None:
    df = pd.read_parquet(SRC, columns=["NIGP Code", "NIGP Description"])
    coded = df[df["NIGP Code"].notna()].copy()

    coded["nigp_code_10digit"] = (
        coded["NIGP Code"].astype("Int64").astype(str).str.zfill(10)
    )
    coded["nigp_class_3digit"] = coded["nigp_code_10digit"].str[:3]
    coded["nigp_item_5digit"] = coded["nigp_code_10digit"].str[:5]

    # ---- 10-digit reference ----
    g10 = coded.groupby("nigp_code_10digit", sort=True)
    ref10 = pd.DataFrame({
        "nigp_class_3digit": g10["nigp_class_3digit"].first(),
        "nigp_item_5digit": g10["nigp_item_5digit"].first(),
        "canonical_description": g10["NIGP Description"].apply(canonical_description),
        "description_variant_count": g10["NIGP Description"]
            .apply(lambda s: s.dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique()),
        "row_count": g10.size(),
    }).reset_index()
    ref10["source"] = SOURCE_LABEL
    ref10 = ref10[[
        "nigp_code_10digit", "nigp_class_3digit", "nigp_item_5digit",
        "canonical_description", "description_variant_count", "row_count", "source",
    ]]
    out10 = OUT_DIR / "nigp_codes_10digit_JHK3.csv"
    ref10.to_csv(out10, index=False)

    # ---- 5-digit (Class-Item) rollup ----
    g5 = coded.groupby("nigp_item_5digit", sort=True)
    ref5 = pd.DataFrame({
        "nigp_class_3digit": g5["nigp_class_3digit"].first(),
        "canonical_description": g5["NIGP Description"].apply(canonical_description),
        "description_variant_count": g5["NIGP Description"]
            .apply(lambda s: s.dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique()),
        "distinct_10digit_codes": g5["nigp_code_10digit"].nunique(),
        "row_count": g5.size(),
    }).reset_index()
    ref5["source"] = SOURCE_LABEL
    ref5 = ref5[[
        "nigp_item_5digit", "nigp_class_3digit", "canonical_description",
        "description_variant_count", "distinct_10digit_codes", "row_count", "source",
    ]]
    out5 = OUT_DIR / "nigp_codes_5digit_JHK3.csv"
    ref5.to_csv(out5, index=False)

    print(f"10-digit reference: {len(ref10):,} rows -> {out10}")
    print(f" 5-digit reference: {len(ref5):,} rows -> {out5}")
    print(f"Distinct 3-digit Classes: {coded['nigp_class_3digit'].nunique()}")
    print(f"Codes with multiple description variants (10-digit): "
          f"{(ref10['description_variant_count'] > 1).sum()}")
    print(f"Codes with multiple description variants (5-digit):  "
          f"{(ref5['description_variant_count'] > 1).sum()}")
    print("Top 10 noisiest 5-digit codes by variant count:")
    print(ref5.nlargest(10, "description_variant_count")[
        ["nigp_item_5digit", "canonical_description", "description_variant_count", "row_count"]
    ].to_string(index=False))


if __name__ == "__main__":
    main()
