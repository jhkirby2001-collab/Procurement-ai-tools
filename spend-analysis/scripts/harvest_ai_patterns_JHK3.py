"""
Harvest AI-classified patterns into rules for the production classifier.

Per locked option B (2026-04-29): AI runs ONCE during build to mine patterns
from the no-rule-fire long tail. This script promotes those AI classifications
into a separate rules file (`keyword_rules_from_ai_JHK3.csv`) that the
production classifier loads alongside the hand-curated rules. Procurement
staff can review/edit the AI-mined file independently of the curated one.

Promotion thresholds (tunable via CLI args):
- HIGH confidence: promote regardless of row_count (defensible AI calls)
- MEDIUM confidence: promote only if row_count >= --medium-min-rows
- LOW confidence: NEVER promote — stays in human-review pile

Each promoted rule is an exact-match keyword rule on the description string.
Output is auditable per row: every rule carries the AI's confidence, reason,
and the row_count it represents in source data.
"""

from __future__ import annotations

import argparse
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
AI_FILE = ROOT / "data" / "processed" / "ai_classified_unique_descriptions_JHK3.csv"
OUT = ROOT / "data" / "reference" / "keyword_rules_from_ai_JHK3.csv"

VALID_CATEGORIES = {
    "Vehicles & Fleet", "Heavy Equipment & Machinery", "Equipment Rental & Leasing",
    "Construction Materials", "Construction & Trades Services",
    "Facilities Operations & Maintenance", "Landscaping, Grounds & Irrigation",
    "Janitorial, Sanitation & Waste", "Chemicals & Water Treatment",
    "Public Safety, Uniforms & PPE", "Medical & Health Services",
    "Animal Care & Veterinary", "IT, Telecom & Audio/Visual",
    "Office, Print & Marketing", "Furniture & Furnishings",
    "Professional & Administrative Services", "Grants & Pass-Through Funding",
}


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--medium-min-rows", type=int, default=5,
                   help="Minimum row_count for medium-confidence rules (default 5)")
    p.add_argument("--include-no-class", action="store_true",
                   help="Include rules where AI couldn't assign an NIGP class (Business Category only)")
    args = p.parse_args()

    if not AI_FILE.exists():
        raise SystemExit(f"AI output not found at {AI_FILE}. Run ai_classify_JHK3.py first.")

    df = pd.read_csv(AI_FILE).fillna("")
    n_total = len(df)
    print(f"AI output rows total: {n_total:,}")

    # Validation: drop any row whose category isn't in our 17 (defense vs schema drift)
    invalid_cat = ~df["business_category"].isin(VALID_CATEGORIES)
    if invalid_cat.any():
        print(f"  Dropping {invalid_cat.sum():,} rows with invalid business_category (not in 17)")
        df = df[~invalid_cat]

    # Tier the rows by confidence
    high = df[df["confidence"] == "high"]
    medium = df[df["confidence"] == "medium"]
    low = df[df["confidence"] == "low"]
    print(f"  High confidence: {len(high):,}")
    print(f"  Medium confidence: {len(medium):,}")
    print(f"  Low confidence (never promoted): {len(low):,}")

    # Apply medium row-count gate
    medium_kept = medium[medium["row_count"] >= args.medium_min_rows]
    medium_dropped = len(medium) - len(medium_kept)
    if medium_dropped > 0:
        print(f"  Medium rows dropped (row_count < {args.medium_min_rows}): {medium_dropped:,}")

    promoted = pd.concat([high, medium_kept], ignore_index=True)

    # Optional: drop rows without an NIGP class assignment
    if not args.include_no_class:
        no_class = promoted["nigp_class_3digit"] == ""
        # Keep grants regardless (no NIGP class by design)
        keep_grants = promoted["business_category"] == "Grants & Pass-Through Funding"
        drop = no_class & ~keep_grants
        if drop.any():
            print(f"  Dropping {drop.sum():,} rows with no NIGP class (use --include-no-class to keep)")
            promoted = promoted[~drop]

    # Build the rule rows
    nigp_class = promoted["nigp_class_3digit"].astype(str)
    grants_mask = promoted["business_category"] == "Grants & Pass-Through Funding"
    has_class = nigp_class != ""
    match_level = pd.Series("review", index=promoted.index, dtype=object)
    match_level[has_class] = "broad"
    match_level[grants_mask] = "broad"  # grants are confidently classified by intent

    rules = pd.DataFrame({
        "pattern": promoted["description"].astype(str),
        "match_type": "exact",
        "business_category": promoted["business_category"],
        "nigp_class_3digit": nigp_class,
        "nigp_item_5digit": "",
        "nigp_match_level": match_level.values,
        "notes": ("AI-mined (Haiku 4.5, conf=" + promoted["confidence"]
                  + ", row_count=" + promoted["row_count"].astype(str)
                  + "): " + promoted["reason"].astype(str).str.slice(0, 120)),
    })

    rules.to_csv(OUT, index=False)
    print()
    print(f"Promoted {len(rules):,} AI-mined rules -> {OUT}")
    print(f"  Rules with NIGP class:    {(rules['nigp_class_3digit'] != '').sum():,}")
    print(f"  Rules without NIGP class: {(rules['nigp_class_3digit'] == '').sum():,}")
    print()
    print("=== Promoted rules by Business Category ===")
    print(rules.groupby("business_category").size().sort_values(ascending=False).to_string())


if __name__ == "__main__":
    main()
