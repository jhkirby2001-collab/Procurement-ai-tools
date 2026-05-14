"""
Post-classifier resolver: fills Business_Category (and NIGP class where AI assigned one)
on every Review_Flag=Yes row by merging in the existing AI-mined classifications from
ai_classified_unique_descriptions_JHK3.csv. Applies the canonical-mapping alignment used
by fix_ai_rule_category_mismatches_JHK3.py so AI category mismatches don't propagate.

Idempotent: only touches rows still flagged for review. Re-running after success is a no-op.

Run from repo root:
  python spend-analysis/scripts/resolve_review_queue_JHK3.py
"""
from __future__ import annotations

import csv
import re
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUTPUTS = ROOT.parent / "outputs"
MAPPING_PATH = OUTPUTS / "NIGP_Mapping_JHK3.csv"
REVIEW_QUEUE_PATH = OUTPUTS / "NIGP_Mapping_Review_Queue_JHK3.csv"
AI_PATH = ROOT / "data" / "processed" / "ai_classified_unique_descriptions_JHK3.csv"
CANON_PATH = ROOT / "data" / "reference" / "business_categories_JHK3.csv"
UNCOVERED_PATH = ROOT / "data" / "processed" / "review_queue_uncovered_uniques_JHK3.csv"

WHITESPACE_RE = re.compile(r"\s+")


def normalize(s: str) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return WHITESPACE_RE.sub(" ", str(s).strip()).upper()


def load_canonical() -> dict[str, str]:
    canon: dict[str, str] = {}
    with open(CANON_PATH, newline="") as f:
        for row in csv.DictReader(f):
            cls = str(row.get("nigp_class_3digit", "")).strip()
            cat = (row.get("business_category", "") or "").strip()
            if cls and cat:
                canon[cls] = cat
    return canon


def fmt_nigp_class(v) -> str:
    """AI output uses float-y class values (e.g. '961.0' or NaN). Normalize to 3-digit string or ''. """
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    if s.isdigit():
        s = s.zfill(3)
    return s


def main() -> None:
    print(f"Loading mapping: {MAPPING_PATH}")
    df = pd.read_csv(MAPPING_PATH, low_memory=False)
    n_total = len(df)

    print(f"Loading AI classifications: {AI_PATH}")
    ai = pd.read_csv(AI_PATH, low_memory=False)
    ai["_k"] = ai["description"].apply(normalize)
    ai = ai.drop_duplicates("_k", keep="first")
    ai_by_key = ai.set_index("_k")

    canon = load_canonical()
    print(f"Canonical NIGP→category map: {len(canon)} entries")

    # Identify rows still in review queue
    is_review = df["Review_Flag"] == "Yes"
    n_review_before = int(is_review.sum())
    print(f"\nReview-queue rows before resolver: {n_review_before:,} ({n_review_before/n_total*100:.2f}% of total)")

    df["_key"] = df["Description_Best"].apply(normalize)

    # Bucket: empty description (cannot resolve via AI)
    is_empty = is_review & (df["_key"] == "")

    # Bucket: has description AND key is in AI output
    can_ai_resolve = is_review & (df["_key"] != "") & (df["_key"].isin(ai_by_key.index))

    # Bucket: has description but not in AI output → top-up needed
    needs_topup = is_review & (df["_key"] != "") & (~df["_key"].isin(ai_by_key.index))

    print(f"  Empty description (cannot AI-resolve): {int(is_empty.sum()):,}")
    print(f"  Resolvable via existing AI output:     {int(can_ai_resolve.sum()):,}")
    print(f"  Needs AI top-up (no AI record yet):    {int(needs_topup.sum()):,}")

    # Apply AI-fill
    fill_idx = df.index[can_ai_resolve]
    keys = df.loc[fill_idx, "_key"]
    ai_subset = ai_by_key.loc[keys]

    # Build new values
    ai_cat = ai_subset["business_category"].fillna("").astype(str).values
    ai_cls_raw = ai_subset["nigp_class_3digit"].values
    ai_cls = [fmt_nigp_class(v) for v in ai_cls_raw]
    ai_conf = ai_subset["confidence"].fillna("").astype(str).values
    ai_reason = ai_subset["reason"].fillna("").astype(str).values

    # Apply canonical alignment: if AI's nigp_class has a canonical category, use that
    aligned_cat = []
    aligned_flag = []  # whether we corrected
    for cat, cls in zip(ai_cat, ai_cls):
        if cls and cls in canon and canon[cls] != cat:
            aligned_cat.append(canon[cls])
            aligned_flag.append(True)
        else:
            aligned_cat.append(cat)
            aligned_flag.append(False)
    n_aligned = sum(aligned_flag)

    df.loc[fill_idx, "Business_Category"] = aligned_cat
    df.loc[fill_idx, "NIGP_Class_3digit"] = ai_cls
    df.loc[fill_idx, "Classification_Method"] = "ai_assist"
    df.loc[fill_idx, "Classification_Confidence"] = ["AI-" + c for c in ai_conf]
    df.loc[fill_idx, "Review_Flag"] = "No"
    df.loc[fill_idx, "Classification_Reason"] = [
        f"AI-assist ({c}): {r[:240]}" for c, r in zip(ai_conf, ai_reason)
    ]
    # NIGP_Match_Level: if AI assigned a class, treat as 3digit; else empty
    df.loc[fill_idx, "NIGP_Match_Level"] = [("3digit" if c else "") for c in ai_cls]

    # Tag empty-description rows as unclassifiable and clear their Review_Flag.
    # NIGP_Class_3digit / NIGP_Item_5digit columns are float64 in the mapping CSV — leave them as NaN
    # (they already are for unclassified rows; setting "" raises LossySetitemError).
    empty_idx = df.index[is_empty]
    if len(empty_idx) > 0:
        df.loc[empty_idx, "Business_Category"] = "Unclassified — No Description"
        df.loc[empty_idx, "Classification_Method"] = "no_description"
        df.loc[empty_idx, "Classification_Confidence"] = "Not Classifiable"
        df.loc[empty_idx, "Review_Flag"] = "No"
        df.loc[empty_idx, "Classification_Reason"] = (
            "Description_Best was empty/null — no rule, account pattern, or AI fill applicable"
        )
        df.loc[empty_idx, "NIGP_Match_Level"] = ""

    # Export the list of uncovered uniques for the AI top-up step
    uncovered_keys = sorted(df.loc[needs_topup, "_key"].unique().tolist())
    pd.DataFrame({"description_norm": uncovered_keys}).to_csv(UNCOVERED_PATH, index=False)
    print(f"\nExported {len(uncovered_keys):,} uncovered unique descriptions → {UNCOVERED_PATH.relative_to(ROOT.parent)}")

    # Drop helper col before writing back
    df = df.drop(columns=["_key"])

    n_review_after = int((df["Review_Flag"] == "Yes").sum())
    n_filled = len(fill_idx)

    print(f"\nResults:")
    print(f"  AI-filled rows:                          {n_filled:,}")
    print(f"  Canonical-aligned rows (cat corrected):  {n_aligned:,}")
    print(f"  Review-queue rows after resolver:        {n_review_after:,} ({n_review_after/n_total*100:.2f}% of total)")
    print(f"    of which empty-description:            {int(is_empty.sum()):,}")
    print(f"    of which needs AI top-up:              {int(needs_topup.sum()):,}")

    print(f"\nWriting back: {MAPPING_PATH}")
    df.to_csv(MAPPING_PATH, index=False)

    # Refresh the review queue file
    rq = df[df["Review_Flag"] == "Yes"].copy()
    rq.to_csv(REVIEW_QUEUE_PATH, index=False)
    print(f"Refreshed review-queue file: {REVIEW_QUEUE_PATH} ({len(rq):,} rows)")


if __name__ == "__main__":
    main()
