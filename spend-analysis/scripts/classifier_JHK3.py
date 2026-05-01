"""
Description-driven classifier for Chicago procurement transactions.

Architecture (locked design):
- INPUTS: description text + Chicago FMPS account/object/fund codes ONLY.
  Vendor and EY-supplied NIGP codes are NOT inputs.
- Rule hierarchy: (1) keyword_rules → (2) account_patterns [TBD] → (3) AI assist [TBD]
- Same core function serves batch (whole file) and single-record (CLI) modes.

This is the rules-only first pass. AI assist hook is a TODO; rows that hit no rule
go to method='human_review' with Review_Flag=Yes.

Usage:
  Batch:         python classifier_JHK3.py --batch
  Single record: python classifier_JHK3.py --describe "REPAIR OF DIESEL GENERATORS"
"""

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RULES_PATH = ROOT / "data" / "reference" / "keyword_rules_DRAFT_JHK3.csv"
AI_RULES_PATH = ROOT / "data" / "reference" / "keyword_rules_from_ai_JHK3.csv"  # optional
ACCOUNT_PATTERNS_PATH = ROOT / "data" / "reference" / "account_patterns_DRAFT_JHK3.csv"
NIGP_3D_PATH = ROOT / "data" / "reference" / "nigp_codes_3digit_JHK3.csv"
SRC_PARQUET = ROOT / "data" / "processed" / "ey_raw_cache.parquet"

DESC_COLS = [
    "PO Description",
    "PO Item Description",
    "AP Invoice Line Description",
    "Invoice Distribution Description",
]
SPEND_COL = "PO Dist Amount Matched to AP"

LOWINFO_RE = re.compile(
    r"^\s*$|^(misc|miscellaneous|various|other|n/?a|none|unknown|tbd)\s*$|"
    r"^per\s+(contract|agreement|spec|specs|specification|po|order|invoice)\s*$|"
    r"^as\s+(needed|specified|per\s+spec)\s*$|"
    r"^(see|refer\s+to)\s+(attached|attachment|invoice|po|contract).*$|"
    r"^(item|line|charge)\s*[#:\-]?\s*\d+\s*$|^[\W_\d]+$",
    flags=re.IGNORECASE,
)

CONFIDENCE_BY_LEVEL = {"exact": 0.95, "broad": 0.80, "review": 0.40}
CONFIDENCE_LABEL = lambda s: "High" if s >= 0.85 else ("Medium" if s >= 0.65 else "Low")
REVIEW_THRESHOLD = 0.75  # locked default


@dataclass
class ClassificationResult:
    business_category: str = ""
    nigp_class_3digit: str = ""
    nigp_item_5digit: str = ""
    nigp_code_assigned: str = ""
    nigp_match_level: str = ""
    classification_confidence: str = ""
    confidence_score: float = 0.0
    classification_method: str = ""
    review_flag: str = "Yes"
    classification_reason: str = ""

    def as_dict(self) -> dict:
        return {
            "Business_Category": self.business_category,
            "NIGP_Class_3digit": self.nigp_class_3digit,
            "NIGP_Item_5digit": self.nigp_item_5digit,
            "NIGP_Code_Assigned": self.nigp_code_assigned,
            "NIGP_Match_Level": self.nigp_match_level,
            "Classification_Confidence": self.classification_confidence,
            "Confidence_Score": self.confidence_score,
            "Classification_Method": self.classification_method,
            "Review_Flag": self.review_flag,
            "Classification_Reason": self.classification_reason,
        }


def load_keyword_rules(path: Path = RULES_PATH,
                        ai_path: Path = AI_RULES_PATH) -> pd.DataFrame:
    """Load hand-curated rules. If AI-mined rules exist, append them.

    Hand-curated rules take priority (loaded first; same match_type still
    sorts by exact > starts_with > contains across both files).
    """
    df = pd.read_csv(
        path,
        dtype={"nigp_class_3digit": str, "nigp_item_5digit": str},
    ).fillna("")
    df["source"] = "curated"

    if ai_path.exists():
        ai_df = pd.read_csv(
            ai_path,
            dtype={"nigp_class_3digit": str, "nigp_item_5digit": str},
        ).fillna("")
        ai_df["source"] = "ai_mined"
        df = pd.concat([df, ai_df], ignore_index=True)

    for col in ("nigp_class_3digit", "nigp_item_5digit"):
        df[col] = df[col].astype(str).str.replace(r"\.0$", "", regex=True)
    order = {"exact": 0, "starts_with": 1, "contains": 2}
    df["__priority"] = df["match_type"].map(order)
    # Stable sort: by priority, then curated-before-AI, then file order
    source_order = {"curated": 0, "ai_mined": 1}
    df["__source_priority"] = df["source"].map(source_order)
    df = df.sort_values(["__priority", "__source_priority"]).reset_index(drop=True)
    df["pattern_upper"] = df["pattern"].astype(str).str.upper()
    return df


def load_account_patterns(path: Path = ACCOUNT_PATTERNS_PATH) -> pd.DataFrame:
    df = pd.read_csv(
        path,
        dtype={"account_code": str, "fund": str, "object_code": str,
               "nigp_class_3digit": str, "nigp_item_5digit": str},
    ).fillna("")
    for col in ("nigp_class_3digit", "nigp_item_5digit", "account_code"):
        df[col] = df[col].astype(str).str.replace(r"\.0$", "", regex=True)
    return df


def best_description(po_desc, item_desc, ap_line_desc, inv_dist_desc) -> str:
    for v in (po_desc, item_desc, ap_line_desc, inv_dist_desc):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            continue
        s = str(v).strip()
        if s and not LOWINFO_RE.match(s):
            return s
    return ""


def classify_one(
    description: str,
    rules_df: pd.DataFrame,
    account_code: Optional[str] = None,
    fund: Optional[str] = None,
    object_code: Optional[str] = None,
) -> ClassificationResult:
    """Single-record classification. Pure function — no I/O."""
    result = ClassificationResult()
    desc = (description or "").strip()
    if not desc:
        result.classification_method = "human_review"
        result.classification_reason = "No usable description"
        return result

    desc_u = desc.upper()
    for _, rule in rules_df.iterrows():
        pat_u = rule["pattern_upper"]
        mt = rule["match_type"]
        if mt == "exact":
            hit = desc_u == pat_u
        elif mt == "starts_with":
            hit = desc_u.startswith(pat_u)
        else:
            hit = pat_u in desc_u
        if hit:
            level = rule["nigp_match_level"]
            score = CONFIDENCE_BY_LEVEL.get(level, 0.50)
            result.business_category = rule["business_category"]
            result.nigp_class_3digit = str(rule["nigp_class_3digit"])
            result.nigp_item_5digit = str(rule["nigp_item_5digit"])
            result.nigp_code_assigned = (
                result.nigp_item_5digit or result.nigp_class_3digit
            )
            result.nigp_match_level = level
            result.confidence_score = score
            result.classification_confidence = CONFIDENCE_LABEL(score)
            result.classification_method = "keyword_rule"
            result.review_flag = (
                "Yes" if (level == "review" or score < REVIEW_THRESHOLD) else "No"
            )
            result.classification_reason = (
                f"matched keyword rule '{rule['pattern']}' ({mt})"
                + (f" — {rule['notes']}" if rule.get("notes") else "")
            )
            return result

    # Pass 2: account_patterns (Chicago FMPS supplemental signal)
    if account_code:
        try:
            account_df = load_account_patterns()
        except FileNotFoundError:
            account_df = None
        if account_df is not None:
            for _, ar in account_df.iterrows():
                if str(account_code) != ar["account_code"]:
                    continue
                if ar["fund"] and str(fund or "") != ar["fund"]:
                    continue
                if ar["object_code"] and str(object_code or "") != ar["object_code"]:
                    continue
                level = ar["nigp_match_level"]
                score = CONFIDENCE_BY_LEVEL.get(level, 0.50)
                result.business_category = ar["business_category"]
                result.nigp_class_3digit = str(ar["nigp_class_3digit"])
                result.nigp_item_5digit = str(ar["nigp_item_5digit"])
                result.nigp_code_assigned = (
                    result.nigp_item_5digit or result.nigp_class_3digit
                )
                result.nigp_match_level = level
                result.confidence_score = score
                result.classification_confidence = CONFIDENCE_LABEL(score)
                result.classification_method = "account_pattern"
                result.review_flag = (
                    "Yes" if (level == "review" or score < REVIEW_THRESHOLD) else "No"
                )
                result.classification_reason = (
                    f"matched account-code pattern '{ar['account_code']}'"
                    + (f" — {ar['notes']}" if ar.get("notes") else "")
                )
                return result

    # TODO: AI assist hook (when AI module is added)

    result.classification_method = "human_review"
    result.classification_reason = "No keyword rule or account pattern matched (AI assist not yet wired)"
    return result


def classify_batch(df: pd.DataFrame, rules_df: pd.DataFrame,
                    account_df: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    """Vectorized batch classification on a DataFrame with EY-style columns."""
    n = len(df)
    print(f"Building Description_Best across {n:,} rows ...")
    desc = df.apply(
        lambda r: best_description(
            r["PO Description"], r["PO Item Description"],
            r["AP Invoice Line Description"], r["Invoice Distribution Description"],
        ),
        axis=1,
    )
    desc_u = desc.str.upper()

    out = pd.DataFrame({
        "Business_Category": [""] * n,
        "NIGP_Class_3digit": [""] * n,
        "NIGP_Item_5digit": [""] * n,
        "NIGP_Code_Assigned": [""] * n,
        "NIGP_Match_Level": [""] * n,
        "Classification_Confidence": [""] * n,
        "Confidence_Score": [0.0] * n,
        "Classification_Method": [""] * n,
        "Review_Flag": ["Yes"] * n,
        "Classification_Reason": [""] * n,
    }, index=df.index)

    # Mask: which rows still need a classification
    unclassified = pd.Series([True] * n, index=df.index)
    rule_hits = []  # (pattern, n_hit) for the coverage report

    print(f"Applying {len(rules_df):,} rules (priority: exact > starts_with > contains) ...")
    for _, rule in rules_df.iterrows():
        if not unclassified.any():
            break
        pat_u = rule["pattern_upper"]
        mt = rule["match_type"]
        if mt == "exact":
            mask = unclassified & (desc_u == pat_u)
        elif mt == "starts_with":
            mask = unclassified & desc_u.str.startswith(pat_u, na=False)
        else:
            mask = unclassified & desc_u.str.contains(re.escape(pat_u), regex=True, na=False)

        n_hit = int(mask.sum())
        rule_hits.append({
            "pattern": rule["pattern"],
            "match_type": mt,
            "business_category": rule["business_category"],
            "nigp_match_level": rule["nigp_match_level"],
            "n_hit": n_hit,
        })
        if n_hit == 0:
            continue

        level = rule["nigp_match_level"]
        score = CONFIDENCE_BY_LEVEL.get(level, 0.50)
        review = "Yes" if (level == "review" or score < REVIEW_THRESHOLD) else "No"
        nigp_code = str(rule["nigp_item_5digit"]) or str(rule["nigp_class_3digit"])
        reason_base = f"matched keyword rule '{rule['pattern']}' ({mt})"
        if rule.get("notes"):
            reason_base += f" — {rule['notes']}"

        out.loc[mask, "Business_Category"] = rule["business_category"]
        out.loc[mask, "NIGP_Class_3digit"] = str(rule["nigp_class_3digit"])
        out.loc[mask, "NIGP_Item_5digit"] = str(rule["nigp_item_5digit"])
        out.loc[mask, "NIGP_Code_Assigned"] = nigp_code
        out.loc[mask, "NIGP_Match_Level"] = level
        out.loc[mask, "Classification_Confidence"] = CONFIDENCE_LABEL(score)
        out.loc[mask, "Confidence_Score"] = score
        out.loc[mask, "Classification_Method"] = "keyword_rule"
        out.loc[mask, "Review_Flag"] = review
        out.loc[mask, "Classification_Reason"] = reason_base

        unclassified &= ~mask

    # Pass 2: account_patterns (Chicago FMPS supplemental signal)
    account_hits = []
    if account_df is not None and "Account" in df.columns:
        accounts_str = df["Account"].astype("Int64").astype(str).fillna("")
        funds_str = df["Fund"].astype(str).fillna("") if "Fund" in df.columns else pd.Series([""] * n, index=df.index)
        print(f"Applying {len(account_df)} account patterns to remaining {unclassified.sum():,} rows ...")
        for _, ar in account_df.iterrows():
            if not unclassified.any():
                break
            mask = unclassified & (accounts_str == str(ar["account_code"]))
            if ar["fund"]:
                mask &= (funds_str == str(ar["fund"]))
            n_hit = int(mask.sum())
            account_hits.append({
                "account_code": ar["account_code"],
                "fund": ar["fund"],
                "business_category": ar["business_category"],
                "nigp_match_level": ar["nigp_match_level"],
                "n_hit": n_hit,
            })
            if n_hit == 0:
                continue

            level = ar["nigp_match_level"]
            score = CONFIDENCE_BY_LEVEL.get(level, 0.50)
            review = "Yes" if (level == "review" or score < REVIEW_THRESHOLD) else "No"
            nigp_code = str(ar["nigp_item_5digit"]) or str(ar["nigp_class_3digit"])
            reason_base = f"matched account-code pattern '{ar['account_code']}'"
            if ar.get("notes"):
                reason_base += f" — {ar['notes']}"

            out.loc[mask, "Business_Category"] = ar["business_category"]
            out.loc[mask, "NIGP_Class_3digit"] = str(ar["nigp_class_3digit"])
            out.loc[mask, "NIGP_Item_5digit"] = str(ar["nigp_item_5digit"])
            out.loc[mask, "NIGP_Code_Assigned"] = nigp_code
            out.loc[mask, "NIGP_Match_Level"] = level
            out.loc[mask, "Classification_Confidence"] = CONFIDENCE_LABEL(score)
            out.loc[mask, "Confidence_Score"] = score
            out.loc[mask, "Classification_Method"] = "account_pattern"
            out.loc[mask, "Review_Flag"] = review
            out.loc[mask, "Classification_Reason"] = reason_base

            unclassified &= ~mask

    # Unclassified = no rule fired
    no_rule = unclassified & (desc != "")
    no_desc = unclassified & (desc == "")
    out.loc[no_rule, "Classification_Method"] = "human_review"
    out.loc[no_rule, "Classification_Reason"] = "No keyword rule or account pattern matched (AI assist not yet wired)"
    out.loc[no_desc, "Classification_Method"] = "human_review"
    out.loc[no_desc, "Classification_Reason"] = "No usable description across all 4 columns"

    out["Description_Best"] = desc
    return out, pd.DataFrame(rule_hits), pd.DataFrame(account_hits)


def assemble_lean_output(raw_df: pd.DataFrame, classified: pd.DataFrame) -> pd.DataFrame:
    """13-column lean output per locked spec."""
    lean = pd.DataFrame({
        "Transaction_ID": raw_df.index.astype(str),  # row-index proxy until composite ID is decided
        "Date": raw_df["PO Creation Date"],
        "Department": raw_df["PO Lead Dept"].astype("Int64").astype(str),
        "Vendor": raw_df["Vendor Name"],
        "Description_Best": classified["Description_Best"],
        "Amount": raw_df[SPEND_COL],
    })
    classified_cols = classified.drop(columns=["Description_Best"])
    return pd.concat([lean, classified_cols], axis=1)


def run_batch() -> None:
    rules = load_keyword_rules()
    print(f"Loaded {len(rules)} keyword rules")
    try:
        account_df = load_account_patterns()
        print(f"Loaded {len(account_df)} account patterns")
    except FileNotFoundError:
        account_df = None
        print("No account_patterns file found — skipping account pass")

    needed_cols = DESC_COLS + ["Vendor Name", "PO Creation Date", "PO Lead Dept",
                                "Account", "Fund", SPEND_COL]
    df = pd.read_parquet(SRC_PARQUET, columns=needed_cols)
    print(f"Loaded {len(df):,} rows from {SRC_PARQUET.name}")

    classified, rule_hits, account_hits = classify_batch(df, rules, account_df)
    final = assemble_lean_output(df, classified)

    out_dir = ROOT / "data" / "processed"
    deliverable_dir = ROOT.parent / "outputs"
    deliverable_dir.mkdir(exist_ok=True)
    out_full = deliverable_dir / "NIGP_Mapping_JHK3.csv"
    out_review = deliverable_dir / "NIGP_Mapping_Review_Queue_JHK3.csv"
    out_coverage = out_dir / "classifier_coverage_report_JHK3.csv"
    out_account = out_dir / "classifier_account_pattern_report_JHK3.csv"

    final.to_csv(out_full, index=False)
    review = final[final["Review_Flag"] == "Yes"]
    review.to_csv(out_review, index=False)
    rule_hits.sort_values("n_hit", ascending=False).to_csv(out_coverage, index=False)
    if not account_hits.empty:
        account_hits.sort_values("n_hit", ascending=False).to_csv(out_account, index=False)

    n = len(final)
    n_review = len(review)
    method_counts = final["Classification_Method"].value_counts()
    cat_counts = final[final["Classification_Method"].isin(["keyword_rule", "account_pattern"])]["Business_Category"].value_counts()

    print()
    print("=" * 72)
    print("CLASSIFICATION COMPLETE")
    print("=" * 72)
    print(f"Total rows:                       {n:,}")
    print(f"Classified by keyword rule:       {method_counts.get('keyword_rule', 0):,}")
    print(f"Classified by account pattern:    {method_counts.get('account_pattern', 0):,}")
    print(f"Sent to human review:             {method_counts.get('human_review', 0):,}")
    print(f"Review_Flag=Yes (review queue):   {n_review:,} ({n_review/n*100:.1f}% of total)")
    print()
    print("Outputs:")
    print(f"  Full lean:      {out_full}")
    print(f"  Review queue:   {out_review}")
    print(f"  Rule coverage:  {out_coverage}")
    if not account_hits.empty:
        print(f"  Account hits:   {out_account}")
    print()
    print("=== Business Category distribution (rule + account-pattern classified rows) ===")
    print(cat_counts.to_string())


def run_single(description: str) -> None:
    rules = load_keyword_rules()
    result = classify_one(description, rules)
    d = result.as_dict()
    print()
    print("=" * 72)
    print("SINGLE-RECORD CLASSIFICATION")
    print("=" * 72)
    print(f"Description: {description!r}")
    print()
    for k, v in d.items():
        print(f"  {k:30s}  {v}")


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--batch", action="store_true", help="Run batch over full EY parquet")
    p.add_argument("--describe", type=str, help="Single-record: classify one description string")
    args = p.parse_args()
    if args.batch:
        run_batch()
    elif args.describe:
        run_single(args.describe)
    else:
        p.print_help()


if __name__ == "__main__":
    main()
