"""
Draft account_patterns_JHK3.csv — supplemental classifier signal from
Chicago FMPS account/object/fund codes.

Purpose: catch transactions where description text alone is ambiguous, using
the account-code as the secondary signal. Per locked design, account_patterns
fire AFTER keyword_rules — only when no keyword rule has matched.

First-pass scope: Subgrant Account Codes
- The 220xxx series in Chicago FMPS includes a set of accounts used almost
  exclusively for subgrant / pass-through funding to community-based
  organizations. Empirical analysis of the EY raw data shows accounts 220005,
  220044, 220100, 220801, 220300, 220999 carry 93-100% program-tagged rows
  (DFSS-/BACP-/CDPH- subgrants). Account 220140 is mixed (17% program-tag,
  83% other) — excluded from this pattern set to avoid false positives.

Future expansion: as we identify other clean account-to-commodity mappings
(e.g., 220157 might map to a specific commodity area), they get added here.
"""

from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "data" / "reference"

# Each rule:
# (account_code, fund, object_code, business_category, nigp_class_3digit,
#  nigp_item_5digit, nigp_match_level, notes)
# Empty string for fund/object means "any". account_code is required.
PATTERNS = [
    # Subgrant / pass-through accounts (>=93% program-tag in EY data)
    ("220005", "", "", "Grants & Pass-Through Funding", "", "", "broad",
     "Subgrant disbursement account (99% program-tagged in EY data)"),
    ("220044", "", "", "Grants & Pass-Through Funding", "", "", "broad",
     "Subgrant disbursement account (99% program-tagged)"),
    ("220100", "", "", "Grants & Pass-Through Funding", "", "", "broad",
     "Subgrant disbursement account (96% program-tagged)"),
    ("220801", "", "", "Grants & Pass-Through Funding", "", "", "broad",
     "Subgrant disbursement account (100% program-tagged)"),
    ("220300", "", "", "Grants & Pass-Through Funding", "", "", "broad",
     "Subgrant disbursement account (93% program-tagged)"),
    ("220999", "", "", "Grants & Pass-Through Funding", "", "", "broad",
     "Subgrant misc disbursement account"),
]


def main() -> None:
    df = pd.DataFrame(PATTERNS, columns=[
        "account_code", "fund", "object_code",
        "business_category", "nigp_class_3digit", "nigp_item_5digit",
        "nigp_match_level", "notes",
    ])
    out = OUT_DIR / "account_patterns_DRAFT_JHK3.csv"
    df.to_csv(out, index=False)
    print(f"Drafted {len(df)} account patterns -> {out}")
    print(df.to_string(index=False))


if __name__ == "__main__":
    main()
