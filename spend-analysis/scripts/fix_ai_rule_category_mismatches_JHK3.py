"""
One-shot fix: bulk-correct business_category in keyword_rules_from_ai_JHK3.csv
where the AI assigned a category that disagrees with the canonical mapping in
business_categories_JHK3.csv. NIGP class stays unchanged.

Run from repo root:
  python spend-analysis/scripts/fix_ai_rule_category_mismatches_JHK3.py
"""
from __future__ import annotations

import csv
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CANON_PATH = ROOT / "data" / "reference" / "business_categories_JHK3.csv"
AI_RULES_PATH = ROOT / "data" / "reference" / "keyword_rules_from_ai_JHK3.csv"


def main() -> None:
    canon: dict[str, str] = {}
    with open(CANON_PATH, newline="") as f:
        for row in csv.DictReader(f):
            cls = str(row.get("nigp_class_3digit", "")).strip()
            cat = row.get("business_category", "").strip()
            if cls:
                canon[cls] = cat
    print(f"Loaded canonical mapping: {len(canon)} NIGP classes → category")

    with open(AI_RULES_PATH, newline="") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames
        rows = list(reader)
    total = len(rows)

    fixed = 0
    note_suffix = "  [auto-corrected 2026-05-06: business_category aligned to canonical mapping]"
    for row in rows:
        cls = str(row.get("nigp_class_3digit", "")).strip()
        cat = row.get("business_category", "").strip()
        if cls and cls in canon and canon[cls] != cat:
            row["business_category"] = canon[cls]
            existing_note = row.get("notes", "") or ""
            if note_suffix.strip() not in existing_note:
                row["notes"] = existing_note + note_suffix
            fixed += 1

    with open(AI_RULES_PATH, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, quoting=csv.QUOTE_MINIMAL)
        writer.writeheader()
        writer.writerows(rows)

    print(f"Rewrote {AI_RULES_PATH}")
    print(f"  total rules: {total:,}")
    print(f"  rules corrected: {fixed:,}")
    print(f"  rules unchanged: {total - fixed:,}")


if __name__ == "__main__":
    main()
