"""
One-shot audit: run plain-English public-sector procurement phrases through
the classifier and report which ones fall to human review.

Goal: surface vocabulary gaps where the rule base (tuned on EY description
formats) fails on phrases a procurement staffer would actually type.

Usage:  python scripts/audit_classifier_coverage_JHK3.py
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

from classifier_JHK3 import classify_one, load_keyword_rules as load_rules

TEST_PHRASES = [
    ("Facilities Operations & Maintenance", "HVAC repair services"),
    ("Facilities Operations & Maintenance", "elevator maintenance"),
    ("Facilities Operations & Maintenance", "boiler service"),
    ("Facilities Operations & Maintenance", "plumbing repair"),
    ("Facilities Operations & Maintenance", "roofing repair"),

    ("Public Safety, Uniforms & PPE", "police uniforms"),
    ("Public Safety, Uniforms & PPE", "ballistic vests"),
    ("Public Safety, Uniforms & PPE", "PPE supplies"),
    ("Public Safety, Uniforms & PPE", "respirators"),
    ("Public Safety, Uniforms & PPE", "tactical gear"),

    ("Construction Materials", "asphalt"),
    ("Construction Materials", "concrete"),
    ("Construction Materials", "rebar"),
    ("Construction Materials", "lumber"),

    ("Vehicles & Fleet", "vehicle parts"),
    ("Vehicles & Fleet", "tires"),
    ("Vehicles & Fleet", "fuel"),
    ("Vehicles & Fleet", "auto repair services"),

    ("Professional & Administrative Services", "consulting services"),
    ("Professional & Administrative Services", "legal services"),
    ("Professional & Administrative Services", "training services"),
    ("Professional & Administrative Services", "background check services"),
    ("Professional & Administrative Services", "auditing services"),

    ("Office, Print & Marketing", "marketing and media services"),
    ("Office, Print & Marketing", "printing services"),
    ("Office, Print & Marketing", "graphic design services"),
    ("Office, Print & Marketing", "office supplies"),
    ("Office, Print & Marketing", "advertising services"),
    ("Office, Print & Marketing", "public relations consulting"),

    ("Equipment Rental & Leasing", "equipment rental"),
    ("Equipment Rental & Leasing", "scaffold rental"),
    ("Equipment Rental & Leasing", "generator rental"),

    ("Janitorial, Sanitation & Waste", "janitorial services"),
    ("Janitorial, Sanitation & Waste", "trash collection"),
    ("Janitorial, Sanitation & Waste", "cleaning supplies"),

    ("Heavy Equipment & Machinery", "bulldozer purchase"),
    ("Heavy Equipment & Machinery", "backhoe lease"),
    ("Heavy Equipment & Machinery", "skid steer loader"),

    ("IT, Telecom & Audio/Visual", "IT services"),
    ("IT, Telecom & Audio/Visual", "software license"),
    ("IT, Telecom & Audio/Visual", "computer hardware"),
    ("IT, Telecom & Audio/Visual", "cybersecurity services"),
    ("IT, Telecom & Audio/Visual", "telephone service"),

    ("Landscaping, Grounds & Irrigation", "landscaping services"),
    ("Landscaping, Grounds & Irrigation", "tree trimming"),
    ("Landscaping, Grounds & Irrigation", "lawn care"),
    ("Landscaping, Grounds & Irrigation", "irrigation system"),

    ("Chemicals & Water Treatment", "water treatment chemicals"),
    ("Chemicals & Water Treatment", "chlorine"),
    ("Chemicals & Water Treatment", "deicer"),

    ("Medical & Health Services", "medical supplies"),
    ("Medical & Health Services", "nursing services"),
    ("Medical & Health Services", "clinical lab services"),

    ("Animal Care & Veterinary", "veterinary services"),
    ("Animal Care & Veterinary", "animal feed"),

    ("Furniture & Furnishings", "office furniture"),
    ("Furniture & Furnishings", "office chairs"),
    ("Furniture & Furnishings", "conference room desks"),

    ("Construction & Trades Services", "electrical work"),
    ("Construction & Trades Services", "carpentry services"),
    ("Construction & Trades Services", "general contractor services"),

    ("Grants & Pass-Through Funding", "DFSS subgrant payment"),
    ("Grants & Pass-Through Funding", "community-based organization grant"),
]


def main() -> None:
    rules = load_rules()
    print(f"Loaded {len(rules):,} rules\n")
    print(f"{'EXPECTED CATEGORY':<42} | {'TYPED PHRASE':<38} | RESULT")
    print("-" * 130)

    fails: list[tuple[str, str, str]] = []
    miscat: list[tuple[str, str, str]] = []
    ok = 0

    for expected, phrase in TEST_PHRASES:
        result = classify_one(phrase, rules).as_dict()
        cat = result.get("Business_Category", "")
        review = result.get("Review_Flag", "")
        match = result.get("NIGP_Match_Level", "")
        confidence = result.get("Classification_Confidence", "")

        if review == "Yes" or not cat:
            status = "FAIL → review queue"
            fails.append((expected, phrase, status))
        elif cat != expected:
            status = f"MISCAT → {cat}"
            miscat.append((expected, phrase, cat))
        else:
            status = f"OK ({match}, {confidence})"
            ok += 1

        print(f"{expected[:40]:<42} | {phrase[:36]:<38} | {status}")

    total = len(TEST_PHRASES)
    print("\n" + "=" * 130)
    print(f"SUMMARY: {ok}/{total} OK | {len(fails)} fell through to review | {len(miscat)} mis-categorized")
    if fails:
        print("\nFAILED (review queue):")
        for exp, phrase, _ in fails:
            print(f"  {phrase!r:<42} (expected: {exp})")
    if miscat:
        print("\nMIS-CATEGORIZED (wrong category):")
        for exp, phrase, got in miscat:
            print(f"  {phrase!r:<42} expected: {exp:<40} got: {got}")


if __name__ == "__main__":
    main()
