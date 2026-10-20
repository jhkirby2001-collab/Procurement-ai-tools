"""
Regression test for water-utility (DWM) core items: each phrase must land in the
expected Business Category AND resolve to the expected supplier NAICS code
(via its NIGP item/class and nigp_naics_crosswalk_JHK3.csv).

Added 2026-10-20 with the DWM water-items rules pass. Run after any edit to
keyword_rules_DRAFT_JHK3.csv or the crosswalk:

    python scripts/audit_dwm_water_items_JHK3.py
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from classifier_JHK3 import classify_one, load_keyword_rules  # noqa: E402
import spend_report_JHK3 as sr  # noqa: E402

# (description, expected Business Category, expected supplier NAICS)
CASES = [
 ("WATER MAIN CONSTRUCTION", "Construction & Trades Services", "237110"),
 ("WATER MAIN REPLACEMENT PROGRAM", "Construction & Trades Services", "237110"),
 ("SEWER MAIN LINING", "Construction & Trades Services", "237110"),
 ("SEWER IMPROVEMENT PROJECT", "Construction & Trades Services", "237110"),
 ("LEAD SERVICE LINE REPLACEMENT", "Construction & Trades Services", "237110"),
 ("WATER SERVICE LINE REPAIR", "Construction & Trades Services", "237110"),
 ("PARKWAY RESTORATION AFTER UTILITY WORK", "Construction & Trades Services", "237310"),
 ("DUCTILE IRON WATER PIPE 8 INCH", "Construction Materials", "423720"),
 ("DUCTILE IRON PIPE CLASS 52", "Construction Materials", "423720"),
 ("DUCTILE IRON WATER MAIN PIPE 12 INCH", "Construction Materials", "423720"),
 ("GATE VALVE 8 INCH MECHANICAL JOINT", "Construction Materials", "423720"),
 ("BUTTERFLY VALVE 24 INCH", "Construction Materials", "423720"),
 ("FIRE HYDRANT CHICAGO STANDARD", "Construction Materials", "423720"),
 ("HYDRANT REPAIR PARTS", "Construction Materials", "423720"),
 ("CURB STOP 1 INCH BRASS", "Construction Materials", "423720"),
 ("CURB BOX AND LID", "Construction Materials", "423720"),
 ("TAPPING SLEEVE 12 X 8", "Construction Materials", "423720"),
 ("PIPE REPAIR CLAMP STAINLESS 8 IN", "Construction Materials", "423720"),
 ("MEGALUG PIPE RESTRAINT 16 IN", "Construction Materials", "423720"),
 ("COPPER TUBING TYPE K 1 INCH", "Construction Materials", "423720"),
 ("MECHANICAL JOINT TEE 12 X 6", "Construction Materials", "423720"),
 ("WATER METER 5/8 INCH", "Construction Materials", "423720"),
 ("PRECAST MANHOLE RISER 48 INCH", "Construction Materials", "327390"),
 ("PRECAST CONCRETE MANHOLE", "Construction Materials", "327390"),
 ("MANHOLE BASE PRECAST", "Construction Materials", "327390"),
 ("PRECAST CATCH BASIN", "Construction Materials", "327390"),
 ("MANHOLE FRAME AND LID", "Construction Materials", "331511"),
 ("CATCH BASIN FRAME AND GRATE", "Construction Materials", "331511"),
 ("SODIUM HYPOCHLORITE BULK DELIVERY", "Chemicals & Water Treatment", "424690"),
 ("LIQUID CHLORINE ONE TON CYLINDERS", "Chemicals & Water Treatment", "424690"),
 ("HYDROFLUOSILICIC ACID", "Chemicals & Water Treatment", "424690"),
 ("FLUOROSILICIC ACID TANK TRUCK", "Chemicals & Water Treatment", "424690"),
 ("BLENDED PHOSPHATE CORROSION INHIBITOR", "Chemicals & Water Treatment", "424690"),
 ("ORTHOPHOSPHATE FOR CORROSION CONTROL", "Chemicals & Water Treatment", "424690"),
 ("POLYMER FOR WATER TREATMENT", "Chemicals & Water Treatment", "424690"),
 ("CAUSTIC SODA 50 PERCENT", "Chemicals & Water Treatment", "424690"),
 ("HYDRATED LIME BULK", "Chemicals & Water Treatment", "424690"),
 ("AQUA AMMONIA", "Chemicals & Water Treatment", "424690"),
 ("ACTIVATED CARBON POWDERED", "Chemicals & Water Treatment", "424690"),
 ("ALUMINUM SULFATE LIQUID ALUM", "Chemicals & Water Treatment", "424690"),
]


def main() -> int:
    rules = load_keyword_rules()
    fails = 0
    for text, cat, naics in CASES:
        d = classify_one(text, rules).as_dict()
        got = sr.naics_for(d["NIGP_Item_5digit"], d["NIGP_Class_3digit"])["NAICS_Code"]
        ok = d["Business_Category"] == cat and got == naics
        fails += not ok
        print(f"{'OK ' if ok else 'XX '} {text[:40]:<40} | {d['Business_Category'][:28]:<28} | "
              f"NIGP {d['NIGP_Code_Assigned'] or '-':<5} | NAICS {got or '-'}"
              + ("" if ok else f"   (expected {cat} / {naics})"))
    print(f"\nSUMMARY: {len(CASES) - fails}/{len(CASES)} OK")
    return 1 if fails else 0


if __name__ == "__main__":
    raise SystemExit(main())
