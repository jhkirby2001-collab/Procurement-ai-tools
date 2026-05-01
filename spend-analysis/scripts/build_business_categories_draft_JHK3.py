"""
Build draft Business Category mapping (Level 1 of taxonomy).

Maps each of the 138 NIGP 3-digit Classes (extracted from EY raw data) to a
business-friendly Business Category. Output is a draft for JHK3 review.

Design principles (per project memory):
- ~12-20 categories, business-friendly, MECE.
- Every NIGP class lands in exactly one category.
- Categories defined by the KIND of thing being purchased, NOT spend dollars.
"""

from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
REF_DIR = ROOT / "data" / "reference"
REF_3D = REF_DIR / "nigp_codes_3digit_JHK3.csv"

CATEGORY_DEFINITIONS = {
    "Vehicles & Fleet": "Vehicles, vehicle parts, vehicle repair services, fuel-related taxes/fees, and watercraft. Anything that moves under its own power or attaches to a moving vehicle.",
    "Heavy Equipment & Machinery": "Purchased operating equipment (loaders, sweepers, lifts, sprayers, conveyors, outdoor power equipment). Excludes rental equipment and excludes vehicle-mounted accessories.",
    "Equipment Rental & Leasing": "Short-term rentals and multi-year leases of equipment, copiers, refrigerators, plotters, stages, linens, and traffic-control gear. Defined by the contract structure (rent/lease), not the asset type.",
    "Construction Materials": "Raw materials and pre-fabricated components installed into infrastructure: concrete, asphalt, steel, rebar, water/sewer pipe, manhole structures, brick, stakes.",
    "Construction & Trades Services": "Skilled construction labor and project services: bridge construction, expansion joint repair, contracted tradesmen.",
    "Facilities Operations & Maintenance": "Building systems and recurring building upkeep: HVAC parts, lighting, plumbing components, doors, pumps, insulation, batteries (facility), airport facility maintenance, water meter test equipment, generic facility hardware.",
    "Landscaping, Grounds & Irrigation": "Outdoor grounds care: landscape services, irrigation design/repair, integrated pest management, topsoil, beekeeping, rodenticide.",
    "Janitorial, Sanitation & Waste": "Cleaning supplies, trash liners, paper towels, wipers, floor machines, disposal/dumpster services, portable-toilet servicing, portable-toilet parts.",
    "Chemicals & Water Treatment": "Industrial and treatment chemicals: chlorine, sodium hexametaphosphate, phosphoric acid, ice/snow melt chemicals, runway deicer, hazmat handling.",
    "Public Safety, Uniforms & PPE": "Police/fire uniforms and accessories, badges, fire extinguishers, fire helmet markings, prison equipment, safety footwear, warning/reflective tape, SCBA breathing-air systems, bunker-gear laundry, institutional food for police lockups.",
    "Medical & Health Services": "Medical gases (and gas service), AEDs, lab spectrometers, medical exams, counseling/behavioral health services.",
    "Animal Care & Veterinary": "Live animals, animal feed (hay), animal litter, veterinary supplies and services, beekeeping (animal-husbandry side).",
    "IT, Telecom & Audio/Visual": "Network services, telephone equipment, telecom cable, in-car video systems, A/V maintenance, printer accessories/supplies, earphones.",
    "Office, Print & Marketing": "Office consumables, paper, envelopes, binding, signs/posters, advertising/PR, printing/embroidery, graphic-art supplies, promotional items.",
    "Furniture & Furnishings": "Office and institutional furniture, mattresses and mattress covers, fitness equipment, furniture repair/installation services.",
    "Professional & Administrative Services": "Advisory and back-office services not fitting elsewhere: management services, security guards, armored car, records storage, weather forecasting, testing/calibration, transportation/logistics, library labor, contract service fees.",
    "Grants & Pass-Through Funding": "Subgrant disbursements and pass-through funding to community-based organizations and subrecipients (DFSS social services, CDPH public health, BACP business development, DOH housing, DPD planning programs). Identified by Chicago FMPS subgrant account codes (220005, 220044, 220100, 220801, 220300, 220999) and/or department-program prefixes in the PO description (DFSS-/BACP-/CDPH-/DOH-/DPD-/DCASE-). Not a commodity purchase — no NIGP class applies.",
}

# (nigp_class_3digit, business_category, judgment_note)
# judgment_note populated ONLY where the call required interpretation
MAPPING = [
    ("010", "Facilities Operations & Maintenance", ""),
    ("020", "Heavy Equipment & Machinery", ""),
    ("025", "Facilities Operations & Maintenance", "Stationary breathing-air compressor system — facility install, not personal SCBA gear"),
    ("031", "Facilities Operations & Maintenance", ""),
    ("035", "Vehicles & Fleet", "Helicopters classified as fleet"),
    ("037", "Office, Print & Marketing", ""),
    ("040", "Animal Care & Veterinary", "Live animals (landscape grazing) — kept with other animal codes for coherence"),
    ("055", "Vehicles & Fleet", ""),
    ("060", "Vehicles & Fleet", ""),
    ("065", "Vehicles & Fleet", ""),
    ("070", "Vehicles & Fleet", ""),
    ("071", "Vehicles & Fleet", ""),
    ("073", "Vehicles & Fleet", ""),
    ("080", "Public Safety, Uniforms & PPE", ""),
    ("135", "Construction Materials", ""),
    ("150", "Facilities Operations & Maintenance", ""),
    ("155", "Janitorial, Sanitation & Waste", "Portable toilet parts — paired with portable-toilet waste service (977)"),
    ("175", "Chemicals & Water Treatment", ""),
    ("180", "Chemicals & Water Treatment", ""),
    ("190", "Chemicals & Water Treatment", ""),
    ("192", "Chemicals & Water Treatment", ""),
    ("195", "Facilities Operations & Maintenance", ""),
    ("200", "Public Safety, Uniforms & PPE", ""),
    ("201", "Public Safety, Uniforms & PPE", ""),
    ("207", "IT, Telecom & Audio/Visual", "Printer supplies treated as IT consumables"),
    ("210", "Construction Materials", ""),
    ("280", "IT, Telecom & Audio/Visual", ""),
    ("285", "Facilities Operations & Maintenance", ""),
    ("287", "Facilities Operations & Maintenance", "Battery packs for Spacesaver shelving — facility hardware"),
    ("305", "Office, Print & Marketing", ""),
    ("310", "Office, Print & Marketing", ""),
    ("315", "Facilities Operations & Maintenance", "Mastic adhesive — building/flooring use"),
    ("320", "Construction Materials", ""),
    ("325", "Animal Care & Veterinary", ""),
    ("340", "Public Safety, Uniforms & PPE", ""),
    ("345", "Public Safety, Uniforms & PPE", ""),
    ("350", "Facilities Operations & Maintenance", "Flagpoles are facility hardware"),
    ("360", "Public Safety, Uniforms & PPE", "Caution/warning tape — incident-scene use"),
    ("365", "Janitorial, Sanitation & Waste", "Floor scrubbers/sweepers — janitorial equipment"),
    ("375", "Public Safety, Uniforms & PPE", "Bread is for police lockup feeding — institutional food service tied to PD"),
    ("390", "Facilities Operations & Maintenance", "Bottled drinking water — general facility provisioning"),
    ("400", "Construction Materials", ""),
    ("405", "Vehicles & Fleet", "Fuel taxes/fees — operational pass-through tied to fleet fueling"),
    ("420", "Furniture & Furnishings", ""),
    ("425", "Furniture & Furnishings", ""),
    ("430", "Medical & Health Services", ""),
    ("450", "Facilities Operations & Maintenance", ""),
    ("465", "Medical & Health Services", ""),
    ("485", "Janitorial, Sanitation & Waste", ""),
    ("493", "Medical & Health Services", "Lab spectrometers — health/lab equipment"),
    ("495", "Animal Care & Veterinary", ""),
    ("525", "Professional & Administrative Services", "Library assembly labor — admin support"),
    ("530", "Public Safety, Uniforms & PPE", ""),
    ("540", "Construction Materials", ""),
    ("545", "Heavy Equipment & Machinery", ""),
    ("550", "Heavy Equipment & Machinery", ""),
    ("557", "Facilities Operations & Maintenance", ""),
    ("560", "Heavy Equipment & Machinery", ""),
    ("570", "Construction Materials", ""),
    ("590", "Public Safety, Uniforms & PPE", "Belt buckles — uniform accessory"),
    ("615", "Office, Print & Marketing", "Elmer's-style glue — office consumable, distinct from 315 mastic"),
    ("630", "Heavy Equipment & Machinery", ""),
    ("640", "Janitorial, Sanitation & Waste", ""),
    ("645", "Office, Print & Marketing", ""),
    ("658", "Construction Materials", ""),
    ("659", "Construction Materials", ""),
    ("665", "Janitorial, Sanitation & Waste", ""),
    ("670", "Construction Materials", ""),
    ("675", "Landscaping, Grounds & Irrigation", "Rodenticide — paired with 910 IPM as pest management"),
    ("680", "Public Safety, Uniforms & PPE", ""),
    ("685", "Professional & Administrative Services", "Shipping crates rental — logistics service"),
    ("699", "Professional & Administrative Services", ""),
    ("700", "Office, Print & Marketing", ""),
    ("720", "Facilities Operations & Maintenance", ""),
    ("725", "IT, Telecom & Audio/Visual", ""),
    ("735", "Janitorial, Sanitation & Waste", ""),
    ("745", "Construction Materials", ""),
    ("750", "Construction Materials", ""),
    ("755", "Construction Materials", ""),
    ("760", "Heavy Equipment & Machinery", ""),
    ("765", "Heavy Equipment & Machinery", "Street sweepers — kept with heavy equipment despite vehicle nature"),
    ("775", "Chemicals & Water Treatment", ""),
    ("785", "Office, Print & Marketing", "Crayons — closest fit"),
    ("790", "Landscaping, Grounds & Irrigation", ""),
    ("800", "Public Safety, Uniforms & PPE", "Safety footwear treated as PPE"),
    ("801", "Office, Print & Marketing", ""),
    ("803", "IT, Telecom & Audio/Visual", ""),
    ("805", "Furniture & Furnishings", "Fitness equipment — treated as facility furnishings"),
    ("810", "Vehicles & Fleet", "Vehicle-mounted salt spreader — fleet accessory"),
    ("815", "Construction & Trades Services", ""),
    ("825", "Animal Care & Veterinary", ""),
    ("832", "Public Safety, Uniforms & PPE", "Reflective tape — traffic-safety marking"),
    ("840", "IT, Telecom & Audio/Visual", "In-car video — police technology system"),
    ("850", "Furniture & Furnishings", ""),
    ("855", "Equipment Rental & Leasing", ""),
    ("875", "Animal Care & Veterinary", ""),
    ("885", "Chemicals & Water Treatment", ""),
    ("890", "Construction Materials", ""),
    ("906", "Landscaping, Grounds & Irrigation", ""),
    ("908", "Office, Print & Marketing", ""),
    ("909", "Facilities Operations & Maintenance", "Airport facility maintenance umbrella"),
    ("910", "Landscaping, Grounds & Irrigation", ""),
    ("912", "Landscaping, Grounds & Irrigation", ""),
    ("913", "Construction & Trades Services", ""),
    ("914", "Construction & Trades Services", ""),
    ("915", "Office, Print & Marketing", "Advertising/PR — marketing function"),
    ("918", "Professional & Administrative Services", ""),
    ("920", "IT, Telecom & Audio/Visual", ""),
    ("926", "Professional & Administrative Services", "Generic monthly service fee — admin contract"),
    ("928", "Vehicles & Fleet", ""),
    ("929", "Vehicles & Fleet", ""),
    ("931", "Furniture & Furnishings", ""),
    ("934", "Landscaping, Grounds & Irrigation", ""),
    ("936", "Public Safety, Uniforms & PPE", "Fire extinguisher inspection — life-safety"),
    ("938", "Medical & Health Services", ""),
    ("939", "IT, Telecom & Audio/Visual", ""),
    ("941", "Facilities Operations & Maintenance", "Water meter test bench — utility ops"),
    ("948", "Medical & Health Services", ""),
    ("952", "Medical & Health Services", "Behavioral health crisis service — health"),
    ("954", "Public Safety, Uniforms & PPE", "Bunker gear laundry — PPE care"),
    ("958", "Professional & Administrative Services", ""),
    ("959", "Vehicles & Fleet", "Tugboat maintenance — watercraft fleet"),
    ("961", "Professional & Administrative Services", ""),
    ("962", "Professional & Administrative Services", ""),
    ("964", "Professional & Administrative Services", "Security guards — contracted service"),
    ("966", "Office, Print & Marketing", ""),
    ("968", "Janitorial, Sanitation & Waste", ""),
    ("972", "Equipment Rental & Leasing", ""),
    ("975", "Equipment Rental & Leasing", "Traffic control rental — defined by rental contract structure"),
    ("977", "Janitorial, Sanitation & Waste", ""),
    ("979", "Equipment Rental & Leasing", ""),
    ("981", "Equipment Rental & Leasing", ""),
    ("983", "Equipment Rental & Leasing", ""),
    ("984", "Equipment Rental & Leasing", "Plotter/scanner lease — lease structure"),
    ("985", "Equipment Rental & Leasing", ""),
    ("988", "Landscaping, Grounds & Irrigation", ""),
    ("990", "Professional & Administrative Services", ""),
    ("992", "Professional & Administrative Services", ""),
]


def main() -> None:
    ref3 = pd.read_csv(REF_3D, dtype={"nigp_class_3digit": str})
    map_df = pd.DataFrame(
        MAPPING, columns=["nigp_class_3digit", "business_category", "judgment_note"]
    )

    assert len(map_df) == 138, f"Expected 138 mappings, got {len(map_df)}"
    assert map_df["nigp_class_3digit"].is_unique, "Duplicate NIGP classes in mapping"
    assigned = set(map_df["business_category"])
    defined = set(CATEGORY_DEFINITIONS)
    missing = assigned - defined
    unused = defined - assigned
    assert not missing, f"Categories used but not defined: {missing}"
    # "Grants & Pass-Through Funding" is intentionally not in the NIGP map —
    # it's keyed by FMPS account code, not commodity class.
    expected_unused = {"Grants & Pass-Through Funding"}
    surprise_unused = unused - expected_unused
    if surprise_unused:
        print(f"WARNING: unexpected categories with no NIGP classes: {surprise_unused}")
    if expected_unused & unused:
        print(f"NOTE: account-keyed categories (no NIGP classes by design): {expected_unused & unused}")

    merged = map_df.merge(
        ref3[["nigp_class_3digit", "canonical_description"]],
        on="nigp_class_3digit",
        how="left",
    ).rename(columns={"canonical_description": "nigp_class_canonical_desc_from_EY"})
    merged["business_category_definition"] = merged["business_category"].map(
        CATEGORY_DEFINITIONS
    )
    merged = merged[[
        "business_category",
        "business_category_definition",
        "nigp_class_3digit",
        "nigp_class_canonical_desc_from_EY",
        "judgment_note",
    ]].sort_values(["business_category", "nigp_class_3digit"]).reset_index(drop=True)

    out_detail = REF_DIR / "business_categories_JHK3.csv"
    merged.to_csv(out_detail, index=False)

    summary = (
        merged.groupby("business_category")
        .agg(
            class_count=("nigp_class_3digit", "count"),
            judgment_calls=("judgment_note", lambda s: (s != "").sum()),
            classes_listed=("nigp_class_3digit", lambda s: ", ".join(sorted(s))),
        )
        .reset_index()
    )
    # Add account-keyed categories (no NIGP classes mapped) so they appear in the summary
    for cat in expected_unused & unused:
        summary = pd.concat([summary, pd.DataFrame([{
            "business_category": cat,
            "class_count": 0,
            "judgment_calls": 0,
            "classes_listed": "(account-keyed — no NIGP classes)",
        }])], ignore_index=True)
    summary["definition"] = summary["business_category"].map(CATEGORY_DEFINITIONS)
    summary = summary[["business_category", "definition", "class_count", "judgment_calls", "classes_listed"]]
    summary = summary.sort_values("class_count", ascending=False).reset_index(drop=True)

    out_summary = REF_DIR / "business_categories_summary_JHK3.csv"
    summary.to_csv(out_summary, index=False)

    print(f"Detail mapping written: {len(merged):,} rows -> {out_detail}")
    print(f"Summary written:        {len(summary)} categories -> {out_summary}")
    print()
    print("=== Business Category summary (sorted by class count) ===")
    pd.set_option("display.max_colwidth", 60)
    pd.set_option("display.width", 200)
    print(summary[["business_category", "class_count", "judgment_calls"]].to_string(index=False))
    print()
    print(f"Total NIGP 3-digit classes mapped: {summary['class_count'].sum()} (expected 138)")
    print(f"Total judgment calls flagged for review: {summary['judgment_calls'].sum()}")


if __name__ == "__main__":
    main()
