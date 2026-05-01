"""
Draft keyword_rules_JHK3.csv for the description-driven classifier.

Strategy: hybrid match types
- 'contains'    — case-insensitive substring match (the bulk of rules)
- 'starts_with' — handles department program-code prefixes (DFSS-, BACP-, etc.)
- 'exact'       — for short cryptic strings that need precise match

Precedence at classify time: exact > starts_with > contains. First rule to fire wins.

NIGP class numbers are pulled from our 138-class working reference (derived from
EY data). Where no clean class exists in the working set (legal services,
environmental services, etc.), the rule sets nigp_match_level='review' or 'broad'
and the row goes to Review_Flag=Yes.

Department program-code prefixes (DFSS-, BACP-, CDPH-, DOH-, DPD-) describe WHO
paid, not WHAT was bought. They cannot be classified from description alone —
they need the supplemental account/object/fund signal. These rules set
match_level='review' and route to human review, per the locked design where
description-only classification is honest about its limits.
"""

from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "data" / "reference"

# Each rule: (pattern, match_type, business_category, nigp_class_3digit,
#             nigp_item_5digit, nigp_match_level, notes)
# match_type: 'exact' | 'starts_with' | 'contains'
# nigp_match_level: 'exact' (5-digit defensible) | 'broad' (only 3-digit) | 'review'
RULES = [
    # ===== Department program-code prefixes -> Grants & Pass-Through Funding =====
    # Confirmed via FMPS account-code analysis: these prefixes correlate 93-100%
    # with subgrant accounts (220005, 220044, 220100, 220801, 220300). They are
    # subgrant disbursements to community-based organizations, not commodity buys.
    ("DFSS-", "starts_with", "Grants & Pass-Through Funding", "", "", "broad",
     "DFSS = Family & Support Services subgrants to CBOs"),
    ("BACP-", "starts_with", "Grants & Pass-Through Funding", "", "", "broad",
     "BACP = Business Affairs & Consumer Protection subgrants"),
    ("CDPH-", "starts_with", "Grants & Pass-Through Funding", "", "", "broad",
     "CDPH = Chicago Dept of Public Health subgrants (Ryan White, etc.)"),
    ("DOH-", "starts_with", "Grants & Pass-Through Funding", "", "", "broad",
     "DOH = Dept of Housing subgrants"),
    ("DPD-", "starts_with", "Grants & Pass-Through Funding", "", "", "broad",
     "DPD = Dept of Planning & Development subgrants"),
    ("DCASE-", "starts_with", "Grants & Pass-Through Funding", "", "", "broad",
     "DCASE = Dept of Cultural Affairs & Special Events subgrants"),

    # ===== Equipment Rental & Leasing =====
    ("RENTAL OF HEAVY EQUIPMENT", "contains", "Equipment Rental & Leasing", "975", "", "broad",
     "Heavy equipment rental contracts (loaders, excavators, etc.)"),
    ("RENTAL AND PLACEMENT OF TRAFFIC CONTROL", "contains", "Equipment Rental & Leasing", "975", "97534", "exact",
     "Traffic control devices rental"),
    ("RENTAL OF AUTOMOBILES", "contains", "Equipment Rental & Leasing", "975", "97586", "exact",
     "Vehicle rental"),
    ("RENTAL OF MOBILE OFFICE TRAILERS", "contains", "Equipment Rental & Leasing", "975", "", "broad", ""),
    ("RENTAL OF TENTS", "contains", "Equipment Rental & Leasing", "855", "", "broad", ""),
    ("RENTAL OF ITEMS FOR CITY FESTIVALS", "contains", "Equipment Rental & Leasing", "855", "", "broad",
     "Festival/event equipment rentals"),
    ("LEASING OF AUTOMOBILES", "contains", "Equipment Rental & Leasing", "975", "97586", "exact", ""),
    ("LEASE OF COPIER", "contains", "Equipment Rental & Leasing", "985", "98527", "exact", ""),
    ("LEASED HIGH VOLUME", "contains", "Equipment Rental & Leasing", "985", "", "broad",
     "Production copier lease"),

    # ===== Janitorial, Sanitation & Waste =====
    ("JANITORIAL", "contains", "Janitorial, Sanitation & Waste", "485", "", "broad", ""),
    ("CUSTODIAL SERVICES", "contains", "Janitorial, Sanitation & Waste", "485", "", "broad", ""),
    ("STEEL DUMPSTERS", "contains", "Janitorial, Sanitation & Waste", "968", "96871", "exact", ""),
    ("WASTE DISPOSAL", "contains", "Janitorial, Sanitation & Waste", "968", "", "broad", ""),
    ("DISPOSAL SERVICES", "contains", "Janitorial, Sanitation & Waste", "968", "", "broad", ""),
    ("PORTABLE CHEMICAL TOILETS", "contains", "Janitorial, Sanitation & Waste", "977", "", "broad",
     "Portable toilet rental + servicing"),
    ("HAULING OF SEWER DEBRIS", "contains", "Janitorial, Sanitation & Waste", "968", "", "broad", ""),

    # ===== Construction Materials =====
    ("READY MIX CONCRETE", "contains", "Construction Materials", "750", "75070", "exact", ""),
    ("READY-MIX CONCRETE", "contains", "Construction Materials", "750", "75070", "exact", ""),
    ("ASPHALT PRIMERS, HOT MIX", "contains", "Construction Materials", "745", "", "broad", ""),
    ("HOT MIX ASPHALT", "contains", "Construction Materials", "745", "", "broad", ""),
    ("REINFORCED CONCRETE SEWER PIPE", "contains", "Construction Materials", "658", "", "broad", ""),
    ("CLAY SEWER PIPE", "contains", "Construction Materials", "658", "", "broad", ""),
    ("PIPES, FITTINGS, VALVES", "contains", "Construction Materials", "670", "", "broad",
     "Generic pipe/fittings — water-side default; could be sewer"),
    ("CONSTRUCTION MATERIALS AND SUPPLIES", "contains", "Construction Materials", "", "", "broad",
     "Generic — broad class only"),
    ("REPROCESSABLE MATERIALS FOR CONSTRUCTION", "contains", "Construction Materials", "", "", "broad", ""),
    ("AGGREGATES", "contains", "Construction Materials", "", "", "broad", ""),
    ("THERMOPLASTIC PAVEMENT MARKINGS", "contains", "Construction Materials", "745", "", "broad", ""),

    # ===== Construction & Trades Services =====
    ("TRADESMEN SERVICES", "contains", "Construction & Trades Services", "914", "", "broad", ""),
    ("BRIDGE CONSTRUCTION", "contains", "Construction & Trades Services", "913", "", "broad", ""),
    ("CONSTRUCTION MANAGEMENT", "contains", "Construction & Trades Services", "", "", "broad",
     "Construction-mgmt-at-risk; could be Pro Services"),
    ("INSPECTION AND CERTIFICATION SERVICES, FIRE EXTINGUISHER", "contains", "Public Safety, Uniforms & PPE", "936", "", "broad", ""),

    # ===== Vehicles & Fleet =====
    ("FUEL SUPPLY", "contains", "Vehicles & Fleet", "405", "", "broad", "Bulk fuel contracts"),
    ("GASOLINE", "contains", "Vehicles & Fleet", "405", "", "broad", ""),
    ("ULTRA LOW SULFUR DIESEL", "contains", "Vehicles & Fleet", "405", "", "broad", ""),
    ("BIO-DIESEL", "contains", "Vehicles & Fleet", "405", "", "broad", ""),
    ("E-10 & E-85 ETHANOL", "contains", "Vehicles & Fleet", "405", "", "broad", ""),
    ("JET FUEL", "contains", "Vehicles & Fleet", "405", "", "broad", "Aviation fuel"),
    ("PARTS AND REPAIR SERVICES - MECHANICAL", "contains", "Vehicles & Fleet", "928", "", "broad",
     "City fleet repair — vehicle body/mech work"),
    ("BODY REPAIRS FOR CITY-OWNED AUTOMOBILES", "contains", "Vehicles & Fleet", "928", "", "broad", ""),
    ("VEHICLE & EQUIPMENT PARTS", "contains", "Vehicles & Fleet", "060", "", "broad", ""),
    ("REPAIR OF SPRINGS AND SUSPENSION", "contains", "Vehicles & Fleet", "060", "", "broad", ""),
    ("PICK-UP TRUCKS", "contains", "Vehicles & Fleet", "070", "", "broad", ""),
    ("POLICE PURSUIT", "contains", "Vehicles & Fleet", "071", "", "broad", ""),
    ("RESPONSE VEHICLES", "contains", "Vehicles & Fleet", "071", "", "broad",
     "Police/emergency vehicles"),
    ("6 X 4 DIESEL POWERED CONVENTIONAL CAB", "contains", "Vehicles & Fleet", "070", "", "broad", ""),
    ("TOWING SERVICES", "contains", "Vehicles & Fleet", "928", "", "broad", ""),
    ("MAINTENANCE, REPAIR AND UPFIT FOR CITY-OWNED VEHICLES", "contains", "Vehicles & Fleet", "928", "", "broad", ""),

    # ===== Facilities Operations & Maintenance =====
    ("HVAC", "contains", "Facilities Operations & Maintenance", "031", "", "broad", ""),
    ("ELEVATORS, DUMBWAITERS", "contains", "Facilities Operations & Maintenance", "", "", "broad",
     "No elevator class in working set"),
    ("DOORS AND MOTORS:OVERHEAD", "contains", "Facilities Operations & Maintenance", "450", "", "broad", ""),
    ("VARIOUS DOORS AND MOTORS", "contains", "Facilities Operations & Maintenance", "450", "", "broad", ""),
    ("PLUMBING SUPPLIES", "contains", "Facilities Operations & Maintenance", "720", "", "broad", ""),
    ("ROOFING", "contains", "Facilities Operations & Maintenance", "", "", "broad",
     "No roofing class in working set"),
    ("FLAGPOLE", "contains", "Facilities Operations & Maintenance", "350", "", "broad", ""),
    ("LOCKSMITH SERVICES", "contains", "Facilities Operations & Maintenance", "450", "", "broad", ""),
    ("AIRPORT MAINTENANCE AND REPAIR", "contains", "Facilities Operations & Maintenance", "909", "", "broad", ""),
    ("AIRPORT FACILITY", "contains", "Facilities Operations & Maintenance", "909", "", "broad", ""),
    ("ELECTRICAL PRODUCTS AND SUPPLIES", "contains", "Facilities Operations & Maintenance", "285", "", "broad", ""),
    ("PURCHASE OF NEW AND REPLACEMENT ELECTRIC MOTORS", "contains", "Facilities Operations & Maintenance", "150", "", "broad", ""),
    ("PURCHASE OF PAINT", "contains", "Facilities Operations & Maintenance", "", "", "broad",
     "Building paint — no clean class"),
    ("AIR FILTERS", "contains", "Facilities Operations & Maintenance", "031", "", "broad",
     "HVAC filters"),
    ("FENCING, GUARDRAIL", "contains", "Facilities Operations & Maintenance", "", "", "broad", ""),
    ("TERRAZZO", "contains", "Facilities Operations & Maintenance", "", "", "broad", ""),
    ("SIGNAGE SYSTEMS FOR O'HARE", "contains", "Facilities Operations & Maintenance", "801", "", "broad",
     "Wayfinding/signage — facility hardware"),
    ("PLYMOVENT EXHAUST", "contains", "Facilities Operations & Maintenance", "031", "", "broad",
     "Vehicle exhaust extraction in fire stations"),

    # ===== Landscaping, Grounds & Irrigation =====
    ("LANDSCAPE SERVICES", "contains", "Landscaping, Grounds & Irrigation", "988", "", "broad", ""),
    ("COMPREHENSIVE LANDSCAPE", "contains", "Landscaping, Grounds & Irrigation", "988", "", "broad", ""),
    ("INTEGRATED PEST MANAGEMENT", "contains", "Landscaping, Grounds & Irrigation", "910", "", "broad", ""),
    ("SNOW PLOWING", "contains", "Landscaping, Grounds & Irrigation", "", "", "broad",
     "No snow-services class in working set"),
    ("AIRSIDE SNOW REMOVAL", "contains", "Landscaping, Grounds & Irrigation", "", "", "broad", ""),

    # ===== Public Safety, Uniforms & PPE =====
    ("FIRE EXTINGUISHER SERVICES", "contains", "Public Safety, Uniforms & PPE", "936", "", "broad", ""),
    ("FIRE SUPPRESSION", "contains", "Public Safety, Uniforms & PPE", "936", "", "broad", ""),
    ("RESCUE EQUIPMENT SUPPLIES", "contains", "Public Safety, Uniforms & PPE", "200", "", "broad", ""),
    ("SAFETY AND TECHNICAL RESCUE EQUIPMENT", "contains", "Public Safety, Uniforms & PPE", "200", "", "broad", ""),
    ("EMERGENCY MEDICAL EQUIPMENT", "contains", "Public Safety, Uniforms & PPE", "465", "", "broad",
     "EMS equipment — life-safety"),
    ("VARIOUS WORK AND BUSINESS UNIFORMS", "contains", "Public Safety, Uniforms & PPE", "200", "", "broad", ""),
    ("UNIFORMS AND COMMISSARY", "contains", "Public Safety, Uniforms & PPE", "200", "", "broad", ""),
    ("BREAD", "exact", "Public Safety, Uniforms & PPE", "375", "", "broad",
     "Police lockup bread (judgment call locked earlier)"),

    # ===== Office, Print & Marketing =====
    ("OFFICE SUPPLIES", "contains", "Office, Print & Marketing", "310", "", "broad", ""),
    ("Office supplies", "exact", "Office, Print & Marketing", "310", "", "broad", ""),
    ("LIBRARY MATERIALS", "contains", "Office, Print & Marketing", "", "", "broad",
     "Books/audiobooks — no library-materials class in working set"),
    ("Library materials", "exact", "Office, Print & Marketing", "", "", "broad", ""),
    ("Library Materials", "exact", "Office, Print & Marketing", "", "", "broad", ""),
    ("Books", "exact", "Office, Print & Marketing", "", "", "broad", ""),
    ("PROMOTIONAL WEARABLES", "contains", "Office, Print & Marketing", "037", "", "broad", ""),
    ("Promotional Wearables", "contains", "Office, Print & Marketing", "037", "", "broad", ""),
    ("DEMOLITION NOTICE DEFAULT PUBLICATION", "contains", "Office, Print & Marketing", "915", "", "broad",
     "Legal-notice publication"),

    # ===== Furniture & Furnishings =====
    ("VARIOUS FURNITURE", "contains", "Furniture & Furnishings", "425", "", "broad", ""),
    ("FURNITURE, ACCESSORIES", "contains", "Furniture & Furnishings", "425", "", "broad", ""),

    # ===== IT, Telecom & Audio/Visual =====
    ("MULTI-FUNCTION DEVICES", "contains", "IT, Telecom & Audio/Visual", "984", "", "broad",
     "MFDs/printers"),
    ("Printers and copiers", "contains", "IT, Telecom & Audio/Visual", "984", "", "broad", ""),
    ("TELECOMMUNICATION SERVICE", "contains", "IT, Telecom & Audio/Visual", "920", "", "broad", ""),
    ("Telecommunications, internet", "contains", "IT, Telecom & Audio/Visual", "920", "", "broad", ""),
    ("CELLULAR BASED WIRELESS", "contains", "IT, Telecom & Audio/Visual", "725", "", "broad", ""),
    ("WIRELESS COMMUNICATION EQUIPMENT", "contains", "IT, Telecom & Audio/Visual", "725", "", "broad", ""),
    ("INFORMATION TECHNOLOGY INFRASTRUCTURE", "contains", "IT, Telecom & Audio/Visual", "920", "", "broad", ""),
    ("AUDIO VISUAL SYSTEMS", "contains", "IT, Telecom & Audio/Visual", "939", "", "broad", ""),
    ("COMDIAL/VERTICAL TIER 1 DEALER TELEPHONE", "contains", "IT, Telecom & Audio/Visual", "725", "", "broad", ""),
    ("ISSOCCS", "contains", "IT, Telecom & Audio/Visual", "920", "", "broad",
     "Integrated safety/security/operations command system at airports"),

    # ===== Medical & Health Services =====
    ("CANINE PRIMARY CARE VETERINARY", "contains", "Animal Care & Veterinary", "875", "", "broad", ""),
    ("VETERINARY", "contains", "Animal Care & Veterinary", "875", "", "broad", ""),
    ("TEMPORARY MEDICAL PERSONNEL", "contains", "Medical & Health Services", "", "", "broad",
     "Locum/temp medical staffing"),
    ("COMPREHENSIVE MEDICAL/PHYSICAL EXAMS", "contains", "Medical & Health Services", "948", "", "broad", ""),
    ("INDUSTRIAL, MEDICAL & PROPANE GAS", "contains", "Medical & Health Services", "430", "", "broad", ""),

    # ===== Heavy Equipment & Machinery =====
    ("INDUSTRIAL EQUIPMENT, TOOLS, ATTACHMENTS", "contains", "Heavy Equipment & Machinery", "550", "", "broad", ""),
    ("HILTI TOOLS", "contains", "Heavy Equipment & Machinery", "550", "", "broad", ""),
    ("OUTDOOR POWER EQUIPMENT", "contains", "Heavy Equipment & Machinery", "550", "", "broad", ""),
    ("MATERIAL HANDLING, SHELVING", "contains", "Heavy Equipment & Machinery", "560", "", "broad", ""),

    # ===== Chemicals & Water Treatment =====
    ("ENVIRONMENTAL RESPONSE SERVICES", "contains", "Chemicals & Water Treatment", "175", "", "broad",
     "Hazmat response"),
    ("HAZARDOUS MATERIAL HANDLING", "contains", "Chemicals & Water Treatment", "175", "", "broad", ""),

    # ===== Professional & Administrative Services =====
    ("OUTSIDE COUNSEL", "contains", "Professional & Administrative Services", "", "", "review",
     "Legal services — no NIGP class in working set; consider broader class"),
    ("OUTSIDE COUNSEL", "exact", "Professional & Administrative Services", "", "", "review", ""),
    ("COURT REPORTER", "exact", "Professional & Administrative Services", "", "", "review",
     "Legal services"),
    ("COURT REPORTING", "contains", "Professional & Administrative Services", "", "", "review", ""),
    ("TRANSCRIPT", "exact", "Professional & Administrative Services", "", "", "review", ""),
    ("Transcript", "exact", "Professional & Administrative Services", "", "", "review", ""),
    ("SUBPOENA", "exact", "Professional & Administrative Services", "", "", "review", ""),
    ("EXPERT", "exact", "Professional & Administrative Services", "", "", "review",
     "Expert witness; legal context"),
    ("SENIOR INSTRUCTOR", "exact", "Professional & Administrative Services", "", "", "review", ""),
    ("TITLE SERVICES FOR PROPERTY", "contains", "Professional & Administrative Services", "", "", "review",
     "Real-estate title work"),
    ("FOREIGN LANGUAGE INTERPRETATION", "contains", "Professional & Administrative Services", "", "", "broad", ""),
    ("AVIATION PLANNING CONSULTING", "contains", "Professional & Administrative Services", "", "", "broad", ""),
    ("PROGRAM MANAGEMENT SERVICES FOR O'HARE", "contains", "Professional & Administrative Services", "958", "", "broad", ""),
    ("PROGRAM MANAGEMENT CONSULTANT", "contains", "Professional & Administrative Services", "958", "", "broad", ""),
    ("PROGRAM FINANCE SERVICES", "contains", "Professional & Administrative Services", "958", "", "broad", ""),
    ("PROFESSIONAL ENVIRONMENTAL ASSESSMENT", "contains", "Professional & Administrative Services", "", "", "broad", ""),
    ("QUALITY CONTROL AND QUALITY ASSURANCE CONSTRUCTION INSPECTION", "contains", "Professional & Administrative Services", "", "", "broad", ""),
    ("QUALITY ASSURANCE AND MATERIALS TESTING", "contains", "Professional & Administrative Services", "992", "", "broad", ""),
    ("LEAD TUNNEL DESIGN AND ENGINEERING", "contains", "Professional & Administrative Services", "", "", "broad", ""),
    ("Consultants or professional services", "exact", "Professional & Administrative Services", "", "", "broad", ""),
    ("Office rental", "exact", "Professional & Administrative Services", "", "", "review",
     "Real-estate lease — not in NIGP working set"),
    ("Utility services", "exact", "Professional & Administrative Services", "", "", "review",
     "Electric/water/gas utility — no class in working set"),
    ("Municipal Electricity Supply", "contains", "Professional & Administrative Services", "", "", "review",
     "Electricity supply contract"),
    ("STAFFING FOR MIGRANT", "contains", "Professional & Administrative Services", "", "", "review",
     "Staffing services"),
    ("UNARMED SECURITY GUARD", "contains", "Professional & Administrative Services", "964", "96400", "exact", ""),
    ("UNARMED SECURITY AND SCREENING", "contains", "Professional & Administrative Services", "964", "", "broad", ""),
    ("Security Services", "contains", "Professional & Administrative Services", "964", "", "broad", ""),
    ("RECORDS STORAGE", "contains", "Professional & Administrative Services", "962", "", "broad", ""),
    ("PROPERTY MANAGEMENT AND EVENT SUPPORT", "contains", "Professional & Administrative Services", "958", "", "broad", ""),
    ("Small Business Improvement Fund", "contains", "Professional & Administrative Services", "", "", "review",
     "Grant/financing program — not a commodity"),
    ("Residential Roadway Lighting Improvement", "contains", "Construction & Trades Services", "", "", "broad", ""),
    ("WARRANTY SERVICES FOR THE AIRPORT RESIDENTIAL SOUND INSULATION", "contains", "Facilities Operations & Maintenance", "557", "", "broad", ""),
]


def main() -> None:
    df = pd.DataFrame(RULES, columns=[
        "pattern", "match_type", "business_category",
        "nigp_class_3digit", "nigp_item_5digit", "nigp_match_level", "notes",
    ])
    # Sort: exact > starts_with > contains, then alphabetical within
    order = {"exact": 0, "starts_with": 1, "contains": 2}
    df["__order"] = df["match_type"].map(order)
    df = df.sort_values(["__order", "business_category", "pattern"]).drop(columns="__order").reset_index(drop=True)

    out = OUT_DIR / "keyword_rules_DRAFT_JHK3.csv"
    df.to_csv(out, index=False)

    print(f"Drafted {len(df)} keyword rules -> {out}")
    print()
    print("=== Rule count by Business Category ===")
    print(df.groupby("business_category").size().sort_values(ascending=False).to_string())
    print()
    print("=== Rule count by match_level ===")
    print(df.groupby("nigp_match_level").size().to_string())
    print()
    print("=== Rule count by match_type ===")
    print(df.groupby("match_type").size().to_string())


if __name__ == "__main__":
    main()
