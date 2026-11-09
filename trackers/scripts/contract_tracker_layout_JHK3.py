"""Final Portfolio Tracker column layout (James asked for a simpler tracker, 2026-11-09).

The build writes the columns in its own order; finalize_layout_JHK3.py then uses LibreOffice to move every column into
FINAL_ORDER (references on every tab follow the move, as in Excel), deletes REMOVED and folds HELPERS into a collapsed
+/- group. Header text must match the build's row-1 headers exactly.
"""
FINAL_ORDER = [
    # contract
    "Reviewed", "PO Number", "PO Description", "Start Date", "End Date", "Vendor Name", "Bureau", "Dept Code", "Analyst",
    "Facility", "Water/Sewer", "Pipes", "Valves", "Meters", "Chemical",          # lane markers (hidden)
    # the one notes column
    "Status & Notes",
    # ownership
    "Spec Reviewer", "Lane (from x columns)", "Proposed Owner (post-transition)",
    # renewal
    "Renewal Path", "Renewal Stage", "Days Until Expiration (auto)", "Start-By Date (auto)", "Expiration Status (auto)",
    "Priority (1 = most urgent) (auto)", "Action Needed (auto)",
    # supply
    "Supply Coverage", "Supply Status (auto)", "Supply Risk (auto)",
    # money
    "Total Contract Value ($)", "Expended to Date ($)", "Remaining Balance ($) (auto)", "% Expended (auto)",
    "% of Term Used (auto)", "Burn Index (auto)", "Projected Exhaustion Date (auto)", "Funds Run Out Before End? (auto)",
    "Funds Flag (auto)",
    # reference
    "REQ No.", "Additional Req No.", "Contract Type", "Owning Dept Code", "Ownership (auto)", "Spec Number",
    "PO Contract Type", "Vendor\nContact", "On DPS 365 Report? (auto)", "DPS Procurement Mgr (auto)",
    # helpers (collapsed group)
    "Lead Time (days) (auto)", "Days Until Start-By (auto)", "Action Category (auto)", "Spending Ahead of Time (pts) (auto)",
]
HELPERS = ["Lead Time (days) (auto)", "Days Until Start-By (auto)", "Action Category (auto)",
           "Spending Ahead of Time (pts) (auto)"]
REMOVED = ["Next Action", "Next Action Due Date", "Next Action Done?", "Days Until Next Action (auto)",
           "Next Action Flag (auto)", "Notes / Next Action", "Status /Notes", "BWS File Notes",
           "DPS 365 Comments (auto)", "Coverage PO / Note", "Coverage Through"]


def final_letter(header):
    from openpyxl.utils import get_column_letter
    return get_column_letter(FINAL_ORDER.index(header) + 1)
