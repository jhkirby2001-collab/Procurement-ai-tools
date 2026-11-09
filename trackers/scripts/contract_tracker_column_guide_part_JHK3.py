# Column Guide tab for the Contract Tracker. Exec'd inside build_contract_tracker_JHK3.py (uses its globals).
# Keep this list in step with the tracker columns: if a column is added, moved or renamed, update it here.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from column_guide_JHK3 import add_column_guide

ORIG, IN_, FX = "From the original list", "You fill in", "Formula - don't type"
CG_TYPES = {
    ORIG: ("1F3864", "FFFFFF", "Came with the contract list. Correct it if it's wrong; nothing recalculates it."),
    IN_: ("2E75B6", "FFFFFF", "You type it or pick it from the dropdown. Keep these up to date - every alert depends on them."),
    FX: ("595959", "FFFFFF", "Calculated for you. Don't type in it: typing over a formula breaks that row. Change the result by changing the inputs or the settings."),
}
PT_, TO_ = "Portfolio Tracker", "Task Order Tracker"
AC_ = "Alert Controls"
CG_ROWS = [
    # ---------------- Portfolio Tracker
    (PT_, "A", "PO Number", ORIG, "The purchase order number of the contract. It is the contract's ID across the workbook.", "Leave as received. Task orders link to the contract by this number.", "232326", ""),
    (PT_, "B", "PO Description", ORIG, "What the contract buys or what service it covers.", "Leave as received.", "Liquid cationic polyelectrolyte", ""),
    (PT_, "C", "Start Date", ORIG, "The day the contract started.", "A real date (mm/dd/yyyy). Used to work out how much of the term has passed.", "07/07/2023", ""),
    (PT_, "D", "End Date", ORIG, "The day the contract ends. Drives every expiration alert.", "A real date (mm/dd/yyyy). Update it when an extension is executed.", "07/07/2028", ""),
    (PT_, "E", "Vendor Name", ORIG, "The company holding the contract.", "Leave as received.", "Polydyne, Inc", ""),
    (PT_, "F", "Bureau", ORIG, "The bureau that uses the contract.", "Leave as received.", "BWS", ""),
    (PT_, "G", "Dept Code", ORIG, "The department code from the original list.", "Leave as received.", "88", ""),
    (PT_, "H", "Analyst", ORIG, "The analyst listed on the original list.", "Leave as received. The current owner is in Proposed Owner (column Q).", "Analyst name", ""),
    (PT_, "I-N", "Facility, Water/Sewer, Pipes, Valves, Meters, Chemical (hidden)", ORIG, "Lane markers from the original list. An x shows which lane the contract belongs to.", "An x in one column. These columns are hidden; Lane (P) shows the result.", "x", ""),
    (PT_, "O", "Spec Reviewer", ORIG, "Who reviews the specification for this contract.", "Leave as received.", "Reviewer name", ""),
    (PT_, "P", "Lane (from x columns)", ORIG, "The commodity group the contract belongs to.", "One of the lane names. Used for the lane breakdowns on the Summary and Dashboard.", "Chemical", f"{AC_} section 8 (lane list for the Summary)"),
    (PT_, "Q", "Proposed Owner (post-transition)", ORIG, "The staff member who owns the contract after the transition.", "A person's name, typed the same way every time so the owner counts add up.", "Staff member name", f"{AC_} section 8 (owner list for the Summary)"),
    (PT_, "R", "Renewal Path", IN_, "How the contract will continue when it ends. Sets how early renewal work must start.", "Pick from the list: Extension available, Competitive rebid, Sole source / NCRB, Let expire.", "Competitive rebid", f"{AC_} section 2"),
    (PT_, "S", "Action Status", IN_, "Where the renewal work stands right now.", "Pick from the list, e.g. Not started, Req drafted, Submitted to DPS, Replaced / executed.", "Submitted to DPS", f"{AC_} section 3"),
    (PT_, "T", "Next Action", IN_, "The one next step on this contract.", "A few words describing the step.", "Send usage estimate to bureau", ""),
    (PT_, "U", "Next Action Due Date", IN_, "When the next step is due.", "A date (mm/dd/yyyy).", "11/15/2026", ""),
    (PT_, "V", "Next Action Done?", IN_, "Whether the next step is finished.", "Pick Yes when done, then type the next step and its date (or clear the date).", "No", f"{AC_} section 1"),
    (PT_, "W", "Total Contract Value ($)", IN_, "The most that can be spent on the contract. For a master contract, this is the ceiling task orders draw from.", "A dollar amount, numbers only.", "500,000", ""),
    (PT_, "X", "Expended to Date ($)", IN_, "How much has been spent on the contract so far.", "A dollar amount from the financial system. Update monthly.", "125,000", ""),
    (PT_, "Y", "Remaining Balance ($)", FX, "Money left on the contract.", "Total Contract Value minus Expended to Date.", "375,000", ""),
    (PT_, "Z", "% Expended", FX, "Share of the contract value already spent.", "Expended to Date divided by Total Contract Value.", "25%", ""),
    (PT_, "AA", "% of Term Used", FX, "Share of the contract's time that has passed.", "Days since Start Date divided by total contract days. Shows 'Check start date' if the start date isn't a real date.", "60%", ""),
    (PT_, "AB", "Spending Ahead of Time (pts)", FX, "Whether money is being spent faster than time is passing.", "% Expended minus % of Term Used. A positive number means spending is running ahead.", "-35%", ""),
    (PT_, "AC", "Funds Flag", FX, "Money alert for the contract.", "RED at 75% spent or 20+ points ahead of time; YELLOW at 50% or 10+ points; otherwise GREEN. Blank until a value is entered.", "GREEN", f"{AC_} section 5"),
    (PT_, "AD", "Days Until Expiration", FX, "How many days until the contract ends.", "End Date minus today. Negative means it has already ended.", "615", ""),
    (PT_, "AE", "Lead Time (days)", FX, "How many days before the end date renewal work should start.", "Looked up from the Renewal Path. 540 when no path is picked yet.", "540", f"{AC_} sections 1 and 2"),
    (PT_, "AF", "Start-By Date", FX, "The last day to start renewal work and still finish before the contract ends.", "End Date minus Lead Time.", "01/14/2027", ""),
    (PT_, "AG", "Days Until Start-By", FX, "How many days until the Start-By Date.", "Start-By Date minus today. Negative means the start date has passed.", "75", ""),
    (PT_, "AH", "Action Category", FX, "Groups the Action Status into Normal, In Process or Closed.", "Looked up from the Action Status. Blank status counts as Normal.", "In Process", f"{AC_} section 3"),
    (PT_, "AI", "Expiration Status", FX, "The main time alert for the contract.", "First rule that fits: CLOSED, LAPSED (past end date), IN PROCESS, RED (180 days or less, or Start-By passed), YELLOW (365 days or less, or Start-By within 90), GREEN.", "YELLOW", f"{AC_} section 4"),
    (PT_, "AJ", "Days Until Next Action", FX, "How many days until the next step is due.", "Next Action Due Date minus today. Blank when there is no date or the step is done.", "14", ""),
    (PT_, "AK", "Next Action Flag", FX, "Alert for the next step.", "OVERDUE when past due, DUE SOON within 14 days, UPCOMING within 30 days.", "DUE SOON", f"{AC_} section 6"),
    (PT_, "AL", "Priority (1 = most urgent)", FX, "How urgent this contract is overall. Sort by this column and work top down.", "The most urgent of the three flags (Expiration, Funds, Next Action), using each status's severity number.", "3", f"{AC_} sections 4-6 (Severity)"),
    (PT_, "AM", "Notes / Next Action", IN_, "Short note on what needs to happen.", "Free text.", "Confirm owner after transition", ""),
    (PT_, "AN", "Status /Notes", IN_, "Running history of what has happened on the contract.", "Free text. Add the date at the start of each new note.", "10/28/26 - TE submitted to DPS, req 748195", ""),
    (PT_, "AO", "REQ No.", IN_, "Requisition number for the renewal, extension or new contract.", "The req number.", "748195", ""),
    (PT_, "AP", "Additional Req No.", IN_, "A second requisition number, if there is one.", "The req number.", "751226", ""),
    (PT_, "AQ", "Vendor Contact", IN_, "Who to contact at the vendor.", "Name, title, phone and email.", "Name | Sales Rep | 555-0100", ""),
    (PT_, "AR", "Contract Type", IN_, "Whether the contract is standard or one that task orders / job order releases are issued against.", "Pick from the list: Standard, Master Service, Construction (JOC).", "Master Service", f"{AC_} section 10"),
    (PT_, "AS", "Owning Dept Code", IN_, "The PO Dept Code of the department that owns the contract.", "A department code, numbers only. Loaded from the shared-contracts file; contracts only on the original list use their Dept Code.", "38", f"{AC_} section 15"),
    (PT_, "AT", "Ownership", FX, "Whether your department owns the contract or shares another department's.", "Owned when Owning Dept Code equals your department code (88); Shared for any other code; Not set when the code is blank.", "Shared", f"{AC_} section 15"),
    (PT_, "AU", "Spec Number", IN_, "The specification number for the contract.", "The spec number. Loaded from the shared-contracts file.", "1258382", ""),
    (PT_, "AV", "PO Contract Type", IN_, "The purchasing category of the PO (commodities, work services, professional services...).", "Type it the same way each time. Spellings from the shared-contracts file were made consistent.", "COMMODITIES", ""),
    (PT_, "AW", "BWS File Notes", IN_, "Notes carried over from the BWS shared-contracts file.", "Free text. Kept separate from Status /Notes so you can see what came from that file.", "Req 650032 submitted to DPS; compliance approved", ""),
    (PT_, "AX", "On DPS 365 Report?", FX, "Whether this contract is on the monthly DPS 365 expiring-contracts report.", "Yes / No by PO. 'Missing - check with DPS' when one of your Owned contracts ends within a year (Alert Controls 16), isn't closed, and isn't on the report. Shared contracts appear on the owning department's report, so they are never flagged.", "Yes", f"{AC_} section 16"),
    (PT_, "AY", "DPS 365 Comments", FX, "The comment the procurement manager entered on the DPS 365 report for this PO.", "Pulled from the DPS 365 Report tab (Additional Comments). Your own notes stay in Status /Notes.", "RX 744110 submitted 8/7/26", ""),
    (PT_, "AZ", "DPS Procurement Mgr", FX, "Procurement manager initials on the DPS 365 report.", "Pulled from the DPS 365 Report tab.", "JK", ""),
    (PT_, "BA", "Burn Index", FX, "How fast money is being spent compared with time. 1.00 = on pace; above 1 = money runs out before the contract ends.", "% Expended divided by % of Term Used. Blank until 10% of the term has passed and money has been spent. Amber at 1.10, red at 1.25.", "1.32", f"{AC_} section 17"),
    (PT_, "BB", "Projected Exhaustion Date", FX, "When the money runs out at the current spending pace.", "Start Date + days elapsed / % Expended. Blank when Burn Index is blank.", "03/15/2027", ""),
    (PT_, "BC", "Funds Run Out Before End?", FX, "Early warning that a vendor limit increase is needed.", "'Yes - N days early' when the Projected Exhaustion Date is before the End Date; otherwise No.", "Yes - 142 days early", ""),
    (PT_, "BD", "Supply Coverage", IN_, "How DWM keeps getting materials while a contract is lapsed or about to lapse.", "Pick from the list (Exhibit B, Alternate PO / contract, Extension pending, Stock on hand, Unconfirmed, No way to buy, Not needed). Yellow = suggested from your notes; confirm or change.", "Exhibit B", f"{AC_} section 18"),
    (PT_, "BE", "Coverage PO / Note", IN_, "Proof of the coverage.", "The PO number, approval or note.", "PO 246545", ""),
    (PT_, "BF", "Coverage Through", IN_, "When the coverage ends, if it does (approval expiry, stock run-out).", "A date (mm/dd/yyyy). The row turns amber 30 days before; after it passes, supply counts as At risk.", "01/31/2027", f"{AC_} section 19"),
    (PT_, "BG", "Supply Status", FX, "Can DWM still buy what this contract covers?", "Your Supply Coverage choice; otherwise Unknown - confirm (lapsed), Not yet planned (ending within 90 days) or Active contract.", "Unknown - confirm", f"{AC_} section 19"),
    (PT_, "BH", "Supply Risk", FX, "Supply Status in three words for the Dashboard.", "Covered / At risk / Not covered (from what the coverage choice counts as), Active or Closed.", "At risk", f"{AC_} section 18"),
    (PT_, "BI", "Action Needed", FX, "Whether the contract needs YOUR action now: Act now (dark red) or Plan (amber). Only this cell is coloured unless the whole-row option is on.", "Act now: next action overdue, supply not covered, lapsed with supply at risk, or Priority 1-2 without a renewal in process. Plan: Priority 3, ending soon with no coverage, or coverage ending soon. Blank when the renewal is in process and supply is covered.", "Act now", f"{AC_} section 19"),
    # ---------------- Task Order Tracker
    (TO_, "A", "Release ID", IN_, "Your team's ID for the release.", "One format for everyone, e.g. R-001.", "R-001", ""),
    (TO_, "B", "Master PO", IN_, "The master contract this release is issued against.", "Pick the PO from the dropdown (it lists the Portfolio Tracker POs).", "271670", ""),
    (TO_, "C", "Master Description", FX, "What the master contract covers.", "Looked up from the Portfolio Tracker by Master PO. Says 'PO not on Portfolio Tracker' if the PO isn't found.", "Electrical power systems maintenance", ""),
    (TO_, "D", "Contract Type", FX, "The master contract's type.", "Looked up from the Portfolio Tracker. 'Not set' means pick a Contract Type there.", "Master Service", f"{AC_} section 10"),
    (TO_, "E", "Release / Task Order No.", IN_, "The number printed on the task order or job order release.", "The document number.", "TO-2026-014", ""),
    (TO_, "F", "Release Type", IN_, "What kind of release this is.", "Pick from the list: Task Order, Job Order Release, Release Change Order.", "Task Order", f"{AC_} section 11"),
    (TO_, "G", "Description / Scope", IN_, "What work the release covers.", "A short description.", "Replace switchgear at pumping station", ""),
    (TO_, "H", "Requesting Bureau", IN_, "The bureau that asked for the work.", "Bureau name or code.", "BWS", ""),
    (TO_, "I", "Issue Date", IN_, "The day the release was issued. Decides the order releases count against the ceiling.", "A date (mm/dd/yyyy).", "01/10/2026", ""),
    (TO_, "J", "Start Date", IN_, "The day the work starts.", "A date (mm/dd/yyyy).", "02/01/2026", ""),
    (TO_, "K", "End Date", IN_, "The day the release ends.", "A date (mm/dd/yyyy). Used for the Overdue and Ending Soon flags.", "01/01/2027", ""),
    (TO_, "L", "Release Amount ($)", IN_, "The dollar amount of the release.", "Numbers only. For a change order, enter only the change (negative for a decrease).", "300,000", ""),
    (TO_, "M", "Paid to Date ($)", IN_, "How much has been paid on this release.", "Numbers only, from invoices. Optional.", "50,000", ""),
    (TO_, "N", "Status", IN_, "Where the release stands. Decides whether its amount counts as committed.", "Pick from the list: Draft, Pending approval, Issued, In progress, Complete, Closed, Cancelled.", "Issued", f"{AC_} section 12"),
    (TO_, "O", "M/WBE Goal %", IN_, "The M/WBE participation goal for the release.", "A percent. Optional.", "26%", ""),
    (TO_, "P", "M/WBE Actual %", IN_, "The M/WBE participation achieved.", "A percent. Optional.", "22%", ""),
    (TO_, "Q", "Counts as Committed?", FX, "Whether this release's amount is subtracted from the master's ceiling.", "Looked up from the Status. Issued, In progress, Complete and Closed count by default.", "Yes", f"{AC_} section 12"),
    (TO_, "R", "Still Open?", FX, "Whether the release can still get date alerts.", "Looked up from the Status.", "Yes", f"{AC_} section 12"),
    (TO_, "S", "Remaining on Release ($)", FX, "Money left on this release.", "Release Amount minus Paid to Date.", "250,000", ""),
    (TO_, "T", "Master Ceiling ($)", FX, "The most that can be committed on the master contract.", "The master's Total Contract Value on the Portfolio Tracker.", "500,000", ""),
    (TO_, "U", "Master End Date", FX, "When the master contract ends.", "The master's End Date on the Portfolio Tracker.", "10/10/2029", ""),
    (TO_, "V", "Master Total Committed ($)", FX, "Everything committed on the master so far, across all its releases.", "Sum of Release Amount for releases on the same master that count as committed.", "380,000", ""),
    (TO_, "W", "Master Available ($)", FX, "Money left on the master for new releases. The number to watch.", "Master Ceiling minus Master Total Committed. Shown in red when negative.", "120,000", ""),
    (TO_, "X", "Committed Through This Release ($)", FX, "Running total on the master up to and including this release.", "Committed releases on the same master issued on or before this release's Issue Date. Used for the Exceeds Ceiling flag.", "300,000", ""),
    (TO_, "Y", "Days to Release End", FX, "How many days until the release ends.", "End Date minus today, for open releases only.", "61", ""),
    (TO_, "Z", "Release Flag", FX, "One alert for the release.", "First rule that fits: EXCEEDS CEILING, NOT A MASTER, OUTLASTS MASTER, OVERPAID, OVERDUE, ENDING SOON. Blank when nothing is wrong.", "ENDING SOON", f"{AC_} section 14"),
    (TO_, "AA", "M/WBE Gap", FX, "How far actual participation is from the goal.", "M/WBE Actual % minus Goal %. Red when below goal.", "-4%", ""),
    (TO_, "AB", "Notes", IN_, "Anything else about the release.", "Free text. For a change order, note the original release number.", "Change order to TO-2026-014", ""),
    # ---------------- DPS 365 Report
    ("DPS 365 Report", "A", "Current End Date", ORIG, "End date on the DPS report.", "As sent by DPS. Never copied to your tracker.", "04/30/2027", ""),
    ("DPS 365 Report", "B", "Original Start Date", ORIG, "Start date on the DPS report.", "As sent by DPS.", "11/01/2018", ""),
    ("DPS 365 Report", "C", "Spec Number", ORIG, "Specification number.", "As sent by DPS.", "338124", ""),
    ("DPS 365 Report", "D", "PO Number", ORIG, "The contract's PO. Links the row to your tracker.", "As sent by DPS.", "84272", ""),
    ("DPS 365 Report", "E", "PO Description", ORIG, "What the contract covers.", "As sent by DPS.", "Reprocessable materials", ""),
    ("DPS 365 Report", "F", "No Longer Needed (X)", IN_, "Mark contracts the department no longer needs.", "Type X.", "X", ""),
    ("DPS 365 Report", "G", "Provide Req # for extension request (Req #)", IN_, "Requisition for the extension.", "The req number from FMPS.", "714339", ""),
    ("DPS 365 Report", "H", "Is Extension request within the term of the contract?", IN_, "Whether the extension fits within the contract's existing term.", "Pick Yes or No.", "Yes", ""),
    ("DPS 365 Report", "I", "Req # for new request that is replacing expiring contract", IN_, "Requisition for the replacement contract.", "The req number from FMPS.", "744110", ""),
    ("DPS 365 Report", "J", "Additional Comments", IN_, "The monthly status comment for DPS. Shows on your tracker in DPS 365 Comments.", "Update monthly. For your rows, start from the Suggested Comment.", "RX 744110 submitted 8/7/26", ""),
    ("DPS 365 Report", "K", "Procurement Manager Initials", ORIG, "Who manages the contract.", "As sent by DPS. Sets the Responsibility group for rows not on your tracker.", "JK", f"{AC_} section 16"),
    ("DPS 365 Report", "L", "On My Tracker?", FX, "Whether the PO is on your Portfolio Tracker.", "Yes / No by PO.", "Yes", ""),
    ("DPS 365 Report", "M", "Responsibility", FX, "Whose contract it is.", "My portfolio when it's on your tracker; otherwise the group for the manager initials (Visibility only by default).", "My portfolio", f"{AC_} section 16"),
    ("DPS 365 Report", "N", "Tracker End Date", FX, "The end date on your tracker.", "Looked up by PO. Shown for comparison only.", "10/31/2026", ""),
    ("DPS 365 Report", "O", "End Dates Differ?", FX, "Whether DPS and your tracker disagree on the end date.", "'Differs' when the two dates aren't the same.", "Differs", ""),
    ("DPS 365 Report", "P", "Tracker Expiration Status", FX, "Your tracker's time alert for the contract.", "Looked up from the Portfolio Tracker.", "LAPSED", f"{AC_} section 4"),
    ("DPS 365 Report", "Q", "Renewal Path", FX, "Your tracker's renewal path.", "Looked up from the Portfolio Tracker.", "Extension available", ""),
    ("DPS 365 Report", "R", "Action Status", FX, "Your tracker's action status.", "Looked up from the Portfolio Tracker.", "Submitted to DPS", ""),
    ("DPS 365 Report", "S", "Next Action", FX, "Your tracker's next step.", "Looked up from the Portfolio Tracker.", "DPS must finish process", ""),
    ("DPS 365 Report", "T", "Next Action Due", FX, "When that step is due.", "Looked up from the Portfolio Tracker.", "11/15/2026", ""),
    ("DPS 365 Report", "U", "Tracker REQ No.", FX, "The req number on your tracker.", "Looked up from the Portfolio Tracker.", "748185", ""),
    ("DPS 365 Report", "V", "Suggested Comment", FX, "A draft monthly comment built from your tracker.", "Today's date, action status, req number and next step. Copy it into Additional Comments and edit as needed.", "11/3/26 - Submitted to DPS; Req 748185; Next: DPS must finish process by 11/15/26", ""),
    # ---------------- other tabs (one line each)
    ("Alert Controls", "Blue cells", "Settings", IN_, "Every list, lead time, rule, threshold and colour the alerts use.", "Change blue cells and dropdowns only. Read the How To Use tab first; rules are checked top to bottom.", "Competitive rebid = 540 days", "-"),
    ("Dashboard, Summary, Task Order Dashboard, Data Check", "All", "All cells", FX, "Pictures and totals of the trackers for you and leadership.", "Nothing to type. They update when the trackers change.", "-", "-"),
]

# Final layout (contract_tracker_layout_JHK3.FINAL_ORDER): drop removed columns, rename, add the new ones, and show the
# Portfolio Tracker letters exactly as they appear after finalize_layout_JHK3.py.
from contract_tracker_layout_JHK3 import FINAL_ORDER
_n = lambda h: " ".join(str(h).replace("(auto)", "").split())
_final = {_n(h): i for i, h in enumerate(FINAL_ORDER)}
_GONE = {"Next Action", "Next Action Due Date", "Next Action Done?", "Days Until Next Action", "Next Action Flag",
         "Notes / Next Action", "Status /Notes", "BWS File Notes", "DPS 365 Comments", "Coverage PO / Note", "Coverage Through"}
_FIX = {
    "Action Status": ("Renewal Stage", "Where the renewal work stands right now (formerly Action Status). Drives IN PROCESS / CLOSED and the lapsed-movement KPI.",
                      "Pick from the list, e.g. Not started, Req drafted, Submitted to DPS, Replaced / executed."),
    "Action Category": (None, "Groups the Renewal Stage into Normal, In Process or Closed. Helper column (collapsed).",
                        "Looked up from the Renewal Stage. Blank counts as Normal."),
    "Priority (1 = most urgent)": (None, "How urgent this contract is overall. Sort by this column and work top down.",
                                   "The more urgent of the Expiration Status and Funds Flag, using each status's severity number."),
    "Supply Status": (None, "Can DWM still buy what this contract covers?",
                      "Your Supply Coverage choice; otherwise Unknown - confirm (lapsed), Not yet planned (ending within 90 days) or Active contract."),
    "Supply Coverage": (None, "How DWM keeps getting materials while a contract is lapsed or about to lapse.",
                        "Pick from the list. Put the proof (PO number, approval) in Status & Notes. Yellow = suggested from your notes; confirm or change."),
    "Action Needed": (None, "Whether the contract needs YOUR action now: Act now (dark red) or Plan (amber). Only this cell is coloured unless the whole-row option is on.",
                      "Act now: supply not covered, lapsed with supply at risk, or Priority 1-2 without a renewal in process. Plan: Priority 3 or ending soon with no coverage. Blank when the renewal is in process and supply is covered."),
}
_HELPER = {"Lead Time (days)", "Days Until Start-By", "Action Category", "Spending Ahead of Time (pts)"}
_pt_rows = []
for _row in CG_ROWS:
    if _row[0] != PT_ or _row[2] in _GONE:
        continue
    _t, _c, _name, _ty, _mean, _how, _ex, _set = _row
    if _name in _FIX:
        _nn, _mm, _hh = _FIX[_name]
        _name = _nn or _name; _mean = _mm; _how = _hh
        if _name == "Renewal Stage": _ex = "Submitted to DPS"
    if _name in _HELPER and "collapsed" not in _mean:
        _mean += " Helper column (click + above the column letters to show)."
    if _c == "I-N":
        _c = "J-O"
    else:
        _c = L(_final[_n(_name)] + 1)
    _pt_rows.append((_t, _c, _name, _ty, _mean, _how, _ex, _set))
_pt_rows.append((PT_, L(_final["Reviewed"] + 1), "Reviewed", IN_, "Your own check-off column for rows you have reviewed.", "Type y (or a date). No formula uses it.", "y", ""))
_pt_rows.append((PT_, L(_final["Status & Notes"] + 1), "Status & Notes", IN_,
                 "The one notes column: status, next steps, BWS shared-list notes and DPS 365 comments for the contract.",
                 "Add new notes at the top with a date. Lines are labelled Status / Notes / Next / BWS / DPS 365 / Coverage. The old separate note columns are on the Archive - Removed Columns tab.",
                 "11/9/26 - Req 748185 with DPS; follow up 11/15", ""))
_pt_rows.sort(key=lambda r: CI(r[1].split("-")[0]))
_first = next(i for i, r in enumerate(CG_ROWS) if r[0] == PT_)
CG_ROWS = CG_ROWS[:_first] + _pt_rows + [r for r in CG_ROWS[_first:] if r[0] != PT_]
_DPSFIX = {"R": ("Renewal Stage", "Your tracker's renewal stage.", "Looked up from the Portfolio Tracker."),
           "S": ("Status & Notes (tracker)", "Your tracker's Status & Notes.", "Looked up from the Portfolio Tracker."),
           "T": ("Supply Status", "Can DWM still buy (from your tracker).", "Looked up from the Portfolio Tracker."),
           "V": ("Suggested Comment", "A draft monthly comment built from your tracker.", "Today's date, renewal stage, req number and supply status. Copy it into Additional Comments and edit as needed.")}
_DPSEX = {"R": "Submitted to DPS", "S": "Status: Req 748185 with DPS...", "T": "Exhibit B",
          "V": "11/9/26 - Submitted to DPS; Req 748185; Supply: Exhibit B"}
CG_ROWS = [((r[0], r[1], _DPSFIX[r[1]][0], r[3], _DPSFIX[r[1]][1], _DPSFIX[r[1]][2], _DPSEX[r[1]], r[7]))
           if (r[0] == "DPS 365 Report" and r[1] in _DPSFIX) else r for r in CG_ROWS]
add_column_guide(
    wb, "Column Guide",
    "Every column on the Portfolio Tracker, DPS 365 Report and Task Order Tracker: who fills it in, what it means, and an example. "
    "Examples are shown here, not as sample rows on the trackers, so they never show up in the dashboards or totals.",
    CG_TYPES, CG_ROWS, index=1)
