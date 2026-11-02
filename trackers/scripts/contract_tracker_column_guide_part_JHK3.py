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
    # ---------------- other tabs (one line each)
    ("Alert Controls", "Blue cells", "Settings", IN_, "Every list, lead time, rule, threshold and colour the alerts use.", "Change blue cells and dropdowns only. Read the How To Use tab first; rules are checked top to bottom.", "Competitive rebid = 540 days", "-"),
    ("Dashboard, Summary, Task Order Dashboard, Data Check", "All", "All cells", FX, "Pictures and totals of the trackers for you and leadership.", "Nothing to type. They update when the trackers change.", "-", "-"),
]

add_column_guide(
    wb, "Column Guide",
    "Every column on the Portfolio Tracker and Task Order Tracker: who fills it in, what it means, and an example. "
    "Examples are shown here, not as sample rows on the trackers, so they never show up in the dashboards or totals.",
    CG_TYPES, CG_ROWS, index=1)
