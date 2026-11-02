# Column Guide tab for the Project Tracker. Exec'd inside build_project_tracker_JHK3.py (uses its globals).
# Keep this list in step with the tracker columns: if a column is added, moved or renamed, update it here.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from column_guide_JHK3 import add_column_guide

IN_, FX = "You fill in", "Formula - don't type"
PG_TYPES = {
    IN_: ("1F3864", "FFFFFF", "You type it or pick it from the dropdown (navy header on the tracker)."),
    FX: ("595959", "FFFFFF", "Calculated for you (grey header marked 'auto'). Don't type in it: typing over a formula breaks that row."),
}
PJ, TK, IS, LI, PM = "Project Tracker", "Task Tracker", "Issue Tracker", "Lists", "Priority Matrix"
PG_ROWS = [
    # ---------------- Project Tracker
    (PJ, "A", "Project ID", IN_, "The project's ID. Tasks and issues link to the project by this ID.", "One format for everyone, e.g. P-001. Duplicates turn red.", "P-001", ""),
    (PJ, "B", "Portfolio", IN_, "The group of work this project belongs to.", "Pick from the list.", "Process Improvement", f"{LI} tab (Portfolios)"),
    (PJ, "C", "Project Name", IN_, "Short name of the project.", "A few words.", "Standardize document filing", ""),
    (PJ, "D", "Objective", IN_, "What the project is meant to achieve.", "One sentence, measurable if possible.", "Cut time to find a file to under 2 minutes", ""),
    (PJ, "E", "Key Deliverables", IN_, "The main things the project will produce.", "Short list.", "Folder structure; naming guide", ""),
    (PJ, "F", "Milestones", IN_, "Key checkpoints and dates.", "Short list with dates.", "Draft 11/15; go-live 12/31", ""),
    (PJ, "G", "Owner", IN_, "The person responsible for the project.", "Pick from the list.", "Staff member name", f"{LI} tab (Staff)"),
    (PJ, "H", "Stakeholders", IN_, "People or groups affected or involved.", "Names or groups.", "Bureau leads, IT", ""),
    (PJ, "I", "Start Date", IN_, "When the project starts.", "A date (mm/dd/yyyy).", "11/01/2026", ""),
    (PJ, "J", "Due Date", IN_, "When the project must be finished.", "A date (mm/dd/yyyy).", "12/31/2026", ""),
    (PJ, "K", "Status", IN_, "Where the project stands.", "Pick from the list: Not Started, In Progress, On Hold, Completed, Cancelled.", "In Progress", f"{LI} tab"),
    (PJ, "L", "Days Left", FX, "Days until the project is due.", "Due Date minus today. Blank once Completed or Cancelled; negative means past due.", "60", ""),
    (PJ, "M", "Total Tasks", FX, "How many tasks are linked to this project.", "Counts Task Tracker rows with this Project ID.", "8", ""),
    (PJ, "N", "Open Tasks", FX, "Tasks not yet finished.", "Total Tasks minus Completed and Cancelled.", "5", ""),
    (PJ, "O", "Tasks Completed", FX, "Tasks marked Completed.", "Counts this project's tasks with status Completed.", "3", ""),
    (PJ, "P", "% Complete", FX, "How far along the project is.", "Completed tasks divided by (open + completed) tasks. Cancelled tasks don't count.", "38%", ""),
    (PJ, "Q", "Overdue Tasks", FX, "Tasks past their due date.", "Counts this project's tasks flagged Overdue.", "1", ""),
    (PJ, "R", "Open Issues", FX, "Issues not yet resolved.", "Counts this project's issues that aren't Resolved or Closed.", "1", ""),
    (PJ, "S", "Health", FX, "Overall health of the project.", "Red = past due, an overdue task, or an open critical issue. Amber = due within 14 days or any open issue. Green = on track.", "Amber", f"{LI} tab (Settings)"),
    (PJ, "T", "Potential Risks / Challenges", IN_, "What could get in the way.", "Short list.", "IT access may be delayed", ""),
    (PJ, "U", "Proposed Solutions", IN_, "How you plan to handle the risks.", "Short list.", "Request access this week", ""),
    (PJ, "V", "Next Steps", IN_, "What happens next on the project.", "Short list.", "Review draft with bureau leads", ""),
    (PJ, "W", "Progress Update", IN_, "Latest update for your leader.", "One or two sentences, dated.", "10/30 - draft structure done", ""),
    (PJ, "X", "Summary of Key Outcomes", IN_, "The results the project delivered.", "Fill in as results come in.", "Files found 4x faster", ""),
    # ---------------- Task Tracker
    (TK, "A", "Task ID", IN_, "The task's ID.", "One format for everyone, e.g. T-001.", "T-001", ""),
    (TK, "B", "Project ID", IN_, "The project this task belongs to.", "Pick from the dropdown (it lists the Project Tracker IDs).", "P-001", ""),
    (TK, "C", "Project Name", FX, "The name of the linked project.", "Looked up from the Project Tracker. Says 'Not in Project Tracker' if the ID isn't found.", "Standardize document filing", ""),
    (TK, "D", "Task Name", IN_, "What needs to be done.", "Start with a verb.", "Draft new folder structure", ""),
    (TK, "E", "Impact", IN_, "How much it matters if this task isn't done.", "Pick High, Medium or Low.", "High", f"{PM} tab (how to rate)"),
    (TK, "F", "Urgency", IN_, "How soon it needs to be done.", "Pick High, Medium or Low.", "Medium", f"{PM} tab (how to rate)"),
    (TK, "G", "Priority (1 = Highest)", FX, "How important the task is overall.", "Looked up from the Priority Matrix using Impact and Urgency.", "2", f"{PM} tab"),
    (TK, "H", "Current Status", IN_, "Where the task stands.", "Pick from the list: Not Started, In Progress, Waiting on Others, On Hold, Completed, Cancelled.", "In Progress", f"{LI} tab"),
    (TK, "I", "Assigned To", IN_, "Who is doing the task.", "Pick from the list.", "Staff member name", f"{LI} tab (Staff)"),
    (TK, "J", "Start Date", IN_, "When work starts.", "A date (mm/dd/yyyy).", "11/02/2026", ""),
    (TK, "K", "Due Date", IN_, "When the task is due.", "A date (mm/dd/yyyy).", "11/15/2026", ""),
    (TK, "L", "Days Left", FX, "Days until the task is due.", "Due Date minus today. Blank once Completed or Cancelled; negative means overdue.", "14", ""),
    (TK, "M", "Due Flag", FX, "Alert for the due date.", "Overdue when past due; Due Soon within 7 days.", "Due Soon", f"{LI} tab (Settings)"),
    (TK, "N", "Frequency", IN_, "How often the task repeats.", "Pick from the list: One-time, Daily, Weekly, Bi-weekly, Monthly, Quarterly, Annually, As Needed.", "Weekly", f"{LI} tab"),
    (TK, "O", "Next Steps", IN_, "The next step on the task.", "A few words.", "Share draft with team", ""),
    (TK, "P", "Notes", IN_, "Anything else about the task.", "Free text.", "Waiting on IT for drive access", ""),
    (TK, "Q", "Results / Outcomes", IN_, "What the task achieved.", "Fill in when done.", "Draft approved", ""),
    (TK, "R", "Date Completed", IN_, "When the task was finished.", "A date (mm/dd/yyyy).", "11/14/2026", ""),
    # ---------------- Issue Tracker
    (IS, "A", "Issue ID", IN_, "The issue's ID.", "One format for everyone, e.g. I-001.", "I-001", ""),
    (IS, "B", "Project ID", IN_, "The project the issue affects.", "Pick from the dropdown.", "P-001", ""),
    (IS, "C", "Project Name", FX, "The name of the linked project.", "Looked up from the Project Tracker.", "Standardize document filing", ""),
    (IS, "D", "Item / Reference", IN_, "The item, part, PO or document the issue is about.", "A number or name.", "Shared drive", ""),
    (IS, "E", "Issue Description", IN_, "What the problem is.", "One or two sentences.", "Drive permissions not granted", ""),
    (IS, "F", "Impact Level", IN_, "How serious the issue is.", "Pick High, Medium or Low.", "High", ""),
    (IS, "G", "Critical Issue?", IN_, "Whether the issue blocks the project. A Yes turns the project Red.", "Pick Yes or No.", "No", ""),
    (IS, "H", "Current Status", IN_, "Where the issue stands.", "Pick from the list: Open, In Progress, Escalated, Resolved, Closed.", "Open", f"{LI} tab"),
    (IS, "I", "Assigned To", IN_, "Who is working the issue.", "Pick from the list.", "Staff member name", f"{LI} tab (Staff)"),
    (IS, "J", "Date of Entry", IN_, "When the issue was logged.", "A date (mm/dd/yyyy).", "10/28/2026", ""),
    (IS, "K", "Target Resolution Date", IN_, "When it should be fixed.", "A date (mm/dd/yyyy).", "11/06/2026", ""),
    (IS, "L", "Days Open", FX, "How long the issue has been open.", "Today minus Date of Entry. Blank once Resolved or Closed.", "4", ""),
    (IS, "M", "Past Target?", FX, "Whether the target date has passed.", "Shows Past Due when the target date has passed and the issue is still open.", "Past Due", ""),
    (IS, "N", "Proposed Solution", IN_, "How you plan to fix it.", "A few words.", "Escalate to IT manager", ""),
    (IS, "O", "Notes", IN_, "Anything else about the issue.", "Free text.", "Ticket 4471", ""),
    (IS, "P", "Date Resolved", IN_, "When it was fixed.", "A date (mm/dd/yyyy).", "11/05/2026", ""),
    # ---------------- other tabs
    (LI, "Yellow cells", "Lists & settings", IN_, "Your portfolio names, team members, dropdown choices and alert windows.", "Fill the yellow cells first. Blue numbers are settings you can change.", "Task 'due soon' window = 7 days", "-"),
    ("Dashboard", "All", "All cells", FX, "Summary and charts of projects, tasks and issues.", "Nothing to type. It updates when the trackers change.", "-", "-"),
]

add_column_guide(
    wb, "Column Guide",
    "Every column on the Project, Task and Issue trackers: whether you fill it in or it is a formula, what it means, and an example. "
    "Examples are shown here, not as sample rows on the trackers, so they never show up on the Dashboard.",
    PG_TYPES, PG_ROWS, index=1)
