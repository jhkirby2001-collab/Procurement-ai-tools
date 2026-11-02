"""Build the blank Project Tracker workbook.

Usage: python build_project_tracker_JHK3.py ../outputs/Project_Tracker_JHK3.xlsx
"""
import os
exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "tracker_helpers_JHK3.py")).read())

wb = Workbook()
readme = wb.active; readme.title = "Read Me"
dash = wb.create_sheet("Dashboard")
proj = wb.create_sheet("Project Tracker")
task = wb.create_sheet("Task Tracker")
iss = wb.create_sheet("Issue Tracker")
mat = wb.create_sheet("Priority Matrix")
guide = wb.create_sheet("Guide")
lists = wb.create_sheet("Lists")

# ---------------------------------------------------------------- Lists
title(lists, "Lists & Settings", "Edit these lists to fit your work. Dropdowns across the workbook read from here.")
LISTS = {
    "Portfolios": [None] * 20,
    "Staff / Assigned To": [None] * 20,
    "Project Status": ["Not Started", "In Progress", "On Hold", "Completed", "Cancelled"],
    "Task Status": ["Not Started", "In Progress", "Waiting on Others", "On Hold", "Completed", "Cancelled"],
    "Level": ["High", "Medium", "Low"],
    "Frequency": ["One-time", "Daily", "Weekly", "Bi-weekly", "Monthly", "Quarterly", "Annually", "As Needed"],
    "Issue Status": ["Open", "In Progress", "Escalated", "Resolved", "Closed"],
    "Yes/No": ["Yes", "No"],
}
REF = {}
for col, (name, vals) in enumerate(LISTS.items(), 1):
    c = L(col)
    h = lists[f"{c}4"]; h.value = name
    h.font = font(bold=True, color="FFFFFF"); h.fill = fill(NAVY); h.border = BOX
    for i, v in enumerate(vals):
        cell = lists[f"{c}{5+i}"]; cell.value = v; cell.border = BOX; cell.font = font(color="0000FF")
        if v is None: cell.fill = fill("FFFF00")
    REF[name] = f"Lists!${c}$5:${c}${4+len(vals)}"
    lists.column_dimensions[c].width = 24
lists["A26"] = "Yellow cells: type your portfolio names and team members. They appear in the dropdowns and on the Dashboard."
lists["A26"].font = font(size=9, italic=True, color=GREY)
lists["A29"] = "Settings"; lists["A29"].font = font(size=12, bold=True, color=NAVY)
SET = [("Project 'due soon' window (days)", 14, "Projects due within this many days turn Amber."),
       ("Task 'due soon' window (days)", 7, "Open tasks due within this many days are flagged 'Due Soon'.")]
for i, (k, v, note) in enumerate(SET):
    r = 30 + i
    lists[f"A{r}"] = k; lists[f"A{r}"].font = font()
    lists[f"B{r}"] = v; lists[f"B{r}"].font = font(color="0000FF", bold=True); lists[f"B{r}"].fill = fill("FFFF00")
    lists[f"B{r}"].border = BOX; lists[f"B{r}"].alignment = Alignment(horizontal="center")
    lists[f"C{r}"] = note; lists[f"C{r}"].font = font(size=9, italic=True, color=GREY)
lists.column_dimensions["A"].width = 32
PRJ_WIN, DUE_WIN = "Lists!$B$30", "Lists!$B$31"
lists.freeze_panes = "A5"

NP, NT, NI = 100, 500, 300
PR = lambda col: f"'Project Tracker'!${col}$5:${col}${4+NP}"
TT = lambda col: f"'Task Tracker'!${col}$5:${col}${4+NT}"
IL = lambda col: f"'Issue Tracker'!${col}$5:${col}${4+NI}"
OPEN_ISS = lambda key, extra="": (f'COUNTIFS({IL("B")},{key}{extra})-COUNTIFS({IL("B")},{key}{extra},{IL("H")},"Resolved")'
                                  f'-COUNTIFS({IL("B")},{key}{extra},{IL("H")},"Closed")')

# ---------------------------------------------------------------- Project Tracker
proj_cols = [
    ("Project ID", 10, "input", None),                                   # A
    ("Portfolio", 22, "list", "Portfolios"),                             # B
    ("Project Name", 30, "text", None),                                  # C
    ("Objective", 32, "text", None),                                     # D
    ("Key Deliverables", 30, "text", None),                              # E
    ("Milestones", 28, "text", None),                                    # F
    ("Owner", 16, "list", "Staff / Assigned To"),                        # G
    ("Stakeholders", 22, "text", None),                                  # H
    ("Start Date", 12, "date", None),                                    # I
    ("Due Date", 12, "date", None),                                      # J
    ("Status", 14, "list", "Project Status"),                            # K
    ("Days Left", 9, "auto_int", lambda r: f'=IF(OR(J{r}="",K{r}="Completed",K{r}="Cancelled"),"",J{r}-TODAY())'),  # L
    ("Total Tasks", 8, "auto_int", lambda r: f'=IF(A{r}="","",COUNTIF({TT("B")},A{r}))'),                            # M
    ("Open Tasks", 8, "auto_int", lambda r: f'=IF(A{r}="","",M{r}-O{r}-COUNTIFS({TT("B")},A{r},{TT("H")},"Cancelled"))'),  # N
    ("Tasks Completed", 9, "auto_int", lambda r: f'=IF(A{r}="","",COUNTIFS({TT("B")},A{r},{TT("H")},"Completed"))'),   # O
    ("% Complete", 9, "auto_pct", lambda r: f'=IF(A{r}="","",IFERROR(O{r}/(N{r}+O{r}),0))'),                           # P
    ("Overdue Tasks", 8, "auto_int", lambda r: f'=IF(A{r}="","",COUNTIFS({TT("B")},A{r},{TT("M")},"Overdue"))'),      # Q
    ("Open Issues", 8, "auto_int", lambda r: f'=IF(A{r}="","",{OPEN_ISS(f"A{r}")})'),                                  # R
    ("Health", 10, "auto", lambda r: (f'=IF(A{r}="","",IF(K{r}="Completed","Complete",IF(K{r}="Cancelled","Cancelled",'
                                      f'IF(OR(AND(L{r}<>"",L{r}<0),Q{r}>0,({OPEN_ISS(f"A{r}", f",{IL(chr(71))},{chr(34)}Yes{chr(34)}")})>0),"Red",'
                                      f'IF(OR(AND(L{r}<>"",L{r}<={PRJ_WIN}),R{r}>0),"Amber","Green")))))')),          # S
    ("Potential Risks / Challenges", 30, "text", None),                  # T
    ("Proposed Solutions", 30, "text", None),                            # U
    ("Next Steps", 30, "text", None),                                    # V
    ("Progress Update", 30, "text", None),                               # W
    ("Summary of Key Outcomes", 32, "text", None),                       # X
]
f0, l0 = tracker(proj, "Project Tracker", "One row per project. Group projects under a Portfolio. Tasks and issues link here by Project ID.",
                 proj_cols, NP, "D5")
cf(proj, f"S{f0}:S{l0}", f'$S{f0}="Red"', RED)
cf(proj, f"S{f0}:S{l0}", f'$S{f0}="Amber"', AMBER)
cf(proj, f"S{f0}:S{l0}", f'OR($S{f0}="Green",$S{f0}="Complete")', GREEN)
cf(proj, f"L{f0}:L{l0}", f'AND($L{f0}<>"",$L{f0}<0)', RED)
cf(proj, f"Q{f0}:Q{l0}", f'AND($Q{f0}<>"",$Q{f0}>0)', RED)
cf(proj, f"A{f0}:A{l0}", f'AND($A{f0}<>"",COUNTIF($A${f0}:$A${l0},$A{f0})>1)', RED)

# ---------------------------------------------------------------- Task Tracker
NOT_IN = "Not in Project Tracker"
lookup = lambda r: f'=IF(B{r}="","",IFERROR(INDEX({PR("C")},MATCH(B{r},{PR("A")},0)),"{NOT_IN}"))'
task_cols = [
    ("Task ID", 9, "input", None),                                                    # A
    ("Project ID", 11, "listref", (f"={PR('A')}", "Project ID not found in the Project Tracker.")),  # B
    ("Project Name", 24, "auto", lookup),                                             # C
    ("Task Name", 38, "text", None),                                                  # D
    ("Impact", 9, "list", "Level"),                                                   # E
    ("Urgency", 9, "list", "Level"),                                                  # F
    ("Priority (1 = Highest)", 10, "auto", lambda r: f"=IF(OR(E{r}=\"\",F{r}=\"\"),\"\",INDEX('Priority Matrix'!$C$6:$E$8,"
                                                     f"MATCH(E{r},'Priority Matrix'!$B$6:$B$8,0),MATCH(F{r},'Priority Matrix'!$C$5:$E$5,0)))"),  # G
    ("Current Status", 15, "list", "Task Status"),                                    # H
    ("Assigned To", 15, "list", "Staff / Assigned To"),                               # I
    ("Start Date", 12, "date", None),                                                 # J
    ("Due Date", 12, "date", None),                                                   # K
    ("Days Left", 9, "auto_int", lambda r: f'=IF(OR(K{r}="",H{r}="Completed",H{r}="Cancelled"),"",K{r}-TODAY())'),  # L
    ("Due Flag", 11, "auto", lambda r: f'=IF(L{r}="","",IF(L{r}<0,"Overdue",IF(L{r}<={DUE_WIN},"Due Soon","")))'),  # M
    ("Frequency", 12, "list", "Frequency"),                                           # N
    ("Next Steps", 30, "text", None),                                                 # O
    ("Notes", 30, "text", None),                                                      # P
    ("Results / Outcomes", 30, "text", None),                                         # Q
    ("Date Completed", 12, "date", None),                                             # R
]
f1, l1 = tracker(task, "Task Tracker", "One row per task. Link it to a Project ID, rate Impact and Urgency, and the Priority calculates itself.",
                 task_cols, NT, "E5")
cf(task, f"G{f1}:G{l1}", f'$G{f1}=1', RED)
cf(task, f"G{f1}:G{l1}", f'$G{f1}=2', AMBER)
cf(task, f"M{f1}:M{l1}", f'$M{f1}="Overdue"', RED)
cf(task, f"M{f1}:M{l1}", f'$M{f1}="Due Soon"', AMBER)
cf(task, f"H{f1}:H{l1}", f'$H{f1}="Completed"', GREEN)
cf(task, f"C{f1}:C{l1}", f'$C{f1}="{NOT_IN}"', RED)

# ---------------------------------------------------------------- Issue Tracker
iss_cols = [
    ("Issue ID", 9, "input", None),                                                   # A
    ("Project ID", 11, "listref", (f"={PR('A')}", "Project ID not found in the Project Tracker.")),  # B
    ("Project Name", 24, "auto", lookup),                                             # C
    ("Item / Reference", 18, "input", None),                                          # D
    ("Issue Description", 38, "text", None),                                          # E
    ("Impact Level", 10, "list", "Level"),                                            # F
    ("Critical Issue?", 9, "list", "Yes/No"),                                         # G
    ("Current Status", 13, "list", "Issue Status"),                                   # H
    ("Assigned To", 15, "list", "Staff / Assigned To"),                               # I
    ("Date of Entry", 12, "date", None),                                              # J
    ("Target Resolution Date", 12, "date", None),                                     # K
    ("Days Open", 9, "auto_int", lambda r: f'=IF(OR(J{r}="",H{r}="Resolved",H{r}="Closed"),"",TODAY()-J{r})'),  # L
    ("Past Target?", 10, "auto", lambda r: f'=IF(OR(K{r}="",H{r}="Resolved",H{r}="Closed"),"",IF(K{r}<TODAY(),"Past Due",""))'),  # M
    ("Proposed Solution", 34, "text", None),                                          # N
    ("Notes", 34, "text", None),                                                      # O
    ("Date Resolved", 12, "date", None),                                              # P
]
f2, l2 = tracker(iss, "Issue Tracker", "One row per problem, risk or blocker. Track it to resolution and record what fixed it.",
                 iss_cols, NI, "E5")
cf(iss, f"G{f2}:G{l2}", f'$G{f2}="Yes"', RED)
cf(iss, f"F{f2}:F{l2}", f'$F{f2}="High"', AMBER)
cf(iss, f"M{f2}:M{l2}", f'$M{f2}="Past Due"', RED)
cf(iss, f"H{f2}:H{l2}", f'OR($H{f2}="Resolved",$H{f2}="Closed")', GREEN)
cf(iss, f"C{f2}:C{l2}", f'$C{f2}="{NOT_IN}"', RED)

# ---------------------------------------------------------------- Priority Matrix
title(mat, "Prioritization Matrix", "Impact x Urgency = Priority. The Task Tracker reads this grid; change a score here and every task updates.")
mat["B4"] = "Impact  ↓   /   Urgency  →"; mat["B4"].font = font(bold=True, color=NAVY)
lv = ["High", "Medium", "Low"]
score = [[1, 2, 3], [2, 3, 4], [3, 4, 5]]
for j, u in enumerate(lv):
    c = mat.cell(5, 3 + j, u); c.font = font(bold=True, color="FFFFFF"); c.fill = fill(NAVY)
    c.alignment = Alignment(horizontal="center"); c.border = BOX
for i, im in enumerate(lv):
    c = mat.cell(6 + i, 2, im); c.font = font(bold=True, color="FFFFFF"); c.fill = fill(NAVY); c.border = BOX
    for j in range(3):
        s = score[i][j]
        c = mat.cell(6 + i, 3 + j, s)
        c.font = font(bold=True, size=14, color="0000FF"); c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BOX; c.fill = fill({1: RED, 2: AMBER, 3: "FFF2CC", 4: LIGHT, 5: AUTO}[s])
    mat.row_dimensions[6 + i].height = 30
mat.column_dimensions["A"].width = 3; mat.column_dimensions["B"].width = 24
for c in "CDE": mat.column_dimensions[c].width = 14
mat["G4"] = "What each priority means"; mat["G4"].font = font(bold=True, color=NAVY)
legend = [("Priority", "Meaning", "Suggested response"),
          (1, "Critical - high impact, needed now", "Do today; escalate blockers immediately"),
          (2, "High - significant impact or time pressure", "Schedule this week"),
          (3, "Medium - important but can be planned", "Plan within the next 2 weeks"),
          (4, "Low - minor impact", "Batch with similar work"),
          (5, "Minimal - nice to have", "Delegate, defer or drop")]
for i, row in enumerate(legend):
    for j, v in enumerate(row):
        c = mat.cell(5 + i, 7 + j, v); c.border = BOX
        c.font = font(bold=(i == 0), color="FFFFFF" if i == 0 else "000000")
        if i == 0: c.fill = fill(NAVY)
mat.column_dimensions["G"].width = 9; mat.column_dimensions["H"].width = 40; mat.column_dimensions["I"].width = 40
mat["B11"] = "How to rate"; mat["B11"].font = font(bold=True, color=NAVY)
rate = [("Impact", "High", "Directly affects goals, budget, compliance, or leadership visibility; many people affected."),
        ("Impact", "Medium", "Affects one team or process; a workaround exists."),
        ("Impact", "Low", "Administrative or housekeeping; little consequence if delayed."),
        ("Urgency", "High", "Due within ~5 business days, or a hard deadline others depend on."),
        ("Urgency", "Medium", "Due within ~2-4 weeks."),
        ("Urgency", "Low", "No firm date, or more than a month out.")]
for i, (a, b, c_) in enumerate(rate):
    for j, v in enumerate((a, b, c_)):
        c = mat.cell(12 + i, 2 + j, v); c.font = font(bold=(j < 2)); c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="top")
    mat.merge_cells(start_row=12 + i, start_column=4, end_row=12 + i, end_column=9)
mat["B19"] = "Blue numbers are editable inputs."; mat["B19"].font = font(size=9, italic=True, color=GREY)

# ---------------------------------------------------------------- Dashboard
title(dash, "Project Tracker - Dashboard", "")
dash["A2"] = "=\"As of \"&TEXT(TODAY(),\"mmmm d, yyyy\")&\"  |  All figures update automatically - do not type in this tab.\""
dash["A2"].font = font(size=9, italic=True, color=GREY)
dash.column_dimensions["A"].width = 2
for c, w in zip("BCDEFGHIJ", [30, 12, 3, 30, 12, 3, 30, 12, 12]):
    dash.column_dimensions[c].width = w

def block(top, colL, head, rows):
    hc = dash[f"{colL}{top}"]; hc.value = head
    hc.font = font(bold=True, color="FFFFFF"); hc.fill = fill(NAVY)
    vcol = L(dash[f"{colL}1"].column + 1)
    dash[f"{vcol}{top}"].fill = fill(NAVY)
    for i, (lab, frm, *fmt) in enumerate(rows, 1):
        a = dash[f"{colL}{top+i}"]; a.value = lab; a.font = font(); a.border = BOX
        b = dash[f"{vcol}{top+i}"]; b.value = "=" + frm; b.font = font(bold=True); b.border = BOX
        b.alignment = Alignment(horizontal="right"); b.number_format = fmt[0] if fmt else "#,##0"

NOT_DONE_T = f'{TT("H")},"<>Completed",{TT("H")},"<>Cancelled"'
NOT_DONE_I = f'{IL("H")},"<>Resolved",{IL("H")},"<>Closed"'
dash["B4"] = "KEY METRICS"; dash["B4"].font = font(bold=True, size=12, color=NAVY)
block(5, "B", "Projects", [
    ("Total projects", f'COUNTIF({PR("A")},"?*")'),
    ("In progress", f'COUNTIF({PR("K")},"In Progress")'),
    ("Completed", f'COUNTIF({PR("K")},"Completed")'),
    ("Past due date", f'COUNTIF({PR("L")},"<0")'),
    ("Health: Red", f'COUNTIF({PR("S")},"Red")'),
    ("Health: Amber", f'COUNTIF({PR("S")},"Amber")'),
    ("Health: Green", f'COUNTIF({PR("S")},"Green")'),
    ("Avg % complete (active projects)", f'IFERROR(AVERAGEIFS({PR("P")},{PR("K")},"<>Completed",{PR("K")},"<>Cancelled",{PR("A")},"?*"),0)', "0%"),
])
block(5, "E", "Tasks", [
    ("Total tasks", f'COUNTIF({TT("D")},"?*")'),
    ("Open tasks", f'COUNTIFS({TT("D")},"?*",{NOT_DONE_T})'),
    ("Overdue", f'COUNTIF({TT("M")},"Overdue")'),
    ("Due soon", f'COUNTIF({TT("M")},"Due Soon")'),
    ("Open Priority 1 (critical)", f'COUNTIFS({TT("G")},1,{NOT_DONE_T})'),
    ("Completed", f'COUNTIF({TT("H")},"Completed")'),
    ("% complete (excl. cancelled)", f'IFERROR(COUNTIF({TT("H")},"Completed")/(COUNTIF({TT("D")},"?*")-COUNTIF({TT("H")},"Cancelled")),0)', "0%"),
])
block(5, "H", "Issues", [
    ("Total issues logged", f'COUNTIF({IL("E")},"?*")'),
    ("Open issues", f'COUNTIFS({IL("E")},"?*",{NOT_DONE_I})'),
    ("Critical - open", f'COUNTIFS({IL("G")},"Yes",{NOT_DONE_I})'),
    ("High impact - open", f'COUNTIFS({IL("F")},"High",{NOT_DONE_I})'),
    ("Escalated", f'COUNTIF({IL("H")},"Escalated")'),
    ("Past target date", f'COUNTIF({IL("M")},"Past Due")'),
    ("Avg days open (open issues)", f'IFERROR(AVERAGE({IL("L")}),0)', "0"),
])
cf(dash, "C9", "$C$9>0", RED); cf(dash, "C10", "$C$10>0", RED); cf(dash, "C11", "$C$11>0", AMBER)
cf(dash, "F8", "$F$8>0", RED); cf(dash, "F10", "$F$10>0", RED)
cf(dash, "I8", "$I$8>0", RED); cf(dash, "I11", "$I$11>0", AMBER)

dash["B16"] = "BREAKDOWNS"; dash["B16"].font = font(bold=True, size=12, color=NAVY)
block(17, "B", "Tasks by status", [(s, f'COUNTIF({TT("H")},"{s}")') for s in LISTS["Task Status"]])
block(17, "E", "Open tasks by priority", [(f"Priority {p}", f'COUNTIFS({TT("G")},{p},{NOT_DONE_T})') for p in range(1, 6)])
block(17, "H", "Issues by status", [(s, f'COUNTIF({IL("H")},"{s}")') for s in LISTS["Issue Status"]])
block(26, "B", "Projects by status", [(s, f'COUNTIF({PR("K")},"{s}")') for s in LISTS["Project Status"]])
block(26, "E", "Open tasks by frequency", [(s, f'COUNTIFS({TT("N")},"{s}",{NOT_DONE_T})') for s in LISTS["Frequency"]])
block(26, "H", "Project health", [(s, f'COUNTIF({PR("S")},"{s}")') for s in ["Red", "Amber", "Green", "Complete"]])
for cell, col in (("H27", RED), ("H28", AMBER), ("H29", GREEN)):
    dash[cell].fill = fill(col)

def roster(top, head, list_name, cols):
    dash[f"B{top-1}"] = head; dash[f"B{top-1}"].font = font(bold=True, size=12, color=NAVY)
    src_col = REF[list_name].split("!")[1][1]
    for cl, (h, _) in cols.items():
        c = dash[f"{cl}{top}"]; c.value = h; c.font = font(bold=True, color="FFFFFF"); c.fill = fill(NAVY)
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    dash.row_dimensions[top].height = 28
    for i in range(20):
        r = top + 1 + i; src = f"Lists!${src_col}${5+i}"
        for cl, (h, fn) in cols.items():
            dash[f"{cl}{r}"] = f'=IF({src}="","",{src})' if fn is None else f'=IF(B{r}="","",{fn(r)})'
            dash[f"{cl}{r}"].font = font(); dash[f"{cl}{r}"].border = BOX
    return top + 21

nt = roster(38, "BY PORTFOLIO", "Portfolios", {
    "B": ("Portfolio (from Lists tab)", None),
    "C": ("Projects", lambda r: f'COUNTIF({PR("B")},B{r})'),
    "E": ("Open tasks", lambda r: f'SUMIF({PR("B")},B{r},{PR("N")})'),
    "F": ("Overdue tasks", lambda r: f'SUMIF({PR("B")},B{r},{PR("Q")})'),
    "H": ("Open issues", lambda r: f'SUMIF({PR("B")},B{r},{PR("R")})'),
    "I": ("Red projects", lambda r: f'COUNTIFS({PR("B")},B{r},{PR("S")},"Red")'),
})
cf(dash, "F39:F58", 'AND(F39<>"",F39>0)', RED); cf(dash, "I39:I58", 'AND(I39<>"",I39>0)', RED)
roster(nt + 3, "WORKLOAD BY PERSON", "Staff / Assigned To", {
    "B": ("Assigned to (from Lists tab)", None),
    "C": ("Open tasks", lambda r: f'COUNTIFS({TT("I")},B{r},{NOT_DONE_T})'),
    "E": ("Overdue tasks", lambda r: f'COUNTIFS({TT("I")},B{r},{TT("M")},"Overdue")'),
    "F": ("Open P1 tasks", lambda r: f'COUNTIFS({TT("I")},B{r},{TT("G")},1,{NOT_DONE_T})'),
    "H": ("Open issues", lambda r: f'COUNTIFS({IL("I")},B{r},{NOT_DONE_I})'),
    "I": ("Projects owned", lambda r: f'COUNTIF({PR("G")},B{r})'),
})
w0 = nt + 4
cf(dash, f"E{w0}:E{w0+19}", f'AND(E{w0}<>"",E{w0}>0)', RED)

def chart(t, cats, vals, anchor):
    ch = BarChart(); ch.type = "bar"; ch.title = t; ch.legend = None
    ch.add_data(vals, titles_from_data=False); ch.set_categories(cats)
    ch.height, ch.width = 7, 13; ch.y_axis.majorGridlines = None
    ch.series[0].graphicalProperties.solidFill = NAVY
    dash.add_chart(ch, anchor)
chart("Tasks by status", Reference(dash, min_col=2, min_row=18, max_row=23), Reference(dash, min_col=3, min_row=18, max_row=23), "K4")
chart("Projects by status", Reference(dash, min_col=2, min_row=27, max_row=31), Reference(dash, min_col=3, min_row=27, max_row=31), "K19")
chart("Issues by status", Reference(dash, min_col=8, min_row=18, max_row=22), Reference(dash, min_col=9, min_row=18, max_row=22), "K34")

# ---------------------------------------------------------------- Guide
title(guide, "Guide", "Working rhythm, prioritization checklist, ownership principles and an escalation template.")
guide.column_dimensions["A"].width = 3; guide.column_dimensions["B"].width = 36; guide.column_dimensions["C"].width = 95
r = 4
def sec(ws, t):
    global r
    r += 1; c = ws[f"B{r}"]; c.value = t; c.font = font(bold=True, size=12, color="FFFFFF"); c.fill = fill(NAVY)
    ws[f"C{r}"].fill = fill(NAVY); r += 1
def line(ws, a, b, f_=None):
    global r
    ws[f"B{r}"] = a; ws[f"C{r}"] = b
    ws[f"B{r}"].font = font(bold=True); ws[f"C{r}"].font = font()
    for c in "BC":
        ws[f"{c}{r}"].alignment = Alignment(wrap_text=True, vertical="top"); ws[f"{c}{r}"].border = BOX
    if f_: ws[f"B{r}"].fill = fill(f_)
    r += 1

sec(guide, "1. Working rhythm")
line(guide, "Daily", "Update task statuses; add new tasks and issues; check the Dashboard for Overdue and Priority 1 items.")
line(guide, "Weekly", "Review every project's Health, Next Steps and Progress Update; clear or escalate blockers; send a status update to your leader.")
line(guide, "Monthly", "Review portfolio-level progress; record Key Outcomes; close completed projects; re-rate priorities.")
line(guide, "Quarterly", "Check projects still align with goals; capture lessons learned; plan the next quarter's projects.")

sec(guide, "2. Prioritization checklist - 10 questions")
for a, b in [
    ("1. Align tasks with strategic goals", "Which task has the greatest impact on achieving business or organizational goals?"),
    ("2. Assess urgency & time sensitivity", "Which task has the most immediate and irreversible consequences if delayed?"),
    ("3. Consider stakeholder importance", "Which task impacts the most important stakeholders?"),
    ("4. Evaluate dependencies", "Which task unlocks progress for other tasks or teams?"),
    ("5. Analyze risk of incomplete tasks", "Which task carries the highest risk or cost if it is not completed on time?"),
    ("6. Apply the 80/20 rule", "Which task delivers the most significant results with the least effort?"),
    ("7. Prioritize based on your expertise", "Which task requires your direct input versus delegation?"),
    ("8. Balance quick wins vs. long-term", "Which task provides immediate visible impact vs. sustained long-term benefit?"),
    ("9. Use a prioritization framework", "How do tasks rank on Impact and Urgency? (see Priority Matrix tab)"),
    ("10. Communicate and adjust", "Are there conflicting priorities, and can adjustments be made?"),
]: line(guide, a, b)

sec(guide, "3. Owning your portfolio - what it means in practice")
for a, b in [
    ("End-to-end accountability", "You own every project in your portfolio from start to finish: objectives, deliverables, timeline and results."),
    ("Proactive leadership", "Anticipate problems and act before they escalate. Send updates before you are asked."),
    ("Decision-making", "Gather data, consult stakeholders, decide, and document your reasoning so you can explain it later."),
    ("Drive results", "Focus on outcomes, not just completed tasks. Record them in Results / Outcomes and Summary of Key Outcomes."),
    ("Manage stakeholders", "Keep your team, partners and leadership informed; address concerns promptly and transparently."),
    ("Communicate effectively", "Use Context -> Challenge -> Solution -> Impact. Escalate with a recommendation, not just a problem."),
    ("Continuous improvement", "After each project, capture lessons learned and suggest process improvements."),
    ("What your leader expects", "Confidence, visibility without having to ask, problem-solving, and results that align with goals."),
]: line(guide, a, b)

sec(guide, "4. Where to record results")
line(guide, "Task Tracker", "Results / Outcomes: the specific result of each task (e.g. 'Report delivered; 3 follow-up items identified').")
line(guide, "Project Tracker", "Summary of Key Outcomes: the high-level result of the project for leadership reporting.")

sec(guide, "5. Issue escalation - steps & email template")
for a, b in [
    ("1. Gather & analyze data", "Confirm the facts, the root cause, current status, and any constraints or delays."),
    ("2. Develop a recommendation", "Root cause - immediate actions - long-term fix."),
    ("3. Communicate to leadership", "Send a short, actionable update (template below)."),
    ("4. Follow up", "No reply in 2 business days? Send a reminder, set the issue to Escalated, and document the outcome."),
]: line(guide, a, b)
r += 1
for t in ["Subject: [Project / Issue] - Update and Recommendation", "", "Hi [Name],", "",
          "Here is the update on [issue]:", "",
          "Situation: [What is happening, in one or two sentences]",
          "Impact: [Effect on schedule, cost, service or stakeholders]",
          "Status: [What has been done so far / current data]",
          "Recommendation: I recommend [clear next step], by [date].", "",
          "Let me know if you would like me to explore any additional angles.", "", "Best,", "[Your name]"]:
    guide[f"C{r}"] = t; guide[f"C{r}"].font = font(); r += 1

# ---------------------------------------------------------------- Read Me
title(readme, "Project Tracker", "Blank template - start here.")
readme.column_dimensions["A"].width = 3; readme.column_dimensions["B"].width = 24; readme.column_dimensions["C"].width = 100
r = 4
sec(readme, "Tabs in this workbook")
line(readme, "Column Guide", "Every column on the Project, Task and Issue trackers: whether you fill it in or it is a formula, what it means, and an example.")
line(readme, "Dashboard", "Live summary of projects, tasks, issues, portfolios and workload. Formulas only - nothing to type.")
line(readme, "Project Tracker", "One row per project, grouped by Portfolio. Task counts, % complete and Health calculate automatically.")
line(readme, "Task Tracker", "Your task list. Link each task to a Project ID, set Impact + Urgency (Priority calculates), Status, Due Date and Frequency.")
line(readme, "Issue Tracker", "Problems, risks and blockers. Track each to resolution.")
line(readme, "Priority Matrix", "The Impact x Urgency scoring grid used by the Task Tracker, plus how to rate each.")
line(readme, "Guide", "Working rhythm, 10-question prioritization checklist, ownership principles, escalation steps and email template.")
line(readme, "Lists", "Dropdown values and settings. Add your portfolio names and team members here first.")

sec(readme, "Getting started")
line(readme, "Step 1", "Lists tab: type your portfolio names and team members in the yellow cells.")
line(readme, "Step 2", "Project Tracker: add each project with a Project ID (e.g. P-001), its Portfolio, objective, dates and status.")
line(readme, "Step 3", "Task Tracker: add tasks (T-001...) and pick the Project ID from the dropdown. Set Impact, Urgency, Status, Due Date, Frequency.")
line(readme, "Step 4", "Issue Tracker: log blockers (I-001...) against a Project ID.")
line(readme, "Step 5", "Check the Dashboard daily. Use the filter arrows on each tracker to sort by due date, priority or person.")

sec(readme, "Colour key")
line(readme, "Navy header", "Column you fill in."); readme[f"B{r-1}"].fill = fill(NAVY); readme[f"B{r-1}"].font = font(bold=True, color="FFFFFF")
line(readme, "Grey header (auto)", "Formula column - do not type over it."); readme[f"B{r-1}"].fill = fill(GREY); readme[f"B{r-1}"].font = font(bold=True, color="FFFFFF")
line(readme, "Red", "Overdue, Priority 1, critical issue, Red health, or duplicate ID.", RED)
line(readme, "Amber", "Due soon, Priority 2, high-impact issue, Amber health.", AMBER)
line(readme, "Green", "Completed, resolved, Green health.", GREEN)
line(readme, "Blue text / yellow fill", "Editable settings and scores (Lists tab, Priority Matrix).", "FFFF00"); readme[f"B{r-1}"].font = font(bold=True, color="0000FF")

sec(readme, "How the numbers are calculated")
line(readme, "Priority", "Looked up from the Priority Matrix using Impact and Urgency (1 = highest, 5 = lowest).")
line(readme, "Days Left / Due Flag", "Due Date minus today, for tasks that are not Completed/Cancelled. Overdue if negative; Due Soon within 7 days (Lists tab).")
line(readme, "% Complete", "Project: completed tasks / (open + completed tasks). Cancelled tasks are excluded.")
line(readme, "Health", "Red = project past due, any overdue task, or an open critical issue.  Amber = due within 14 days (Lists tab) or any open issue.  Green = on track.")
line(readme, "By Portfolio", "Dashboard totals per portfolio roll up from the projects in that portfolio.")
line(readme, "Capacity", "Projects: 100 rows. Tasks: 500 rows. Issues: 300 rows.")

sec(readme, "Example (illustration only - not part of the trackers)")
line(readme, "Project Tracker", "P-001 | Portfolio: Process Improvement | Project: Standardize document filing | Owner: [name] | Due 12/31/2026 | In Progress")
line(readme, "Task Tracker", "T-001 | P-001 | Draft new folder structure | Impact High, Urgency Medium -> Priority 2 | In Progress | Due 11/15/2026 | One-time")
line(readme, "Issue Tracker", "I-001 | P-001 | Shared drive permissions not granted | High | Critical: No | Open | Target 11/06/2026")

exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "project_tracker_column_guide_part_JHK3.py")).read())
wb.properties.title = "Project Tracker"
wb.save(OUT)
print("saved", OUT)
