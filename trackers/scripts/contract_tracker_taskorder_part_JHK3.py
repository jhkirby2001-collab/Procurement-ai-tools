# Task Order Tracker tab + master roll-up. Exec'd inside build_contract.py after dashboard_part.py.
NR = 500                                   # release rows
TOL = 4 + NR                               # last release row
TOS = "'Task Order Tracker'"
TO = lambda c: f"{TOS}!${c}$5:${c}${TOL}"

# =============== Alert Controls sections 10-14
def slot_rows(start, n):
    for i in range(n):
        r = start + i
        ctl[f"B{r}"] = i + 1; ctl[f"B{r}"].font = font(color=GREY_HDR); ctl[f"B{r}"].border = BOX
        yield i, r

sec(120, "10. Contract Types  (dropdown on the Portfolio Tracker, last column)",
    "'Takes releases?' = Yes means task orders / job order releases can be issued against it and its ceiling is tracked on the Task Order Tracker.")
hdr(122, {"B": "Slot", "C": "Contract Type", "D": "Takes releases?", "I": "What it means"})
CT = [("Standard", "No", "Single contract; spending tracked on the Portfolio Tracker."),
      ("Master Service", "Yes", "Master / blanket agreement; work is ordered through task orders."),
      ("Construction (JOC)", "Yes", "Job order contract; work is ordered through job order releases.")]
CT0, CT1 = 123, 130
for i, r in slot_rows(CT0, 8):
    a, b, n = CT[i] if i < len(CT) else (None, None, None)
    inp(f"C{r}", a); inp(f"D{r}", b); note(f"I{r}", n); ctl.merge_cells(f"E{r}:H{r}")
dv(ctl, '"Yes,No"', f"D{CT0}:D{CT1}")
CT_NAME, CT_TAKES = f"{AC}!$C${CT0}:$C${CT1}", f"{AC}!$D${CT0}:$D${CT1}"

sec(132, "11. Release Types  (dropdown on the Task Order Tracker)")
hdr(134, {"B": "Slot", "C": "Release Type", "I": "What it means"})
RT = [("Task Order", "Work ordered under a master service contract."),
      ("Job Order Release", "Work ordered under a construction / JOC contract."),
      ("Release Change Order", "Change to an existing release. Enter the change amount (negative for a decrease) and reference the original release in Notes.")]
RT0, RT1 = 135, 142
for i, r in slot_rows(RT0, 8):
    a, n = RT[i] if i < len(RT) else (None, None)
    inp(f"C{r}", a); note(f"I{r}", n); ctl.merge_cells(f"D{r}:H{r}")
RT_NAME = f"{AC}!$C${RT0}:$C${RT1}"

sec(144, "12. Release Statuses  (dropdown on the Task Order Tracker)",
    "'Counts as committed?' = the amount is subtracted from the master's available balance. 'Still open?' = the release can still get date alerts.")
hdr(146, {"B": "Slot", "C": "Release Status", "D": "Counts as committed?", "E": "Still open?", "I": "What it means"})
RS = [("Draft", "No", "Yes", "Being prepared. Not yet committed."),
      ("Pending approval", "No", "Yes", "Waiting on approval. Shown so you can see what is coming, but not committed yet."),
      ("Issued", "Yes", "Yes", "Approved and issued to the vendor. Funds committed."),
      ("In progress", "Yes", "Yes", "Work under way."),
      ("Complete", "Yes", "No", "Work finished; awaiting or done with final payment."),
      ("Closed", "Yes", "No", "Fully paid and closed out."),
      ("Cancelled", "No", "No", "Cancelled; amount released back to the master.")]
RS0, RS1 = 147, 156
for i, r in slot_rows(RS0, 10):
    a, b, c, n = RS[i] if i < len(RS) else (None,) * 4
    inp(f"C{r}", a); inp(f"D{r}", b); inp(f"E{r}", c); note(f"I{r}", n); ctl.merge_cells(f"F{r}:H{r}")
dv(ctl, '"Yes,No"', f"D{RS0}:E{RS1}")
RS_NAME, RS_COMMIT, RS_OPEN = f"{AC}!$C${RS0}:$C${RS1}", f"{AC}!$D${RS0}:$D${RS1}", f"{AC}!$E${RS0}:$E${RS1}"

sec(158, "13. Commitment Flag rules  (how much of a master's ceiling is committed)",
    "First match wins. Matches when % committed >= its limit (blank = always). Uses Total Contract Value on the Portfolio Tracker as the ceiling.")
hdr(160, {"B": "Order", "C": "Flag Label", "D": "% Committed at or over", "G": "Colour", "H": "Severity (1 = most urgent)", "I": "What it means / why"})
CF_ = [("OVER CEILING", 1.0, "Dark Red", 1, "Releases exceed the master's ceiling. Stop issuing; a ceiling increase or new contract is needed."),
       ("RED", 0.90, "Red", 2, "90%+ committed. Start the increase or replacement now."),
       ("YELLOW", 0.75, "Amber", 3, "75%+ committed. Check upcoming needs with the bureau."),
       ("GREEN", None, "Green", 8, "Room left on the master.")]
CM0, CM1 = 161, 166
for i, r in slot_rows(CM0, 6):
    a, b, c, s, n = CF_[i] if i < len(CF_) else (None,) * 5
    inp(f"C{r}", a); inp(f"D{r}", b, "0%"); ctl[f"E{r}"].border = BOX; ctl[f"F{r}"].border = BOX
    inp(f"G{r}", c); inp(f"H{r}", s, "0"); note(f"I{r}", n)
dv(ctl, "=" + COLOR_REF, f"G{CM0}:G{CM1}")

sec(168, "14. Release Flag rules  (one flag per release)",
    "First match wins. Pick a Condition from the fixed list in column K; 'Value' is only used by the days condition.")
hdr(170, {"B": "Order", "C": "Flag Label", "D": "Condition", "E": "Value (days)", "G": "Colour", "H": "Severity (1 = most urgent)", "I": "What it means / why", "K": "Conditions (fixed)"})
CONDS = ["Master over ceiling", "Ends after master", "Days to release end at or under", "Master not a release contract",
         "Paid more than release amount", "Always"]
for i, c in enumerate(CONDS):
    ctl[f"K{171+i}"] = c; ctl[f"K{171+i}"].font = font(); ctl[f"K{171+i}"].border = BOX
COND_LIST = f"{AC}!$K$171:$K${170+len(CONDS)}"
RF = [("EXCEEDS CEILING", "Master over ceiling", None, "Dark Red", 1, "This release (counting releases issued on or before it) pushes the master past its ceiling."),
      ("NOT A MASTER", "Master not a release contract", None, "Purple", 2, "The Master PO's Contract Type does not take releases. Check the PO or set its Contract Type."),
      ("OUTLASTS MASTER", "Ends after master", None, "Red", 2, "Release end date is after the master contract's end date."),
      ("OVERPAID", "Paid more than release amount", None, "Red", 2, "Paid to date is higher than the release amount."),
      ("OVERDUE", "Days to release end at or under", -1, "Red", 2, "Release is still open past its end date."),
      ("ENDING SOON", "Days to release end at or under", 30, "Amber", 3, "Open release ends within 30 days.")]
RF0, RF1 = 171, 178
for i, r in slot_rows(RF0, 8):
    a, b, v, c, s, n = RF[i] if i < len(RF) else (None,) * 6
    inp(f"C{r}", a); inp(f"D{r}", b); inp(f"E{r}", v, "0"); ctl[f"F{r}"].border = BOX
    inp(f"G{r}", c); inp(f"H{r}", s, "0"); note(f"I{r}", n)
dv(ctl, "=" + COND_LIST, f"D{RF0}:D{RF1}")
dv(ctl, "=" + COLOR_REF, f"G{RF0}:G{RF1}")
ctl.column_dimensions["K"].width = 30
for a, b in [(CM0, CM1), (RF0, RF1)]:
    colour_cf(ctl, f"C{a}:C{b}", f'$G{a}="@"')
    colour_cf(ctl, f"G{a}:G{b}", f'$G{a}="@"')

# =============== Portfolio Tracker: one new input column (AR)
pt = wb["Portfolio Tracker"]
c = pt["AR1"]; c.value = "Contract Type"
c.font = font(bold=True, color="FFFFFF"); c.fill = fill(BLUE_HDR); c.border = BOX
c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
pt.column_dimensions["AR"].width = 18
for r in range(2, LAST + 1):
    pt[f"AR{r}"].border = BOX; pt[f"AR{r}"].font = font()
dv(pt, "=" + CT_NAME, f"AR2:AR{LAST}")
pt.auto_filter.ref = f"A1:AR{LAST}"

# =============== Task Order Tracker tab
to = wb.create_sheet("Task Order Tracker")
to.sheet_view.showGridLines = False
to["A1"] = "Task Order Tracker"; to["A1"].font = font(size=16, bold=True, color=NAVY)
to["A2"] = ("One row per task order or job order release, linked to its master contract by PO. Each committed release reduces the "
            "master's available balance. Lists, statuses and flag rules are on Alert Controls sections 10-14.")
to["A2"].font = font(size=9, italic=True, color=GREY_HDR)
to["A3"] = "Blue headers = you type.   Grey headers marked (auto) = formulas, do not type over them."
to["A3"].font = font(size=9, color=GREY_HDR)
PTA = f"'Portfolio Tracker'!$A$2:$A${LAST}"
PTc = lambda c: f"'Portfolio Tracker'!${c}$2:${c}${LAST}"

def look(rr, col):
    return f'INDEX({PTc(col)},MATCH($B{rr},{PTA},0))'

def rf_cond(rr):
    def cnd(s):
        D, E = f"{AC}!$D${s}", f"{AC}!$E${s}"
        return (f'AND({AC}!$C${s}<>"",OR('
                f'AND({D}="Master over ceiling",Q{rr}="Yes",ISNUMBER(T{rr}),ISNUMBER(X{rr}),X{rr}>T{rr}),'
                f'AND({D}="Ends after master",R{rr}="Yes",ISNUMBER(K{rr}),ISNUMBER(U{rr}),K{rr}>U{rr}),'
                f'AND({D}="Days to release end at or under",R{rr}="Yes",ISNUMBER(Y{rr}),{E}<>"",Y{rr}<={E}),'
                f'AND({D}="Master not a release contract",B{rr}<>"",IFERROR(INDEX({CT_TAKES},MATCH(D{rr},{CT_NAME},0)),"No")<>"Yes"),'
                f'AND({D}="Paid more than release amount",ISNUMBER(L{rr}),ISNUMBER(M{rr}),L{rr}>=0,M{rr}>L{rr}),'
                f'{D}="Always"))')
    return cnd

TOCOLS = [
    ("A", "Release ID", "in", 10, None), ("B", "Master PO", "in", 12, None),
    ("C", "Master Description", "auto", 30, lambda r: f'=IF(B{r}="","",IFERROR({look(r,"B")},"PO not on Portfolio Tracker"))'),
    ("D", "Contract Type", "auto", 16, lambda r: f'=IF(B{r}="","",IFERROR(IF({look(r,"AR")}="","Not set",{look(r,"AR")}),""))'),
    ("E", "Release / Task Order No.", "in", 14, None), ("F", "Release Type", "in", 18, None),
    ("G", "Description / Scope", "in", 34, None), ("H", "Requesting Bureau", "in", 14, None),
    ("I", "Issue Date", "in", 11, None), ("J", "Start Date", "in", 11, None), ("K", "End Date", "in", 11, None),
    ("L", "Release Amount ($)", "in", 14, None), ("M", "Paid to Date ($)", "in", 14, None),
    ("N", "Status", "in", 15, None), ("O", "M/WBE Goal %", "in", 9, None), ("P", "M/WBE Actual %", "in", 9, None),
    ("Q", "Counts as Committed?", "auto", 10, lambda r: f'=IF(OR(B{r}="",N{r}=""),"",IFERROR(INDEX({RS_COMMIT},MATCH(N{r},{RS_NAME},0)),"No"))'),
    ("R", "Still Open?", "auto", 8, lambda r: f'=IF(OR(B{r}="",N{r}=""),"",IFERROR(INDEX({RS_OPEN},MATCH(N{r},{RS_NAME},0)),"No"))'),
    ("S", "Remaining on Release ($)", "auto", 14, lambda r: f'=IF(L{r}="","",L{r}-N(M{r}))'),
    ("T", "Master Ceiling ($)", "auto", 14, lambda r: f'=IF(B{r}="","",IFERROR(IF({look(r,"W")}="","",{look(r,"W")}),""))'),
    ("U", "Master End Date", "auto", 11, lambda r: f'=IF(B{r}="","",IFERROR(IF({look(r,"D")}="","",{look(r,"D")}),""))'),
    ("V", "Master Total Committed ($)", "auto", 14, lambda r: f'=IF(B{r}="","",SUMIFS({TO("L")},{TO("B")},B{r},{TO("Q")},"Yes"))'),
    ("W", "Master Available ($)", "auto", 14, lambda r: f'=IF(OR(B{r}="",T{r}=""),"",T{r}-V{r})'),
    ("X", "Committed Through This Release ($)", "auto", 14,
     lambda r: f'=IF(OR(B{r}="",Q{r}<>"Yes"),"",IF(ISNUMBER(I{r}),SUMIFS({TO("L")},{TO("B")},B{r},{TO("Q")},"Yes",{TO("I")},"<="&I{r}),V{r}))'),
    ("Y", "Days to Release End", "auto", 9, lambda r: f'=IF(OR(NOT(ISNUMBER(K{r})),R{r}<>"Yes"),"",K{r}-TODAY())'),
    ("Z", "Release Flag", "auto", 15, lambda r: f'=IF(B{r}="","",' + slot_chain(range(RF0, RF1 + 1), rf_cond(r), "C") + ")"),
    ("AA", "M/WBE Gap", "auto", 9, lambda r: f'=IF(OR(O{r}="",P{r}=""),"",P{r}-O{r})'),
    ("AB", "Notes", "in", 40, None),
]
for col, h, kind, w, fn in TOCOLS:
    c = to[f"{col}4"]; c.value = h + (" (auto)" if kind == "auto" else "")
    c.font = font(bold=True, color="FFFFFF"); c.fill = fill(BLUE_HDR if kind == "in" else GREY_HDR)
    c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center"); c.border = BOX
    to.column_dimensions[col].width = w
    for r in range(5, TOL + 1):
        x = to[f"{col}{r}"]; x.border = BOX; x.font = font()
        x.alignment = Alignment(vertical="top", wrap_text=(col in ("G", "AB", "C")))
        if fn:
            x.value = fn(r); x.fill = fill("F2F2F2"); x.alignment = Alignment(vertical="top", horizontal="center", wrap_text=(col == "C"))
        if col in ("I", "J", "K", "U"): x.number_format = "mm/dd/yyyy"
        elif col in ("L", "M", "S", "T", "V", "W", "X"): x.number_format = '$#,##0;($#,##0);"-"'
        elif col in ("O", "P"): x.number_format = "0%"
        elif col == "AA": x.number_format = '+0%;-0%;0%'
        elif col == "Y": x.number_format = "0"
    if col == "Z":
        for r in range(5, TOL + 1): to[f"Z{r}"].font = font(bold=True)
to.row_dimensions[4].height = 45
dv(to, f"={PTA}", f"B5:B{TOL}", "PO not found on the Portfolio Tracker.")
dv(to, "=" + RT_NAME, f"F5:F{TOL}")
dv(to, "=" + RS_NAME, f"N5:N{TOL}")
v = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True, showErrorMessage=True, error="Enter a date (mm/dd/yyyy).")
to.add_data_validation(v); v.add(f"I5:K{TOL}")
v = DataValidation(type="decimal", allow_blank=True, showErrorMessage=True, error="Enter a dollar amount.", operator="between", formula1="-1000000000000", formula2="1000000000000")
to.add_data_validation(v); v.add(f"L5:M{TOL}")
for nm, fc, tc in PALETTE:
    t = f'IFERROR(INDEX({AC}!$G${RF0}:$G${RF1},MATCH($Z5,{AC}!$C${RF0}:$C${RF1},0)),"")="{nm}"'
    to.conditional_formatting.add(f"Z5:Z{TOL}", FormulaRule(formula=[t], fill=fill(fc), font=Font(color=tc, bold=True), stopIfTrue=True))
to.conditional_formatting.add(f"W5:W{TOL}", FormulaRule(formula=['AND(ISNUMBER($W5),$W5<0)'], fill=fill("F4A6A6"), font=Font(color="7A0000", bold=True)))
to.conditional_formatting.add(f"C5:D{TOL}", FormulaRule(formula=['OR($C5="PO not on Portfolio Tracker",$D5="Not set")'], font=Font(color="C00000", italic=True)))
to.conditional_formatting.add(f"AA5:AA{TOL}", FormulaRule(formula=['AND(ISNUMBER($AA5),$AA5<0)'], fill=fill("F4A6A6"), font=Font(color="7A0000")))
to.freeze_panes = "C5"
to.auto_filter.ref = f"A4:AB{TOL}"
to.sheet_view.zoomScale = 90

# =============== master roll-up per Portfolio Tracker row (hidden Dashboard Data, cols Y-AE)
PT_ = lambda c, r: f"'Portfolio Tracker'!{c}{r}"
for h, col in [("Takes releases?", "Y"), ("Committed ($)", "Z"), ("Paid on releases ($)", "AA"), ("% Committed", "AB"),
               ("Commitment Flag", "AC"), ("Flag severity", "AD"), ("Master list key", "AE"), ("Available ($)", "AF")]:
    dd[f"{col}1"] = h
for r in range(2, LAST + 1):
    dd[f"Y{r}"] = f'=IF({PT_("AR",r)}="","No",IFERROR(INDEX({CT_TAKES},MATCH({PT_("AR",r)},{CT_NAME},0)),"No"))'
    dd[f"Z{r}"] = f'=IF(Y{r}<>"Yes","",SUMIFS({TO("L")},{TO("B")},{PT_("A",r)},{TO("Q")},"Yes"))'
    dd[f"AA{r}"] = f'=IF(Y{r}<>"Yes","",SUMIFS({TO("M")},{TO("B")},{PT_("A",r)}))'
    dd[f"AB{r}"] = f'=IF(AND(Y{r}="Yes",ISNUMBER({PT_("W",r)})),IF({PT_("W",r)}>0,Z{r}/{PT_("W",r)},""),"")'
    cm = lambda s, r=r: f'AND({AC}!$C${s}<>"",IF({AC}!$D${s}="",TRUE,AB{r}>={AC}!$D${s}))'
    dd[f"AC{r}"] = f'=IF(AB{r}="","",' + slot_chain(range(CM0, CM1 + 1), cm, "C") + ")"
    dd[f"AD{r}"] = f'=IF(AC{r}="","",{sev(f"AC{r}",CM0,CM1)})'
    dd[f"AE{r}"] = f'=IF(Y{r}="Yes",ROW(),"")'
    dd[f"AF{r}"] = f'=IF(OR(Y{r}<>"Yes",NOT(ISNUMBER({PT_("W",r)}))),"",{PT_("W",r)}-Z{r})'
DDc = lambda c: f"'Dashboard Data'!${c}$2:${c}${LAST}"

# =============== Dashboard: tiles row + master contracts section
sh("B14", "Task orders and releases")
tile("B", 15, "MASTER CONTRACTS", f'=COUNTIF({DDc("Y")},"Yes")', "Contract Type takes releases")
tile("E", 15, "OPEN RELEASES", f'=COUNTIF({TO("R")},"Yes")', "Still open on the Task Order Tracker")
tile("H", 15, "COMMITTED ON MASTERS", f'=SUM({DDc("Z")})', "Committed releases", MON)
tile("K", 15, "AVAILABLE ON MASTERS", f'=SUM({DDc("AF")})', "Ceiling minus committed", MON)
tile("N", 15, "MASTERS NEAR CEILING", f'=COUNTIF({DDc("AD")},"<=3")', "Commitment flag severity 1-3", alert=("FFF4D6", "8A6100"))
tile("Q", 15, "RELEASES FLAGGED", f'=COUNTIF({TO("Z")},"?*")', "Detail on the Task Order Dashboard", alert=("FBE3E3", "C00000"))

exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "contract_tracker_todash_part_JHK3.py")).read())

# =============== Summary: release status counts
sr = sm.max_row + 2
shead(f"A{sr}", "Task orders and releases")
th(f"A{sr+1}", "Release Status"); th(f"B{sr+1}", "Releases"); th(f"C{sr+1}", "Amount ($)")
for i, s in enumerate(range(RS0, RS1 + 1)):
    r = sr + 2 + i
    td(f"A{r}", f'=IF({AC}!$C${s}="","",{AC}!$C${s})', "@", bold=True)
    td(f"B{r}", f'=IF(A{r}="","",COUNTIF({TO("N")},A{r}))')
    td(f"C{r}", f'=IF(A{r}="","",SUMIFS({TO("L")},{TO("N")},A{r}))', '$#,##0;($#,##0);"-"')
