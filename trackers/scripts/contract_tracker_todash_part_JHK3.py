# Task Order Dashboard + Task Order How To. Exec'd inside taskorder_part.py (uses its globals).
_main_db = db
tdb = wb.create_sheet("Task Order Dashboard")
db = tdb                                  # tile() / sh() write to the global `db`
tdb.sheet_view.showGridLines = False
tdb.column_dimensions["A"].width = 2
for c in range(2, 20): tdb.column_dimensions[L(c)].width = 11
tdb.column_dimensions["K"].width = 14
tdb["B1"] = "Task Order Dashboard"; tdb["B1"].font = font(size=18, bold=True, color=NAVY)
tdb["B2"] = ('="Task orders and job order releases against master contracts  |  As of "&TEXT(TODAY(),"mmmm d, yyyy")&'
             '"  |  Updates automatically from the Task Order Tracker"')
tdb["B2"].font = font(size=9, italic=True, color=GREY_HDR)

# ---------- helper data (Dashboard Data, columns AH onward)
NRS, NRF, NRT = len(RS), len(RF), len(RT)
dd["AH1"], dd["AI1"], dd["AJ1"] = "Release status", "Releases", "Amount ($)"
for i, s in enumerate(range(RS0, RS1 + 1)):
    r = 2 + i
    dd[f"AH{r}"] = f'=IF({AC}!$C${s}="","",{AC}!$C${s})'
    dd[f"AI{r}"] = f'=IF(AH{r}="","",COUNTIF({TO("N")},AH{r}))'
    dd[f"AJ{r}"] = f'=IF(AH{r}="","",SUMIFS({TO("L")},{TO("N")},AH{r}))'
dd["AL1"], dd["AM1"] = "Release flag", "Releases"
for i, s in enumerate(range(RF0, RF1 + 1)):
    r = 2 + i
    dd[f"AL{r}"] = f'=IF({AC}!$C${s}="","",{AC}!$C${s})'
    dd[f"AM{r}"] = f'=IF(AL{r}="","",COUNTIF({TO("Z")},AL{r}))'
dd["AO1"], dd["AP1"] = "Release type", "Committed ($)"
for i, s in enumerate(range(RT0, RT1 + 1)):
    r = 2 + i
    dd[f"AO{r}"] = f'=IF({AC}!$C${s}="","",{AC}!$C${s})'
    dd[f"AP{r}"] = f'=IF(AO{r}="","",SUMIFS({TO("L")},{TO("F")},AO{r},{TO("Q")},"Yes"))'
dd["AR1"], dd["AS1"], dd["AT1"] = "Month start", "Committed issued ($)", "Month"
for k in range(12):
    r = 2 + k
    dd[f"AR{r}"] = f"=DATE(YEAR(TODAY()),MONTH(TODAY())-{11-k},1)"
    dd[f"AT{r}"] = f'=TEXT(AR{r},"mmm yy")'
    dd[f"AS{r}"] = f'=SUMIFS({TO("L")},{TO("Q")},"Yes",{TO("I")},">="&AR{r},{TO("I")},"<"&EDATE(AR{r},1))'
dd["AV4"] = "Flagged-release key"
for r in range(5, TOL + 1):
    sv = f'IFERROR(INDEX({AC}!$H${RF0}:$H${RF1},MATCH({TOS}!Z{r},{AC}!$C${RF0}:$C${RF1},0)),99)'
    dd[f"AV{r}"] = (f'=IF({TOS}!Z{r}="","",{sv}*100000+MAX(0,MIN(99998,IF(ISNUMBER({TOS}!Y{r}),{TOS}!Y{r},70000)+20000))+ROW()/1000)')
RKEY = f"'Dashboard Data'!$AV$5:$AV${TOL}"

# ---------- tiles
NOT_CXL = lambda: " - ".join([])  # placeholder (unused)
sh("B4", "Masters")
tile("B", 5, "MASTER CONTRACTS", f'=COUNTIF({DDc("Y")},"Yes")', "Contract Type takes releases")
tile("E", 5, "TOTAL CEILING", f'=SUMIFS({PTc("W")},{DDc("Y")},"Yes")', "Total Contract Value of masters", MON)
tile("H", 5, "COMMITTED", f'=SUM({DDc("Z")})', "Releases that count as committed", MON)
tile("K", 5, "AVAILABLE", f'=SUM({DDc("AF")})', "Ceiling minus committed", MON)
tile("N", 5, "PAID ON RELEASES", f'=SUM({DDc("AA")})', "Paid to date, all releases", MON)
tile("Q", 5, "MASTERS NEAR CEILING", f'=COUNTIF({DDc("AD")},"<=3")', "Commitment flag severity 1-3", alert=("FFF4D6", "8A6100"))
sh("B9", "Releases")
tile("B", 10, "RELEASES LOGGED", f'=COUNTA({TO("B")})', "All rows on the Task Order Tracker")
tile("E", 10, "OPEN RELEASES", f'=COUNTIF({TO("R")},"Yes")', "Status counts as still open")
tile("H", 10, "COMMITTED RELEASES", f'=COUNTIF({TO("Q")},"Yes")', "Status counts as committed")
tile("K", 10, "RELEASES FLAGGED", f'=COUNTIF({TO("Z")},"?*")', "Any release flag (list below)", alert=("FBE3E3", "C00000"))
tile("N", 10, "ENDING IN 30 DAYS", f'=COUNTIFS({TO("Y")},">=0",{TO("Y")},"<=30")', "Open releases", alert=("FFF4D6", "8A6100"))
tile("Q", 10, "PENDING APPROVAL", f'=SUMIFS({TO("L")},{TO("N")},{AC}!$C${RS0+1})', '="$ in status: "&' + f"{AC}!$C${RS0+1}", MON)

# ---------- master table (rows 15-31) + chart beside it
M0 = 14
sh(f"B{M0}", "Master contracts: committed vs. available")
mcols = [("B", "PO", None), ("C", "Description", "E"), ("F", "Ceiling", None), ("G", "Committed", None),
         ("H", "Available", None), ("I", "% Committed", None), ("J", "Paid", None), ("K", "Flag", None)]
hr2 = M0 + 1
for col, h, mg in mcols:
    c = tdb[f"{col}{hr2}"]; c.value = h; c.font = font(bold=True, color="FFFFFF"); c.fill = fill(NAVY); c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center")
    if mg:
        for cc in range(CI(col) + 1, CI(mg) + 1): tdb.cell(hr2, cc).fill = fill(NAVY)
        tdb.merge_cells(f"{col}{hr2}:{mg}{hr2}")
for k in range(1, 16):
    r = hr2 + k
    tdb[f"V{r}"] = f'=IFERROR(SMALL({DDc("AE")},{k})-1,"")'
    idx = lambda rng: f'IF($V{r}="","",INDEX({rng},$V{r}))'
    vals = {"B": f"={idx(PTc('A'))}", "C": f"={idx(PTc('B'))}",
            "F": f'=IF($V{r}="","",IF(INDEX({PTc("W")},$V{r})="","",INDEX({PTc("W")},$V{r})))',
            "G": f"={idx(DDc('Z'))}", "H": f"={idx(DDc('AF'))}", "I": f"={idx(DDc('AB'))}", "J": f"={idx(DDc('AA'))}", "K": f"={idx(DDc('AC'))}"}
    for col, h, mg in mcols:
        c = tdb[f"{col}{r}"]; c.value = vals[col]; c.border = BOX; c.font = font(size=9)
        c.alignment = Alignment(vertical="center", horizontal="left" if col == "C" else "center", wrap_text=(col == "C"))
        if mg:
            for cc in range(CI(col) + 1, CI(mg) + 1): tdb.cell(r, cc).border = BOX
            tdb.merge_cells(f"{col}{r}:{mg}{r}")
    for col in "FGHJ": tdb[f"{col}{r}"].number_format = '$#,##0;($#,##0);"-"'
    tdb[f"I{r}"].number_format = "0%"; tdb[f"K{r}"].font = font(size=9, bold=True)
    tdb.row_dimensions[r].height = 24
tdb.column_dimensions["V"].hidden = True
for nm, fc, tc in PALETTE:
    t = f'IFERROR(INDEX({AC}!$G${CM0}:$G${CM1},MATCH($K{hr2+1},{AC}!$C${CM0}:$C${CM1},0)),"")="{nm}"'
    tdb.conditional_formatting.add(f"K{hr2+1}:K{hr2+15}", FormulaRule(formula=[t], fill=fill(fc), font=Font(color=tc, bold=True), stopIfTrue=True))
tdb.conditional_formatting.add(f"H{hr2+1}:H{hr2+15}", FormulaRule(formula=[f'AND(ISNUMBER(H{hr2+1}),H{hr2+1}<0)'], font=Font(color="C00000", bold=True)))
tdb[f"B{hr2+16}"] = (f'=IF(COUNT({DDc("AE")})=0,"No master contracts yet. Set Contract Type on the Portfolio Tracker (Contract Type column).",'
                     f'IF(COUNT({DDc("AE")})>15,"+ "&(COUNT({DDc("AE")})-15)&" more masters; see the Task Order Tracker.","Ceiling = Total Contract Value on the Portfolio Tracker."))')
tdb[f"B{hr2+16}"].font = font(size=9, italic=True, color=GREY_HDR)

mb = BarChart(); mb.type = "bar"; mb.grouping = "stacked"; mb.overlap = 100; mb.title = "Committed vs. available by master"
mb.add_data(Reference(tdb, min_col=7, max_col=8, min_row=hr2, max_row=hr2 + 15), titles_from_data=True)
mb.set_categories(Reference(tdb, min_col=2, min_row=hr2 + 1, max_row=hr2 + 15))
mb.series[0].graphicalProperties.solidFill = "1F3864"; mb.series[1].graphicalProperties.solidFill = "A9C6E8"
mb.x_axis.scaling.orientation = "maxMin"; mb.y_axis.majorGridlines = None; mb.y_axis.numFmt = '$#,##0'
mb.height, mb.width = 9.6, 13
tdb.add_chart(mb, f"M{M0}")

# ---------- charts row 2 (status pie, flags bar) and row 3 (monthly $, type $)
C1 = 34
sh(f"B{C1}", "Release activity")
STATUS_HEX = ["D9D9D9", "FFD966", "1F3864", "6FA8DC", "6AA84F", "A6A6A6", "E06666"]
pie = PieChart(); pie.title = "Releases by status"
pie.add_data(Reference(dd, min_col=35, min_row=2, max_row=1 + NRS), titles_from_data=False)   # AI
pie.set_categories(Reference(dd, min_col=34, min_row=2, max_row=1 + NRS))                   # AH
colour_series(pie.series[0], STATUS_HEX[:NRS])
pie.dataLabels = DataLabelList(showVal=True, showPercent=False, showCatName=False, showSerName=False, showLegendKey=False, showLeaderLines=False)
pie.height, pie.width = 8.5, 13
tdb.add_chart(pie, f"B{C1+1}")

fb = BarChart(); fb.type = "bar"; fb.title = "Releases by flag"; fb.legend = None
fb.add_data(Reference(dd, min_col=39, min_row=2, max_row=1 + NRF), titles_from_data=False)    # AM
fb.set_categories(Reference(dd, min_col=38, min_row=2, max_row=1 + NRF))                    # AL
colour_series(fb.series[0], [CHART_HEX[x[3]] for x in RF])
fb.dataLabels = DataLabelList(showVal=True, showPercent=False, showCatName=False, showSerName=False, showLegendKey=False)
fb.x_axis.scaling.orientation = "maxMin"; fb.y_axis.majorGridlines = None
fb.height, fb.width = 8.5, 17
tdb.add_chart(fb, f"J{C1+1}")

C2 = C1 + 19
mc = BarChart(); mc.type = "col"; mc.title = "Committed $ issued by month (last 12 months)"; mc.legend = None
mc.add_data(Reference(dd, min_col=45, min_row=2, max_row=13), titles_from_data=False)       # AS
mc.set_categories(Reference(dd, min_col=46, min_row=2, max_row=13))                        # AT
mc.series[0].graphicalProperties.solidFill = "1F3864"
mc.y_axis.majorGridlines = None; mc.y_axis.numFmt = '$#,##0'
mc.height, mc.width = 8.5, 13
tdb.add_chart(mc, f"B{C2}")

tc_ = BarChart(); tc_.type = "col"; tc_.title = "Committed $ by release type"; tc_.legend = None
tc_.add_data(Reference(dd, min_col=42, min_row=2, max_row=1 + NRT), titles_from_data=False)   # AP
tc_.set_categories(Reference(dd, min_col=41, min_row=2, max_row=1 + NRT))                   # AO
tc_.series[0].graphicalProperties.solidFill = "0E5E5A"
tc_.y_axis.majorGridlines = None; tc_.y_axis.numFmt = '$#,##0'
tc_.height, tc_.width = 8.5, 17
tdb.add_chart(tc_, f"J{C2}")

# ---------- flagged releases list
F0 = C2 + 19
sh(f"B{F0}", "Top 10 flagged releases")
tdb[f"B{F0+1}"] = "Most severe flag first, then soonest end date. Full detail on the Task Order Tracker."
tdb[f"B{F0+1}"].font = font(size=9, italic=True, color=GREY_HDR)
fcols = [("B", "#", None, None), ("C", "Release ID", "A", None), ("D", "Master PO", "B", None), ("E", "Release No.", "E", None),
         ("F", "Scope", "G", "H"), ("I", "Amount", "L", None), ("J", "Status", "N", None), ("K", "End Date", "K", None),
         ("L", "Days Left", "Y", None), ("M", "Master Available", "W", None), ("N", "Flag", "Z", "O")]
hr3 = F0 + 2
for col, h, _, mg in fcols:
    c = tdb[f"{col}{hr3}"]; c.value = h; c.font = font(bold=True, color="FFFFFF"); c.fill = fill(NAVY); c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if mg:
        for cc in range(CI(col) + 1, CI(mg) + 1): tdb.cell(hr3, cc).fill = fill(NAVY)
        tdb.merge_cells(f"{col}{hr3}:{mg}{hr3}")
tdb.row_dimensions[hr3].height = 28
for k in range(1, 11):
    r = hr3 + k
    tdb[f"W{r}"] = f'=IFERROR(MATCH(SMALL({RKEY},{k}),{RKEY},0),"")'
    for col, h, src, mg in fcols:
        c = tdb[f"{col}{r}"]; c.border = BOX; c.font = font(size=9)
        c.alignment = Alignment(vertical="center", horizontal="left" if col == "F" else "center", wrap_text=(col == "F"))
        if src is None:
            c.value = f'=IF($W{r}="","",{k})'
        else:
            rng = f"{TOS}!${src}$5:${src}${TOL}"
            c.value = f'=IF($W{r}="","",IF(INDEX({rng},$W{r})="","",INDEX({rng},$W{r})))'
        if mg:
            for cc in range(CI(col) + 1, CI(mg) + 1): tdb.cell(r, cc).border = BOX
            tdb.merge_cells(f"{col}{r}:{mg}{r}")
    tdb[f"I{r}"].number_format = '$#,##0;($#,##0);"-"'; tdb[f"M{r}"].number_format = '$#,##0;($#,##0);"-"'
    tdb[f"K{r}"].number_format = "mm/dd/yyyy"; tdb[f"L{r}"].number_format = "0"; tdb[f"N{r}"].font = font(size=9, bold=True)
    tdb.row_dimensions[r].height = 26
tdb.column_dimensions["W"].hidden = True
for nm, fc, tc in PALETTE:
    t = f'IFERROR(INDEX({AC}!$G${RF0}:$G${RF1},MATCH($N{hr3+1},{AC}!$C${RF0}:$C${RF1},0)),"")="{nm}"'
    tdb.conditional_formatting.add(f"N{hr3+1}:N{hr3+10}", FormulaRule(formula=[t], fill=fill(fc), font=Font(color=tc, bold=True), stopIfTrue=True))
tdb[f"B{hr3+11}"] = f'=IF(COUNT({RKEY})=0,"No flagged releases.",IF(COUNT({RKEY})>10,"+ "&(COUNT({RKEY})-10)&" more on the Task Order Tracker (filter Release Flag).",""))'
tdb[f"B{hr3+11}"].font = font(size=9, italic=True, color=GREY_HDR)
tdb.page_setup.orientation = "landscape"; tdb.page_setup.fitToWidth = 1; tdb.page_setup.fitToHeight = 0
tdb.sheet_properties.pageSetUpPr.fitToPage = True
db = _main_db

# =============== Task Order How To
th_ = wb.create_sheet("Task Order How To")
th_.sheet_view.showGridLines = False
th_.column_dimensions["A"].width = 2; th_.column_dimensions["B"].width = 30; th_.column_dimensions["C"].width = 110
th_["B1"] = "Task Order How To"; th_["B1"].font = font(size=16, bold=True, color=NAVY)
th_["B2"] = "How to track task orders and job order releases against master and construction (JOC) contracts. For staff training and leadership questions."
th_["B2"].font = font(size=9, italic=True, color=GREY_HDR)
_r = [4]
def TH(t):
    r = _r[0] + 1
    c = th_[f"B{r}"]; c.value = t; c.font = font(bold=True, size=12, color="FFFFFF"); c.fill = fill("0E5E5A")
    th_[f"C{r}"].fill = fill("0E5E5A"); _r[0] = r + 1
def TP(a, b=""):
    r = _r[0]
    th_[f"B{r}"] = a; th_[f"C{r}"] = b
    th_[f"B{r}"].font = font(bold=True); th_[f"C{r}"].font = font()
    for c in "BC":
        th_[f"{c}{r}"].alignment = Alignment(wrap_text=True, vertical="top"); th_[f"{c}{r}"].border = BOX
    _r[0] = r + 1

TH("1. What this is for")
TP("Purpose", "Some contracts are master service agreements or construction job order contracts (JOC). Work is ordered against them through task orders or job order releases. Each release commits part of the master's ceiling. This section tracks every release and shows how much of each master is left.")
TP("The three tabs", "Task Order Tracker = where you enter releases.  Task Order Dashboard = charts and lists for you and leadership.  Task Order How To = this guide.  Settings are on Alert Controls, sections 10-14.")

TH("2. Four numbers to understand")
TP("Ceiling", "The most that can be spent on the master. Comes from Total Contract Value on the Portfolio Tracker.")
TP("Committed", "Total of all releases whose status counts as committed (Issued, In progress, Complete, Closed by default). This money is promised even if not yet paid.")
TP("Available", "Ceiling minus Committed. What is left to issue new releases. This is the number to watch.")
TP("Paid", "What has actually been paid on releases. A master can be lightly paid but fully committed.")

TH("3. One-time setup for each master contract")
TP("Step 1", "Portfolio Tracker: in the Contract Type column, pick Master Service or Construction (JOC).")
TP("Step 2", "Portfolio Tracker: enter the master's Total Contract Value. This is the ceiling.")
TP("Step 3", "The master now appears on the Task Order Dashboard. If its Contract Type is blank or Standard, releases against it are flagged NOT A MASTER.")

TH("4. Entering a release (Task Order Tracker)")
TP("Column Guide", "The Column Guide tab lists every Task Order Tracker column with an example value.")
TP("Release ID", "Your own ID, e.g. R-001. Use one format for the whole team.")
TP("Master PO", "Pick the master contract's PO from the dropdown. Master Description, Contract Type, Ceiling and End Date fill in automatically.")
TP("Release / Task Order No.", "The number on the task order or job order release document.")
TP("Release Type", "Task Order, Job Order Release or Release Change Order.")
TP("Description / Scope, Bureau", "What the work is and which bureau requested it.")
TP("Issue / Start / End Date", "Issue Date matters: it decides the order releases are counted against the ceiling.")
TP("Release Amount ($)", "The amount of the release. For a change order, enter only the change (negative for a decrease).")
TP("Paid to Date ($)", "Optional. Update from invoices or the financial system.")
TP("Status", "Draft, Pending approval, Issued, In progress, Complete, Closed or Cancelled. Status decides whether the amount counts as committed.")
TP("M/WBE Goal % / Actual %", "Optional. The gap shows red when actual is below goal.")
TP("Grey (auto) columns", "Do not type over them: Counts as Committed?, Still Open?, Remaining, Master Ceiling, Master End Date, Master Total Committed, Master Available, Committed Through This Release, Days to Release End, Release Flag, M/WBE Gap.")

TH("5. Release flags (default order, first match wins)")
for a, b in [("EXCEEDS CEILING", "Counting this release and every committed release issued on or before it, the master is over its ceiling. Stop and get a ceiling increase or a new contract."),
             ("NOT A MASTER", "The Master PO's Contract Type does not take releases (blank or Standard). Fix the PO or set its Contract Type."),
             ("OUTLASTS MASTER", "The release ends after the master contract ends. Confirm this is allowed, or shorten the release, or extend the master."),
             ("OVERPAID", "Paid to Date is more than the Release Amount. Check invoices or record a change order."),
             ("OVERDUE", "Still open after its end date. Close it out or extend it."),
             ("ENDING SOON", "Open release ending within 30 days.")]:
    TP(a, b)
TP("Commitment flag (per master)", "OVER CEILING at 100% committed, RED at 90%, YELLOW at 75%, otherwise GREEN. Shown on the Task Order Dashboard.")

TH("6. Weekly routine")
TP("Step 1", "Log any new release the day it is issued, with its Issue Date and amount. Update status as it moves.")
TP("Step 2", "Open the Task Order Dashboard. Check Masters Near Ceiling and the Top 10 flagged releases.")
TP("Step 3", "For any master at RED or OVER CEILING, start the funding increase or replacement contract (track it on the Portfolio Tracker's Next Action).")
TP("Step 4", "Monthly: update Paid to Date and close out completed releases.")

TH("7. Change orders and special cases")
TP("Release change order", "Add a new row, type Release Change Order, same Master PO, amount = the change (+ or -). Put the original release number in Notes. This keeps an audit trail.")
TP("Cancelled release", "Set Status to Cancelled. The amount stops counting as committed and goes back to Available. Don't delete the row.")
TP("Pending approval", "Not committed by default, so it doesn't reduce Available. The Dashboard shows the pending total so you can see what's coming. Change this on Alert Controls section 12 if your team prefers.")
TP("No issue date yet", "The release is compared against the master's total committed instead of in date order.")

TH("8. Changing the settings (Alert Controls)")
TP("Section 10", "Contract Types and whether each takes releases.")
TP("Section 11", "Release Types.")
TP("Section 12", "Release Statuses: 'Counts as committed?' and 'Still open?' for each.")
TP("Section 13", "Commitment flag thresholds (100% / 90% / 75%), colours and severity.")
TP("Section 14", "Release flag rules: label, condition, days, colour, severity and order. Conditions come from a fixed list; a new kind of condition needs a formula change.")
TP("Charts", "Chart colours and ranges follow the starting lists. New statuses, flags or types appear in the tables right away; charts may need their data range widened.")

TH("9. Why it works this way (for leadership)")
TP("Track commitments, not just spending", "In public purchasing, a release commits funds when it is issued, before invoices arrive. Watching only spending hides how much of a master is already promised. Tracking committed and available balances is the same idea as encumbrance accounting used in government budgeting.")
TP("One row per release", "Each release is its own record, so it can be totalled, filtered, audited and reported without changing the contract list.")
TP("Rules on one tab", "Thresholds and statuses are team decisions. Keeping them on Alert Controls makes changes visible and consistent.")
TP("Sources", "Encumbrance and budgetary control: Government Finance Officers Association (GFOA) guidance on budgetary control and encumbrances. City procurement rules: Municipal Code of Chicago, Title 2, Chapter 2-92, and Department of Procurement Services (DPS) policies. The 75% / 90% thresholds are team operating assumptions, not outside requirements; confirm release rules (for example whether a release may outlast its master) with DPS.")
