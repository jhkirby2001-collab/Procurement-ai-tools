# DPS 365 Report tab: the monthly DPS expiring-contracts report (every department contract ending within a year),
# linked to the Portfolio Tracker. Exec'd inside build_contract_tracker_JHK3.py after the shared-contract parts
# (uses their globals). The report rows are for visibility; only contracts on the Portfolio Tracker are managed here.
DPS_SRC = sys.argv[4] if len(sys.argv) > 4 else None
DPS_N = 300                                   # report rows the tab can hold
D0 = 6                                        # first data row
D1 = D0 + DPS_N - 1
DPS = "'DPS 365 Report'"
DP = lambda c: f"{DPS}!${c}${D0}:${c}${D1}"
PTR = lambda c: f"'Portfolio Tracker'!${c}$2:${c}${LAST}"

# ---------------- Alert Controls section 16
sec(201, "16. DPS 365 report",
    "Settings for the DPS 365 Report tab. Contracts on your tracker are always 'My portfolio'; other report rows are grouped by procurement manager initials.")
hdr(203, {"B": "#", "C": "Setting", "D": "Value", "I": "What it does"})
ctl["B204"] = 1; ctl["B204"].font = font(); ctl["B204"].border = BOX
ctl["C204"] = "Report window (days)"; ctl["C204"].font = font(); ctl["C204"].border = BOX
inp("D204", 365, "0"); ctl["D204"].fill = fill("FFFF00")
note("I204", "An Owned tracker contract ending within this many days that is NOT on the DPS report is flagged 'Missing - check with DPS'. Shared contracts are on the owning department's report and are never flagged.")
ctl.merge_cells("E204:H204")
DPS_WIN = f"{AC}!$D$204"
hdr(206, {"B": "Slot", "C": "Procurement manager initials", "D": "Group", "I": "Notes"})
INI = [("JK", "My portfolio"), ("MJ", "My portfolio"), ("LJ", "My portfolio"), ("JR", "Visibility only"), ("MA", "Visibility only")]
I0, I1 = 207, 216
for i, rr in enumerate(range(I0, I1 + 1)):
    ctl[f"B{rr}"] = i + 1; ctl[f"B{rr}"].font = font(color=GREY_HDR); ctl[f"B{rr}"].border = BOX
    a, b = INI[i] if i < len(INI) else (None, None)
    inp(f"C{rr}", a); inp(f"D{rr}", b)
    ctl.merge_cells(f"E{rr}:H{rr}")
note(f"I{I0}", "Report rows whose PO is on your Portfolio Tracker are 'My portfolio' whatever the initials.")
INI_K, INI_G = f"{AC}!$C${I0}:$C${I1}", f"{AC}!$D${I0}:$D${I1}"

# ---------------- the tab
dp = wb.create_sheet("DPS 365 Report")
dp.sheet_view.showGridLines = False
dp["A1"] = "DPS 365 Report"; dp["A1"].font = font(size=16, bold=True, color=NAVY)
dp["A2"] = ("Every department contract ending within a year, as sent to DPS each month. Columns A-K are the DPS report (procurement managers "
            "fill in A-K and return it to DPS by the first of next month). Grey columns compare each row with your Portfolio Tracker. "
            "Dates on the tracker are never changed from here.")
dp["A2"].font = font(size=9, italic=True, color=GREY_HDR); dp.merge_cells("A2:N2"); dp["A2"].alignment = Alignment(wrap_text=True, vertical="top")
dp.row_dimensions[2].height = 30
dp["A3"] = "DPS legend: RED = 12 months or less, YELLOW = 13-18 months, GREEN = 19-24 months."
dp["A3"].font = font(size=9, color=GREY_HDR)

src_heads, src_rows = [], []
if DPS_SRC:
    sws = openpyxl.load_workbook(DPS_SRC, data_only=True).worksheets[0]
    hdr_row = next(r for r in range(1, 10) if str(sws.cell(r, 4).value).strip().upper() == "PO NUMBER")
    src_heads = [sws.cell(hdr_row, c).value for c in range(1, 12)]
    for r in range(hdr_row + 1, sws.max_row + 1):
        v = [sws.cell(r, c).value for c in range(1, 12)]
        if any(x not in (None, "") for x in v):
            src_rows.append(v)
    # untouched copy of the report
    asis = wb.create_sheet("DPS 365 (as received)")
    for row in sws.iter_rows():
        for c in row:
            d = asis.cell(c.row, c.column, c.value)
            if c.has_style:
                d.font = copy(c.font); d.fill = copy(c.fill); d.border = copy(c.border)
                d.alignment = copy(c.alignment); d.number_format = c.number_format
    for k, dim in sws.column_dimensions.items():
        asis.column_dimensions[k].width = dim.width
DPS_HEADS = [h if h else d for h, d in zip(src_heads or [None] * 11, [
    "Current End Date", "Original Start Date", "Spec Number", "PO Number", "PO Description", "No Longer Needed (X)",
    "Provide Req # for extension request", "Is Extension request within the term of the contract?",
    "Req # for new request that is replacing", "Additional Comments", "Procurement Manager Initials"])]
DPS_HEADS = [" ".join(str(h).split()) for h in DPS_HEADS]

look = lambda rr, col: f"INDEX({PTR(col)},MATCH($D{rr},{PTR('A')},0))"
HELP = [  # (col, header, width, formula)
    ("L", "On My Tracker?", 9, lambda rr: f'=IF($D{rr}="","",IF(ISNUMBER(MATCH($D{rr},{PTR("A")},0)),"Yes","No"))'),
    ("M", "Responsibility", 15, lambda rr: f'=IF($D{rr}="","",IF(L{rr}="Yes","My portfolio",IF($K{rr}="","Not assigned",IFERROR(INDEX({INI_G},MATCH($K{rr},{INI_K},0)),"Visibility only"))))'),
    ("N", "Tracker End Date", 11, lambda rr: f'=IF(L{rr}<>"Yes","",IF({look(rr,"D")}="","",{look(rr,"D")}))'),
    ("O", "End Dates Differ?", 9, lambda rr: f'=IF(OR(N{rr}="",NOT(ISNUMBER($A{rr}))),"",IF(N{rr}<>$A{rr},"Differs",""))'),
    ("P", "Tracker Expiration Status", 12, lambda rr: f'=IF(L{rr}<>"Yes","",{look(rr,"AI")})'),
    ("Q", "Renewal Path", 16, lambda rr: f'=IF(L{rr}<>"Yes","",IF({look(rr,"R")}="","",{look(rr,"R")}))'),
    ("R", "Renewal Stage", 16, lambda rr: f'=IF(L{rr}<>"Yes","",IF({look(rr,"S")}="","",{look(rr,"S")}))'),
    ("S", "Status & Notes (tracker)", 40, lambda rr: f'=IF(L{rr}<>"Yes","",IF({look(rr,"BJ")}="","",{look(rr,"BJ")}))'),
    ("T", "Supply Status", 14, lambda rr: f'=IF(L{rr}<>"Yes","",IF({look(rr,"BG")}="","",{look(rr,"BG")}))'),
    ("U", "Tracker REQ No.", 11, lambda rr: f'=IF(L{rr}<>"Yes","",IF({look(rr,"AO")}="","",{look(rr,"AO")}))'),
    ("V", "Suggested Comment", 46, lambda rr: (
        f'=IF(L{rr}<>"Yes","",TRIM(TEXT(TODAY(),"m/d/yy")&" - "&IF(R{rr}<>"",R{rr},"Stage not set")'
        f'&IF(ISNUMBER(U{rr}),"; Req "&U{rr},"")&IF(T{rr}<>"","; Supply: "&T{rr},"")))')),
]
HR = 5
for j, h in enumerate(DPS_HEADS):
    c = dp.cell(HR, 1 + j, h)
    c.font = font(bold=True, color="FFFFFF"); c.fill = fill(NAVY); c.border = BOX
    c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
for col, h, w, _ in HELP:
    c = dp[f"{col}{HR}"]; c.value = h + " (auto)"
    c.font = font(bold=True, color="FFFFFF"); c.fill = fill(GREY_HDR); c.border = BOX
    c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    dp.column_dimensions[col].width = w
dp.row_dimensions[HR].height = 54
for col, w in zip("ABCDEFGHIJK", [11, 11, 11, 10, 40, 9, 12, 12, 14, 44, 10]):
    dp.column_dimensions[col].width = w
for i in range(DPS_N):
    rr = D0 + i
    v = src_rows[i] if i < len(src_rows) else [None] * 11
    for j in range(11):
        val = v[j]
        if isinstance(val, str) and j in (3,):
            val = val.strip()
            if val.isdigit():
                val = int(val)
        c = dp.cell(rr, 1 + j, val); c.border = BOX; c.font = font(size=9)
        c.alignment = Alignment(vertical="top", wrap_text=(j in (4, 9)))
        if j in (0, 1): c.number_format = "mm/dd/yyyy"
    for col, h, w, fn in HELP:
        c = dp[f"{col}{rr}"]; c.value = fn(rr); c.fill = fill("F2F2F2"); c.border = BOX
        c.font = font(size=9); c.alignment = Alignment(vertical="top", wrap_text=(col in ("S", "V")))
    dp[f"N{rr}"].number_format = "mm/dd/yyyy"
dv(dp, '"X"', f"F{D0}:F{D1}", "Type X if the contract is no longer needed.")
dv(dp, '"Yes,No"', f"H{D0}:H{D1}")
for nm, fc, tc in PALETTE:
    t = f'IFERROR(INDEX({AC}!$G${E0}:$G${E1},MATCH($P{D0},{AC}!$C${E0}:$C${E1},0)),"")="{nm}"'
    dp.conditional_formatting.add(f"P{D0}:P{D1}", FormulaRule(formula=[t], fill=fill(fc), font=Font(color=tc, bold=True), stopIfTrue=True))
dp.conditional_formatting.add(f"O{D0}:O{D1}", FormulaRule(formula=[f'$O{D0}="Differs"'], fill=fill("FFD966"), font=Font(color="5C4400", bold=True)))
dp.conditional_formatting.add(f"M{D0}:M{D1}", FormulaRule(formula=[f'$M{D0}="My portfolio"'], fill=fill("BDD7EE"), font=Font(color="1F3864", bold=True)))
dp.conditional_formatting.add(f"A{D0}:K{D1}", FormulaRule(formula=[f'$M{D0}="My portfolio"'], fill=fill("EEF4FB")))
dp.freeze_panes = f"E{D0}"
dp.auto_filter.ref = f"A{HR}:V{D1}"
dp.sheet_view.zoomScale = 90
dp.page_setup.orientation = "landscape"; dp.page_setup.fitToWidth = 1; dp.page_setup.fitToHeight = 0
dp.sheet_properties.pageSetUpPr.fitToPage = True
dp.print_title_rows = f"{HR}:{HR}"

# ---------------- Portfolio Tracker: AX-AZ
PT_NEW = [("AX", "On DPS 365 Report?", 14,
           lambda r: f'=IF(A{r}="","",IF(ISNUMBER(MATCH(A{r},{DP("D")},0)),"Yes",IF(AND(AT{r}="Owned",ISNUMBER(AD{r}),AD{r}>=0,AD{r}<={DPS_WIN},AH{r}<>{CLOSED_CAT}),"Missing - check with DPS","No")))'),
          ("AY", "DPS 365 Comments", 44,
           lambda r: f'=IF(AX{r}<>"Yes","",IF(INDEX({DP("J")},MATCH(A{r},{DP("D")},0))="","",INDEX({DP("J")},MATCH(A{r},{DP("D")},0))))'),
          ("AZ", "DPS Procurement Mgr", 10,
           lambda r: f'=IF(AX{r}<>"Yes","",IF(INDEX({DP("K")},MATCH(A{r},{DP("D")},0))="","",INDEX({DP("K")},MATCH(A{r},{DP("D")},0))))')]
for col, h, w, fn in PT_NEW:
    c = new[f"{col}1"]; c.value = h + " (auto)"
    c.font = font(bold=True, color="FFFFFF"); c.fill = fill(GREY_HDR); c.border = BOX
    c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    new.column_dimensions[col].width = w
    for r in range(2, LAST + 1):
        x = new[f"{col}{r}"]; x.value = fn(r); x.fill = fill("F2F2F2"); x.border = BOX; x.font = font()
        x.alignment = Alignment(vertical="top", horizontal="left" if col == "AY" else "center", wrap_text=(col == "AY"))
new.conditional_formatting.add(f"AX2:AX{LAST}", FormulaRule(formula=['LEFT($AX2,7)="Missing"'], fill=fill("FFD966"), font=Font(color="5C4400", bold=True)))
new.auto_filter.ref = f"A1:AZ{LAST}"

# ---------------- Dashboard: DPS 365 section
P0 = hr4 + NCODES + 4
sh(f"B{P0}", "DPS 365 report (department contracts ending within a year)")
tile("B", P0 + 1, "ON DPS REPORT", f'=COUNTA({DP("D")})', "All department contracts listed")
tile("E", P0 + 1, "MY PORTFOLIO", f'=COUNTIF({DP("M")},"My portfolio")', "On your tracker or your team's initials")
tile("H", P0 + 1, "VISIBILITY ONLY", f'=COUNTA({DP("D")})-COUNTIF({DP("M")},"My portfolio")', "Managed by others")
tile("K", P0 + 1, "MISSING FROM REPORT", f'=COUNTIF({PTR("AX")},"Missing*")', "Owned, ending within a year, not listed", alert=("FFF4D6", "8A6100"))
tile("N", P0 + 1, "END DATES DIFFER", f'=COUNTIF({DP("O")},"Differs")', "Report vs. your tracker", alert=("FFF4D6", "8A6100"))
tile("Q", P0 + 1, "MY ROWS - NO COMMENT", f'=COUNTIFS({DP("M")},"My portfolio",{DP("J")},"")', "Comments to fill before sending", alert=("FBE3E3", "C00000"))
