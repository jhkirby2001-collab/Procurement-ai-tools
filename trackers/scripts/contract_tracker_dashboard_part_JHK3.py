# Dashboard tab + hidden Dashboard Data tab. Exec'd inside build_contract.py (uses its globals).
from openpyxl.chart import PieChart, BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList

# ---- Alert Controls section 9: dashboard settings
sec(113, "9. Dashboard settings")
hdr(115, {"B": "#", "C": "Setting", "D": "Value", "I": "What it does"})
for i, (k, v, n) in enumerate([
        ("'Expiring soon' window on the Dashboard (days)", 180, "Dashboard tile counts open contracts ending within this many days."),
        ("Top list shows priorities up to", 2, "Contracts with Priority at or below this number appear in the Dashboard's Top 10 list.")]):
    r = 116 + i
    ctl[f"B{r}"] = i + 1; ctl[f"B{r}"].font = font(); ctl[f"B{r}"].border = BOX
    ctl[f"C{r}"] = k; ctl[f"C{r}"].font = font(); ctl[f"C{r}"].border = BOX; ctl[f"C{r}"].alignment = Alignment(wrap_text=True)
    inp(f"D{r}", v, "0"); ctl[f"D{r}"].fill = fill("FFFF00"); note(f"I{r}", n)
    ctl.merge_cells(f"E{r}:H{r}")
SOON, TOPMAX = f"{AC}!$D$116", f"{AC}!$D$117"
CLOSED_CAT = f"{AC}!$K$30"          # the 'Closed' category name
OPEN = f'{PT("AH")},"<>"&{CLOSED_CAT}'

CHART_HEX = {"Dark Red": "C00000", "Red": "E06666", "Amber": "F1C232", "Light Amber": "FFE599",
             "Green": "6AA84F", "Blue": "6FA8DC", "Grey": "A6A6A6", "Purple": "8E7CC3"}
start_colours = [e[4] for e in EXP]           # colour per expiration rule slot at build time
NSTAT = len(EXP)

# ---- hidden data sheet
dd = wb.create_sheet("Dashboard Data")
dd["A1"] = "Helper data for the Dashboard charts and Top 10 list. Do not edit."; dd["A1"].font = font(italic=True, color=GREY_HDR)
# status mix (A3:B12)
dd["A2"], dd["B2"] = "Status", "Contracts"
for i, s in enumerate(range(E0, E1 + 1)):
    r = 3 + i
    dd[f"A{r}"] = f'=IF({AC}!$C${s}="","",{AC}!$C${s})'
    dd[f"B{r}"] = f'=IF(A{r}="","",COUNTIF({PT("AI")},A{r}))'
# expirations by quarter (D3:E13)
dd["D2"], dd["E2"] = "When", "Contracts ending"
dd["G3"] = "=DATE(YEAR(TODAY()),INT((MONTH(TODAY())-1)/3)*3+1,1)"; dd["G2"] = "Quarter start"
dd["D3"] = "Already lapsed"
dd["E3"] = f'=COUNTIFS({PT("AD")},"<0",{OPEN})'
for k in range(8):
    r = 4 + k
    dd[f"G{r}"] = f"=EDATE($G$3,{3*k})"
    lo = "TODAY()" if k == 0 else f"G{r}"
    dd[f"D{r}"] = f'="Q"&(INT((MONTH(G{r})-1)/3)+1)&" "&YEAR(G{r})'
    dd[f"E{r}"] = f'=COUNTIFS({PT("D")},">="&{lo},{PT("D")},"<"&EDATE(G{r},3),{OPEN})'
dd["D12"] = "Later"; dd["G12"] = "=EDATE($G$3,24)"
dd["E12"] = f'=COUNTIFS({PT("D")},">="&G12,{OPEN})'
# lane x status (I2:..)
NL = len(lanes)
dd["I2"] = "Lane"
for j, s in enumerate(range(E0, E0 + NSTAT)):
    dd.cell(2, 10 + j, f'=IF({AC}!$C${s}="","",{AC}!$C${s})')
for i in range(NL):
    r = 3 + i
    dd[f"I{r}"] = f'=IF({AC}!$C${LANES0+i}="","",{AC}!$C${LANES0+i})'
    for j in range(NSTAT):
        col = L(10 + j)
        dd[f"{col}{r}"] = f'=IF(OR($I{r}="",{col}$2=""),0,COUNTIFS({PT("P")},$I{r},{PT("AI")},{col}$2))'
# owner workload (S2:U..)
NO = len(owners)
dd["S2"], dd["T2"], dd["U2"] = "Owner", "Open contracts", "Need action now"
for i in range(NO):
    r = 3 + i
    dd[f"S{r}"] = f'=IF({AC}!$D${OWN0+i}="","",{AC}!$D${OWN0+i})'
    dd[f"T{r}"] = f'=IF(S{r}="",0,COUNTIFS({PT("Q")},S{r},{OPEN}))'
    dd[f"U{r}"] = f'=IF(S{r}="",0,COUNTIFS({PT("Q")},S{r},{PT("AL")},"<="&{TOPMAX}))'
# sort key per tracker row (W2:W300) : priority, then fewest days left
dd["W1"] = "Top-list key"
for r in range(2, LAST + 1):
    dd[f"W{r}"] = (f"=IF(OR('Portfolio Tracker'!AL{r}=\"\",'Portfolio Tracker'!AL{r}>{TOPMAX}),\"\","
                   f"'Portfolio Tracker'!AL{r}*100000+MAX(0,MIN(99998,'Portfolio Tracker'!AD{r}+20000))+ROW()/1000)")
KEY = "'Dashboard Data'!$W$2:$W$%d" % LAST
dd.sheet_state = "hidden"

# ---- Dashboard sheet
db = wb.create_sheet("Dashboard")
db.sheet_view.showGridLines = False
db.column_dimensions["A"].width = 2
for c in range(2, 20): db.column_dimensions[L(c)].width = 11
db["B1"] = "Contract Portfolio Dashboard"; db["B1"].font = font(size=18, bold=True, color=NAVY)
db["B2"] = '="Goods & Services portfolio  |  As of "&TEXT(TODAY(),"mmmm d, yyyy")&"  |  Updates automatically from the Portfolio Tracker"'
db["B2"].font = font(size=9, italic=True, color=GREY_HDR)

TILE_BG, TILE_LINE = "F3F6FB", "C9D3E6"
def tile(col, row, label, frm, sub, fmt="#,##0", alert=None):
    c0, c1 = col, L(CI(col) + 2)
    for rr in range(row, row + 3):
        for cc in range(CI(c0), CI(c1) + 1):
            x = db.cell(rr, cc); x.fill = fill(TILE_BG)
    db.merge_cells(f"{c0}{row}:{c1}{row}"); db.merge_cells(f"{c0}{row+1}:{c1}{row+1}"); db.merge_cells(f"{c0}{row+2}:{c1}{row+2}")
    a = db[f"{c0}{row}"]; a.value = label; a.font = font(size=9, bold=True, color=GREY_HDR); a.alignment = Alignment(horizontal="left", indent=1)
    b = db[f"{c0}{row+1}"]; b.value = frm; b.font = font(size=24, bold=True, color=NAVY); b.number_format = fmt
    b.alignment = Alignment(horizontal="left", indent=1)
    d = db[f"{c0}{row+2}"]; d.value = sub; d.font = font(size=8, color=GREY_HDR); d.alignment = Alignment(horizontal="left", indent=1, wrap_text=True)
    db.row_dimensions[row + 1].height = 34; db.row_dimensions[row + 2].height = 24
    if alert:
        colour, txt = alert
        db.conditional_formatting.add(f"{c0}{row+1}", FormulaRule(formula=[f"AND(ISNUMBER({c0}{row+1}),{c0}{row+1}>0)"],
                                      font=Font(color=txt, bold=True), stopIfTrue=True))
        for rr in range(row, row + 3):
            db.conditional_formatting.add(f"{c0}{rr}:{c1}{rr}", FormulaRule(formula=[f"AND(ISNUMBER(${c0}${row+1}),${c0}${row+1}>0)"],
                                          fill=fill(colour), stopIfTrue=False))

def sh(cell, text):
    db[cell] = text; db[cell].font = font(size=12, bold=True, color=NAVY)

sh("B4", "Contract health")
tile("B", 5, "OPEN CONTRACTS", f'=COUNTA({PT("A")})-COUNTIF({PT("AH")},{CLOSED_CAT})', "Tracked, not closed")
tile("E", 5, "NEED ACTION NOW", f'=COUNTIF({PT("AL")},"<="&{TOPMAX})', "Priority 1-2 (see Top 10 below)", alert=("FBE3E3", "C00000"))
tile("H", 5, "LAPSED", f'=COUNTIFS({PT("AD")},"<0",{OPEN})', "Past end date, not closed", alert=("FBE3E3", "C00000"))
tile("K", 5, "EXPIRING SOON", f'=COUNTIFS({PT("AD")},">=0",{PT("AD")},"<="&{SOON},{OPEN})',
     '="Ending within "&' + SOON + '&" days"', alert=("FFF4D6", "8A6100"))
tile("N", 5, "IN PROCESS", f'=COUNTIF({PT("AH")},{AC}!$K$29)', "Renewal work submitted (incl. lapsed)")
tile("Q", 5, "NEXT ACTIONS OVERDUE", f'=COUNTIF({PT("AK")},{AC}!$C${N0})', "Past their due date", alert=("FBE3E3", "C00000"))
db.column_dimensions["S"].width = 11

sh("B20", "Money")
MON = '$#,##0;($#,##0);"-"'
tile("B", 21, "TOTAL CONTRACT VALUE", f'=SUM({PT("W")})', "Entered so far", MON)
tile("E", 21, "EXPENDED TO DATE", f'=SUM({PT("X")})', "Entered so far", MON)
tile("H", 21, "REMAINING BALANCE", f'=SUM({PT("Y")})', "Value minus expended", MON)
tile("K", 21, "FUNDS AT RISK", f'=COUNTIF({PT("AC")},{AC}!$C${U0})', "Contracts with a red funds flag", alert=("FBE3E3", "C00000"))
tile("N", 21, "FUNDS RUN OUT EARLY", f'=COUNTIF({PT("BC")},"Yes*")', '="Before end date; "&COUNT(' + PT("X") + ')&" of "&COUNTA(' + PT("A") + ')&" have spending entered"', alert=("FBE3E3", "C00000"))
tile("Q", 21, "RENEWAL PATH SET", f'=IFERROR(COUNTA({PT("R")})/COUNTA({PT("A")}),0)', "Share of contracts with a plan", "0%")

# ---- charts
def colour_series(series, hexes):
    for i, h in enumerate(hexes):
        pt = DataPoint(idx=i); pt.graphicalProperties.solidFill = h; pt.graphicalProperties.line.solidFill = "FFFFFF"
        series.dPt.append(pt)

sh("B30", "Where the portfolio stands")
pie = PieChart(); pie.title = "Contracts by expiration status"
pie.add_data(Reference(dd, min_col=2, min_row=3, max_row=2 + NSTAT), titles_from_data=False)
pie.set_categories(Reference(dd, min_col=1, min_row=3, max_row=2 + NSTAT))
colour_series(pie.series[0], [CHART_HEX[c] for c in start_colours])
pie.dataLabels = DataLabelList(showVal=True, showPercent=False, showCatName=False, showSerName=False, showLegendKey=False, showLeaderLines=False)
pie.height, pie.width = 8.5, 13
db.add_chart(pie, "B32")

qb = BarChart(); qb.type = "col"; qb.title = "When contracts end (open contracts)"; qb.legend = None
qb.add_data(Reference(dd, min_col=5, min_row=3, max_row=12), titles_from_data=False)
qb.set_categories(Reference(dd, min_col=4, min_row=3, max_row=12))
colour_series(qb.series[0], ["C00000"] + ["1F3864"] * 9)
qb.dataLabels = DataLabelList(showVal=True, showPercent=False, showCatName=False, showSerName=False, showLegendKey=False)
qb.y_axis.majorGridlines = None; qb.y_axis.title = "Contracts"
qb.height, qb.width = 8.5, 17
db.add_chart(qb, "J32")

sh("B50", "By lane and by owner")
lb = BarChart(); lb.type = "bar"; lb.grouping = "stacked"; lb.overlap = 100; lb.title = "Expiration status by lane"
lb.add_data(Reference(dd, min_col=10, max_col=9 + NSTAT, min_row=2, max_row=2 + NL), titles_from_data=True)
lb.set_categories(Reference(dd, min_col=9, min_row=3, max_row=2 + NL))
for sr, c in zip(lb.series, start_colours):
    sr.graphicalProperties.solidFill = CHART_HEX[c]; sr.graphicalProperties.line.solidFill = CHART_HEX[c]
lb.x_axis.scaling.orientation = "maxMin"; lb.y_axis.majorGridlines = None
lb.height, lb.width = 9, 13
db.add_chart(lb, "B52")

ob = BarChart(); ob.type = "bar"; ob.title = "Workload by owner"
ob.add_data(Reference(dd, min_col=20, max_col=21, min_row=2, max_row=2 + NO), titles_from_data=True)
ob.set_categories(Reference(dd, min_col=19, min_row=3, max_row=2 + NO))
ob.series[0].graphicalProperties.solidFill = "6FA8DC"; ob.series[1].graphicalProperties.solidFill = "C00000"
ob.x_axis.scaling.orientation = "maxMin"; ob.y_axis.majorGridlines = None
ob.height, ob.width = 9, 17
db.add_chart(ob, "J52")

# ---- Top 10 list
T0 = 72
sh(f"B{T0}", "Top 10 contracts needing action")
db[f"B{T0+1}"] = '="Priority 1-"&' + TOPMAX + '&", most urgent first, then fewest days left. Full detail is on the Portfolio Tracker."'
db[f"B{T0+1}"].font = font(size=9, italic=True, color=GREY_HDR)
cols = [("B", "#", None, 4), ("C", "PO Number", "A", 11), ("D", "Description", "B", 30), ("G", "Vendor", "E", 22),
        ("I", "Owner", "Q", 16), ("K", "Lane", "P", 11), ("L", "End Date", "D", 11), ("M", "Days Left", "AD", 9),
        ("N", "Expiration", "AI", 12), ("O", "Funds", "AC", 9), ("P", "Next Action", "AK", 11), ("Q", "Action Status", "S", 16)]
merges = {"D": "F", "G": "H", "I": "J", "Q": "S"}
hr = T0 + 2
for col, h, _, _ in cols:
    c = db[f"{col}{hr}"]; c.value = h; c.font = font(bold=True, color="FFFFFF"); c.fill = fill(NAVY); c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if col in merges:
        for cc in range(CI(col) + 1, CI(merges[col]) + 1): db.cell(hr, cc).fill = fill(NAVY)
        db.merge_cells(f"{col}{hr}:{merges[col]}{hr}")
db.row_dimensions[hr].height = 28
for k in range(1, 11):
    r = hr + k
    db[f"U{r}"] = f'=IFERROR(MATCH(SMALL({KEY},{k}),{KEY},0),"")'   # tracker offset (helper, hidden col)
    for col, h, src, _ in cols:
        c = db[f"{col}{r}"]; c.border = BOX; c.font = font(); c.alignment = Alignment(vertical="center", wrap_text=(col in "DG"))
        if src is None:
            c.value = f'=IF($U{r}="","",{k})'
        else:
            c.value = f"=IF($U{r}=\"\",\"\",IF(INDEX('Portfolio Tracker'!${src}$2:${src}${LAST},$U{r})=\"\",\"\",INDEX('Portfolio Tracker'!${src}$2:${src}${LAST},$U{r})))"
        if col in merges:
            for cc in range(CI(col) + 1, CI(merges[col]) + 1): db.cell(r, cc).border = BOX
            db.merge_cells(f"{col}{r}:{merges[col]}{r}")
    db[f"L{r}"].number_format = "mm/dd/yyyy"; db[f"M{r}"].number_format = "0"
    for col in "BKLMNOP": db[f"{col}{r}"].alignment = Alignment(horizontal="center", vertical="center")
    db.row_dimensions[r].height = 42
db.column_dimensions["U"].hidden = True
db[f"B{hr+11}"] = '=IF(COUNT(' + KEY + ')>10,"+ "&(COUNT(' + KEY + ')-10)&" more on the Portfolio Tracker (sort by Priority).","")'
db[f"B{hr+11}"].font = font(size=9, italic=True, color="C00000")
for col, a, b in [("N", E0, E1), ("O", U0, U1), ("P", N0, N1)]:
    rng = f"{col}{hr+1}:{col}{hr+10}"
    for nm, fc, tc in PALETTE:
        t = f'IFERROR(INDEX({AC}!$G${a}:$G${b},MATCH(${col}{hr+1},{AC}!$C${a}:$C${b},0)),"")="{nm}"'
        db.conditional_formatting.add(rng, FormulaRule(formula=[t], fill=fill(fc), font=Font(color=tc, bold=True), stopIfTrue=True))

db.page_setup.orientation = "landscape"; db.page_setup.fitToWidth = 1; db.page_setup.fitToHeight = 0
db.sheet_properties.pageSetUpPr.fitToPage = True
