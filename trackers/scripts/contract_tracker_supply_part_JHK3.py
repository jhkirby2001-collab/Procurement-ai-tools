# Spending pace (burn index, projected exhaustion), supply coverage for lapsed / soon-to-lapse contracts,
# whole-row "Action Needed" highlight, and the lapsed renewal x supply KPIs on the Dashboard and Summary.
# Exec'd inside build_contract_tracker_JHK3.py after the DPS 365 part (uses its globals). Column letters here are
# the build layout; the optional Reviewed column (insert_reviewed_column_JHK3.py) shifts everything one to the right.

# ---------------- Alert Controls sections 17-19
sec(219, "17. Spending pace  (Burn Index = % spent / % of term used)",
    "1.00 = spending in step with time. Above 1 = money is running out faster than time. Change the levels here.")
hdr(221, {"B": "#", "C": "Setting", "D": "Value", "I": "What it does"})
for i, (k, v, fmt, n) in enumerate([
        ("Burn Index - amber at or above", 1.10, "0.00", "Burn Index cell turns amber."),
        ("Burn Index - red at or above", 1.25, "0.00", "Burn Index cell turns red."),
        ("Only calculate after this share of the term has passed", 0.10, "0%",
         "Early in a contract a little spending swings the index wildly, so it stays blank until then (and while nothing is spent).")]):
    r = 222 + i
    ctl[f"B{r}"] = i + 1; ctl[f"B{r}"].font = font(); ctl[f"B{r}"].border = BOX
    ctl[f"C{r}"] = k; ctl[f"C{r}"].font = font(); ctl[f"C{r}"].border = BOX; ctl[f"C{r}"].alignment = Alignment(wrap_text=True)
    inp(f"D{r}", v, fmt); ctl[f"D{r}"].fill = fill("FFFF00"); note(f"I{r}", n); ctl.merge_cells(f"E{r}:H{r}")
BURN_AMB, BURN_RED, MIN_EL = f"{AC}!$D$222", f"{AC}!$D$223", f"{AC}!$D$224"

sec(227, "18. Supply coverage  (how DWM keeps getting materials when a contract lapses or is about to)",
    "Choices for the Supply Coverage dropdown. 'Counts as' drives Supply Risk, the Dashboard and row highlighting. Add your own in the empty slots.")
hdr(229, {"B": "Slot", "C": "Supply Coverage", "D": "Counts as", "I": "What it means", "K": "Counts-as values (fixed)"})
COV = [("Exhibit B", "Covered", "Buying under Exhibit B while the renewal is in process."),
       ("Alternate PO / contract", "Covered", "Another PO or another department's contract covers the need. Put the PO in Coverage PO / Note."),
       ("Extension pending - vendor honoring", "Covered", "Vendor keeps supplying while the extension is processed."),
       ("Stock on hand", "Covered", "Inventory covers needs until the new contract. Put the run-out date in Coverage Through."),
       ("Unconfirmed", "At risk", "Coverage asked about but not confirmed (e.g. 'Can we use Exh B?')."),
       ("No way to buy", "Not covered", "Nothing in place. Escalate."),
       ("Not needed", "Closed", "Requirement ended; no coverage needed.")]
CV0, CV1 = 230, 241
for i, r in enumerate(range(CV0, CV1 + 1)):
    ctl[f"B{r}"] = i + 1; ctl[f"B{r}"].font = font(color=GREY_HDR); ctl[f"B{r}"].border = BOX
    a, b, n = COV[i] if i < len(COV) else (None, None, None)
    inp(f"C{r}", a); inp(f"D{r}", b); note(f"I{r}", n); ctl.merge_cells(f"E{r}:H{r}")
for i, v in enumerate(["Covered", "At risk", "Not covered", "Closed"]):
    ctl[f"K{CV0+i}"] = v; ctl[f"K{CV0+i}"].font = font(); ctl[f"K{CV0+i}"].border = BOX
dv(ctl, f"={AC}!$K${CV0}:$K${CV0+3}", f"D{CV0}:D{CV1}")
COV_NAME, COV_CAT = f"{AC}!$C${CV0}:$C${CV1}", f"{AC}!$D${CV0}:$D${CV1}"

sec(244, "19. Whole-row highlight and supply windows",
    "A row is highlighted when the contract needs YOUR action. Renewals in process with supply covered are not highlighted.")
hdr(246, {"B": "#", "C": "Setting", "D": "Value", "I": "What it does"})
for i, (k, v, n) in enumerate([
        ("'Act now' (red row) = Priority at or below", 2, "Also: supply not covered, or lapsed with supply at risk."),
        ("'Plan' (amber row) = Priority at or below", 3, "Also: contract ending inside the supply window with no coverage."),
        ("Plan supply coverage for contracts ending within (days)", 90, "Before a contract lapses, Supply Status asks for a plan ('Not yet planned')."),
        ("(not used)", None, "Coverage Through was removed with the Coverage PO / Note column (11/2026).")]):
    r = 247 + i
    ctl[f"B{r}"] = i + 1; ctl[f"B{r}"].font = font(); ctl[f"B{r}"].border = BOX
    ctl[f"C{r}"] = k; ctl[f"C{r}"].font = font(); ctl[f"C{r}"].border = BOX; ctl[f"C{r}"].alignment = Alignment(wrap_text=True)
    if v is None:
        ctl[f"D{r}"].border = BOX; note(f"I{r}", n); ctl.merge_cells(f"E{r}:H{r}"); continue
    inp(f"D{r}", v, "0"); ctl[f"D{r}"].fill = fill("FFFF00"); note(f"I{r}", n); ctl.merge_cells(f"E{r}:H{r}")
ACT_N, PLAN_N, PLAN_WIN = (f"{AC}!$D$247", f"{AC}!$D$248", f"{AC}!$D$249")
INPROC, NORMAL = f"{AC}!$K$29", f"{AC}!$K$28"

# ---------------- Portfolio Tracker columns BA-BI
SUP = [  # col, header, kind, width, formula
    ("BA", "Burn Index", "auto", 9,
     lambda r: f'=IF(AND(ISNUMBER(Z{r}),ISNUMBER(AA{r}),N(X{r})>0),IF(AA{r}>={MIN_EL},Z{r}/AA{r},""),"")'),
    ("BB", "Projected Exhaustion Date", "auto", 12,
     lambda r: f'=IF(OR(BA{r}="",NOT(ISNUMBER(C{r}))),"",C{r}+(TODAY()-C{r})/Z{r})'),
    ("BC", "Funds Run Out Before End?", "auto", 14,
     lambda r: f'=IF(OR(BB{r}="",NOT(ISNUMBER(D{r}))),"",IF(BB{r}<D{r},"Yes - "&TEXT(D{r}-BB{r},"0")&" days early","No"))'),
    ("BD", "Supply Coverage", "in", 20, None),
    ("BE", "Coverage PO / Note", "in", 24, None),
    ("BF", "Coverage Through", "in", 11, None),
    ("BG", "Supply Status", "auto", 20,
     lambda r: (f'=IF(A{r}="","",IF(BD{r}<>"",BD{r},'
                f'IF(AH{r}={CLOSED_CAT},"",IF(AND(ISNUMBER(AD{r}),AD{r}<0),"Unknown - confirm",'
                f'IF(AND(ISNUMBER(AD{r}),AD{r}<={PLAN_WIN}),"Not yet planned","Active contract")))))')),
    ("BH", "Supply Risk", "auto", 11,
     lambda r: (f'=IF(A{r}="","",IF(BD{r}<>"",'
                f'IFERROR(INDEX({COV_CAT},MATCH(BD{r},{COV_NAME},0)),"At risk"),'
                f'IF(AH{r}={CLOSED_CAT},"Closed",IF(AND(ISNUMBER(AD{r}),AD{r}<={PLAN_WIN}),"At risk","Active"))))')),
    ("BI", "Action Needed", "auto", 11, None),
]
def _action(r):
    ok_ip = f'AND(AH{r}={INPROC},OR(BH{r}="Covered",BH{r}="Active"))'
    near = f'AND(ISNUMBER(AD{r}),AD{r}<={PLAN_WIN})'
    needs = (f'OR(AND({near},OR(BH{r}="At risk",BH{r}="Not covered")),'
             f'AND(ISNUMBER(AL{r}),AL{r}<={PLAN_N},NOT({ok_ip})))')
    act = (f'OR(BH{r}="Not covered",AND(ISNUMBER(AL{r}),AL{r}<={ACT_N},NOT({ok_ip})),'
           f'AND(ISNUMBER(AD{r}),AD{r}<0,BH{r}="At risk"))')
    return f'=IF(OR(A{r}="",AH{r}={CLOSED_CAT}),"",IF(NOT({needs}),"",IF({act},"Act now","Plan")))'
SUP[-1] = ("BI", "Action Needed", "auto", 11, _action)

for col, h, kind, w, fn in SUP:
    c = new[f"{col}1"]; c.value = h + (" (auto)" if kind == "auto" else "")
    c.font = font(bold=True, color="FFFFFF"); c.fill = fill(GREY_HDR if kind == "auto" else BLUE_HDR); c.border = BOX
    c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    new.column_dimensions[col].width = w
    for r in range(2, LAST + 1):
        x = new[f"{col}{r}"]; x.border = BOX; x.font = font()
        x.alignment = Alignment(vertical="top", horizontal="center" if kind == "auto" else "left", wrap_text=(col in ("BE", "BG")))
        if fn:
            x.value = fn(r); x.fill = fill("F2F2F2")
        if col in ("BB", "BF"): x.number_format = "mm/dd/yyyy"
        if col == "BA": x.number_format = "0.00"
    if col in ("BH", "BI"):
        for r in range(2, LAST + 1): new[f"{col}{r}"].font = font(bold=True)
dv(new, "=" + COV_NAME, f"BD2:BD{LAST}")
# Status & Notes (BJ): the one notes column. Filled by the build from the old note columns; finalize_layout moves it to
# right after the lane columns.
c = new["BJ1"]; c.value = "Status & Notes"
c.font = font(bold=True, color="FFFFFF"); c.fill = fill(BLUE_HDR); c.border = BOX
c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
new.column_dimensions["BJ"].width = 60
for r in range(2, LAST + 1):
    x = new[f"BJ{r}"]; x.border = BOX; x.font = font(); x.alignment = Alignment(vertical="top", wrap_text=True)
# cell colours
new.conditional_formatting.add(f"BA2:BA{LAST}", FormulaRule(formula=[f'AND(ISNUMBER($BA2),$BA2>={BURN_RED})'], fill=fill("F4A6A6"), font=Font(color="7A0000", bold=True), stopIfTrue=True))
new.conditional_formatting.add(f"BA2:BA{LAST}", FormulaRule(formula=[f'AND(ISNUMBER($BA2),$BA2>={BURN_AMB})'], fill=fill("FFD966"), font=Font(color="5C4400", bold=True), stopIfTrue=True))
new.conditional_formatting.add(f"BC2:BC{LAST}", FormulaRule(formula=['LEFT($BC2,3)="Yes"'], fill=fill("F4A6A6"), font=Font(color="7A0000", bold=True), stopIfTrue=True))
for val, fc, tc in [("Covered", "C6EFCE", "1E5631"), ("At risk", "FFD966", "5C4400"), ("Not covered", "C00000", "FFFFFF")]:
    new.conditional_formatting.add(f"BH2:BH{LAST}", FormulaRule(formula=[f'$BH2="{val}"'], fill=fill(fc), font=Font(color=tc, bold=True), stopIfTrue=True))
new.conditional_formatting.add(f"BI2:BI{LAST}", FormulaRule(formula=['$BI2="Act now"'], fill=fill("C00000"), font=Font(color="FFFFFF", bold=True), stopIfTrue=True))
new.conditional_formatting.add(f"BI2:BI{LAST}", FormulaRule(formula=['$BI2="Plan"'], fill=fill("FFD966"), font=Font(color="5C4400", bold=True), stopIfTrue=True))
# whole-row highlight: added last, so the status cells keep their own stronger colours
new.conditional_formatting.add(f"A2:BJ{LAST}", FormulaRule(formula=['$BI2="Act now"'], fill=fill("F8D7D7")))
new.conditional_formatting.add(f"A2:BJ{LAST}", FormulaRule(formula=['$BI2="Plan"'], fill=fill("FFF2D0")))
new.auto_filter.ref = f"A1:BJ{LAST}"

# ---------------- Dashboard: lapsed renewal x supply (rows 9-19)
LAP = f'{PT("AD")},"<0",{PT("AH")},"<>"&{CLOSED_CAT}'
sh("B9", "Lapsed contracts: renewal and supply")
tile("B", 10, "LAPSED", f'=COUNTIFS({LAP})', "Past end date, not closed", alert=("FBE3E3", "C00000"))
tile("E", 10, "LAPSED - RENEWAL IN PROCESS", f'=COUNTIFS({PT("AD")},"<0",{PT("AH")},{INPROC})', "Renewal submitted and moving")
tile("H", 10, "LAPSED - NO ACTION YET", f'=COUNTIFS({PT("AD")},"<0",{PT("AH")},{NORMAL})', "Nothing started", alert=("FBE3E3", "C00000"))
tile("K", 10, "% OF LAPSED WITH MOVEMENT", '=IFERROR(E11/B11,0)', "In process / lapsed", "0%")
tile("N", 10, "LAPSED - SUPPLY COVERED", f'=COUNTIFS({LAP},{PT("BH")},"Covered")', "DWM can still buy")
tile("Q", 10, "LAPSED - SUPPLY AT RISK", f'=COUNTIFS({LAP},{PT("BH")},"At risk")+COUNTIFS({LAP},{PT("BH")},"Not covered")',
     "Unknown, unconfirmed or not covered", alert=("FBE3E3", "C00000"))
db["B14"] = "Renewal x supply (lapsed contracts)"; db["B14"].font = font(bold=True, color=NAVY)
heads = [("C", "Supply covered", "C6EFCE", "1E5631"), ("D", "At risk / unknown", "FFD966", "5C4400"),
         ("E", "Not covered", "C00000", "FFFFFF"), ("F", "Total", NAVY, "FFFFFF")]
db["B15"].fill = fill(NAVY); db["B15"].border = BOX
for col, h, fc, tc in heads:
    c = db[f"{col}15"]; c.value = h; c.fill = fill(fc); c.font = font(bold=True, color=tc, size=9); c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
for r, (lab, cat) in zip((16, 17), [("In process", INPROC), ("No action", NORMAL)]):
    db[f"B{r}"] = lab; db[f"B{r}"].font = font(bold=True); db[f"B{r}"].border = BOX
    for col, risk in (("C", '"Covered"'), ("D", '"At risk"'), ("E", '"Not covered"')):
        db[f"{col}{r}"] = f'=COUNTIFS({PT("AD")},"<0",{PT("AH")},{cat},{PT("BH")},{risk})'
    db[f"F{r}"] = f"=SUM(C{r}:E{r})"
    for col in "CDEF":
        db[f"{col}{r}"].border = BOX; db[f"{col}{r}"].font = font(bold=True); db[f"{col}{r}"].alignment = Alignment(horizontal="center")
db.conditional_formatting.add("E16:E17", FormulaRule(formula=['E16>0'], fill=fill("F8D7D7")))
db.conditional_formatting.add("C17:E17", FormulaRule(formula=['C17>0'], fill=fill("FFF2D0")))
tile("H", 14, "LAPSED 90+ DAYS, NO ACTION", f'=COUNTIFS({PT("AD")},"<=-90",{PT("AH")},{NORMAL})', "Escalate", alert=("FBE3E3", "C00000"))
tile("K", 14, "AVERAGE DAYS LAPSED", f'=IFERROR(-AVERAGEIFS({PT("AD")},{LAP}),0)', "Across lapsed contracts", "0")
db.merge_cells("B18:S18")
db["B18"] = ('=B11&" contracts are past their end date. "&E11&" have renewals in process ("&TEXT(K11,"0%")&"), and DWM can still buy on "'
             '&N11&" of them. "&Q11&" need supply confirmed or arranged; "&H11&" have no action yet."')
db["B18"].font = font(bold=True, color=NAVY, size=11); db["B18"].alignment = Alignment(wrap_text=True, vertical="center")
db.row_dimensions[18].height = 44
lc = BarChart(); lc.type = "bar"; lc.grouping = "stacked"; lc.overlap = 100; lc.title = "Lapsed: renewal x supply"
lc.add_data(Reference(db, min_col=3, max_col=5, min_row=15, max_row=17), titles_from_data=True)
lc.set_categories(Reference(db, min_col=2, min_row=16, max_row=17))
for srs, hx in zip(lc.series, ("6AA84F", "F1C232", "C00000")):
    srs.graphicalProperties.solidFill = hx; srs.graphicalProperties.line.solidFill = hx
lc.y_axis.majorGridlines = None; lc.x_axis.scaling.orientation = "maxMin"
lc.height, lc.width = 4.6, 11
db.add_chart(lc, "N14")

# ---------------- Summary: lapsed renewal x supply
sr0 = sm.max_row + 2
shead(f"A{sr0}", "Lapsed contracts: renewal and supply")
th(f"A{sr0+1}", "Lapsed contracts"); th(f"B{sr0+1}", "Supply covered"); th(f"C{sr0+1}", "At risk / unknown"); th(f"D{sr0+1}", "Not covered"); th(f"E{sr0+1}", "Total")
for i, (lab, cat) in enumerate([("Renewal in process", INPROC), ("No action yet", NORMAL)]):
    r = sr0 + 2 + i
    td(f"A{r}", lab, "@", bold=True)
    for col, risk in (("B", '"Covered"'), ("C", '"At risk"'), ("D", '"Not covered"')):
        td(f"{col}{r}", f'=COUNTIFS({PT("AD")},"<0",{PT("AH")},{cat},{PT("BH")},{risk})')
    td(f"E{r}", f"=SUM(B{r}:D{r})", bold=True)
r = sr0 + 4
td(f"A{r}", "% of lapsed with movement", "@"); td(f"B{r}", f"=IFERROR(E{sr0+2}/(E{sr0+2}+E{sr0+3}),0)", "0%", bold=True)
td(f"A{r+1}", "Lapsed 90+ days, no action", "@"); td(f"B{r+1}", f'=COUNTIFS({PT("AD")},"<=-90",{PT("AH")},{NORMAL})', bold=True)
td(f"A{r+2}", "Average days lapsed", "@"); td(f"B{r+2}", f'=IFERROR(-AVERAGEIFS({PT("AD")},{LAP}),0)', "0", bold=True)
td(f"A{r+3}", "Funds run out before end date", "@"); td(f"B{r+3}", f'=COUNTIF({PT("BC")},"Yes*")', bold=True)
