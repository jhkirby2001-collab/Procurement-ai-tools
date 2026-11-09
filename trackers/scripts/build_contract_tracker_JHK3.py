"""Build the Contract Tracker workbook.

Usage: python build_contract_tracker_JHK3.py ../data/Contract_List_Source_JHK3.xlsx ../outputs/Contract_Tracker_2026_JHK3.xlsx \
           ../data/Shared_Contracts_BWS_Source_JHK3.xlsx
The third argument (shared contracts list) and fourth (DPS 365 expiring-contracts report) are optional.
To keep a working copy's typed entries: CARRY_USER=<working copy> carries every input cell from it into the new build.
If the working copy has the Reviewed column in A: add REVIEWED_COL=1. Then: recalc.py <out>, then finalize_layout_JHK3.py <out>.
Then recalculate (open and save in Excel, or LibreOffice headless) so formula values are cached.
Optional env: TEST_TO=1 loads sample task-order releases; RENDER_DASH=<sheet> hides every other sheet (for rendering checks).
"""
import sys, os
from copy import copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI

SRC, OUT = sys.argv[1], sys.argv[2]
F = "Arial"
NAVY, BLUE_HDR, GREY_HDR = "1F3864", "2E75B6", "595959"
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
LAST = 300                      # tracker formulas / dropdowns / colours run to this row
AC = "'Alert Controls'"

def font(**k):
    k.setdefault("name", F); k.setdefault("size", 10); return Font(**k)
def fill(c): return PatternFill("solid", start_color=c, end_color=c)

# Colour palette: name -> (fill, text). Rules pick a colour by name from the Alert Controls tab.
PALETTE = [("Dark Red", "C00000", "FFFFFF"), ("Red", "F4A6A6", "7A0000"), ("Amber", "FFD966", "5C4400"),
           ("Light Amber", "FFF2CC", "5C4400"), ("Green", "C6EFCE", "1E5631"), ("Blue", "BDD7EE", "1F3864"),
           ("Grey", "D9D9D9", "404040"), ("Purple", "E1D5F0", "4B2C73")]

wb = openpyxl.load_workbook(SRC)
old = wb["Portfolio Tracker"]

# =====================================================================================
# 1. ALERT CONTROLS
# =====================================================================================
ctl = wb.create_sheet("Alert Controls")
ctl.sheet_view.showGridLines = False
for c, w in zip("ABCDEFGHIJK", [2, 8, 30, 24, 18, 18, 14, 12, 70, 3, 22]):
    ctl.column_dimensions[c].width = w
ctl["B1"] = "Alert Controls"; ctl["B1"].font = font(size=16, bold=True, color=NAVY)
ctl["B2"] = ("Every number, list and rule the alerts use lives on this tab. Change a blue cell or a dropdown and the "
             "Portfolio Tracker and Summary update. See the How To Use tab before changing rules.")
ctl["B2"].font = font(size=9, italic=True, color=GREY_HDR)

def sec(row, text, note=None):
    c = ctl[f"B{row}"]; c.value = text; c.font = font(size=12, bold=True, color="FFFFFF"); c.fill = fill(NAVY)
    for col in "CDEFGHI": ctl[f"{col}{row}"].fill = fill(NAVY)
    if note:
        ctl[f"B{row+1}"] = note; ctl[f"B{row+1}"].font = font(size=9, italic=True, color=GREY_HDR)

def hdr(row, cols):
    for col, text in cols.items():
        c = ctl[f"{col}{row}"]; c.value = text; c.font = font(bold=True, color="FFFFFF"); c.fill = fill(BLUE_HDR)
        c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = BOX
    ctl.row_dimensions[row].height = 30

def inp(cell, value, fmt=None, wrap=False):
    c = ctl[cell]; c.value = value; c.font = font(color="0000FF"); c.border = BOX
    c.alignment = Alignment(wrap_text=wrap, vertical="top")
    if fmt: c.number_format = fmt

def note(cell, value):
    c = ctl[cell]; c.value = value; c.font = font(size=9, color="404040"); c.border = BOX
    c.alignment = Alignment(wrap_text=True, vertical="top")

def dv(ws, ref, rng, msg="Pick from the list. To add choices, add them on the Alert Controls tab."):
    v = DataValidation(type="list", formula1=ref, allow_blank=True, showErrorMessage=True,
                       errorStyle="warning", error=msg)
    ws.add_data_validation(v); v.add(rng)

# --- 1. General
sec(4, "1. General settings")
hdr(6, {"B": "#", "C": "Setting", "D": "Value", "I": "What it does"})
GEN = [("Default lead time when no Renewal Path is chosen (days)", 540, None,
        "Used until a Renewal Path is picked. Set to the longest path so nothing is under-warned."),
       ("Value that marks a Next Action as done", "Yes", None,
        "When the 'Next Action Done?' column equals this, the Next Action flag clears.")]
for i, (k, v, f, n) in enumerate(GEN):
    r = 7 + i
    ctl[f"B{r}"] = i + 1; ctl[f"B{r}"].font = font(); ctl[f"B{r}"].border = BOX
    ctl[f"C{r}"] = k; ctl[f"C{r}"].font = font(); ctl[f"C{r}"].border = BOX; ctl[f"C{r}"].alignment = Alignment(wrap_text=True)
    inp(f"D{r}", v, f); ctl[f"D{r}"].fill = fill("FFFF00"); note(f"I{r}", n)
    ctl.merge_cells(f"E{r}:H{r}")
DEFAULT_LEAD, DONE_VAL = f"{AC}!$D$7", f"{AC}!$D$8"

# --- 2. Renewal paths
sec(10, "2. Renewal Paths  (dropdown on the tracker)", "Rename, add or remove paths in the 10 slots. Lead time = days before the end date that work must start.")
hdr(12, {"B": "Slot", "C": "Renewal Path", "D": "Lead Time (days)", "I": "What it means / when to use it"})
PATHS = [("Extension available", 240, "An option or time extension is left. Lead time covers the extension request through DPS approval."),
         ("Competitive rebid", 540, "No options left; a new solicitation is needed. Covers usage estimate, spec and M/WBE goal setting, DPS solicitation, award, plus cushion."),
         ("Sole source / NCRB", 365, "Non-competitive review. Covers the NCRB package, market research and price reasonableness review."),
         ("Let expire", 0, "No replacement planned. Alert at the end date only; set Renewal Stage to a Closed stage once confirmed."),
         ("Owning dept renews", 365, "Shared contract: another department owns it and runs the renewal. Lead time is when your department's usage estimate and spec input are due.")]
P0, P1 = 13, 22
for i in range(10):
    r = P0 + i
    ctl[f"B{r}"] = i + 1; ctl[f"B{r}"].font = font(color=GREY_HDR); ctl[f"B{r}"].border = BOX
    name, lead, n = PATHS[i] if i < len(PATHS) else (None, None, None)
    inp(f"C{r}", name); inp(f"D{r}", lead, "0"); note(f"I{r}", n)
    ctl.merge_cells(f"E{r}:H{r}")
PATH_NAME, PATH_LEAD = f"{AC}!$C${P0}:$C${P1}", f"{AC}!$D${P0}:$D${P1}"

# --- 3. Action statuses + categories
sec(24, "3. Renewal Stage  (dropdown on the tracker; formerly Action Status)", "Each status 'counts as' a category. The rules in section 4 use the category, so new statuses work as soon as you give them one.")
hdr(26, {"B": "Slot", "C": "Renewal Stage", "D": "Counts As (category)", "I": "What it means", "K": "Categories (editable)"})
STAT = [("Not started", "Normal", "Nothing submitted yet."),
        ("Req drafted", "Normal", "Requisition or package being prepared by our team. Still in our hands."),
        ("Submitted to DPS", "In Process", "Package is with DPS. Clock is running on their side."),
        ("In legal review", "In Process", "With Law / DPS legal."),
        ("Advertised / bid open", "In Process", "Solicitation is out."),
        ("Awarded - pending execution", "In Process", "Award made; contract not yet executed."),
        ("Replaced / executed", "Closed", "New contract or extension is in place. This row no longer needs alerts."),
        ("Let expire", "Closed", "Bureau confirmed no replacement is needed."),
        ("Voided / cancelled", "Closed", "PO voided or requirement cancelled."),
        ("Usage estimate sent to owning dept", "In Process", "Shared contract: your input has gone to the department that owns the contract."),
        ("Waiting on owning dept", "In Process", "Shared contract: the owning department is running the renewal."),
        ("DWM included in new contract", "Closed", "Shared contract: the replacement is in place and covers your department.")]
S0, S1 = 27, 38
for i in range(12):
    r = S0 + i
    ctl[f"B{r}"] = i + 1; ctl[f"B{r}"].font = font(color=GREY_HDR); ctl[f"B{r}"].border = BOX
    s, cat, n = STAT[i] if i < len(STAT) else (None, None, None)
    inp(f"C{r}", s); inp(f"D{r}", cat); note(f"I{r}", n)
    ctl.merge_cells(f"E{r}:H{r}")
STAT_NAME, STAT_CAT = f"{AC}!$C${S0}:$C${S1}", f"{AC}!$D${S0}:$D${S1}"
# categories list: K27 = "Any" (fixed), K28:K35 editable categories
ctl["K27"] = "Any"; ctl["K27"].font = font(bold=True); ctl["K27"].border = BOX
for i, c in enumerate(["Normal", "In Process", "Closed"] + [None] * 5):
    inp(f"K{28+i}", c)
ctl["K36"] = "'Any' is fixed. Add categories below Closed."; ctl["K36"].font = font(size=8, italic=True, color=GREY_HDR)
ctl["K36"].alignment = Alignment(wrap_text=True)
CAT_LIST, APPLIES_LIST = f"{AC}!$K$28:$K$35", f"{AC}!$K$27:$K$35"
dv(ctl, "=" + CAT_LIST, f"D{S0}:D{S1}")

COLOR_REF = f"{AC}!$C$84:$C$91"   # set in section 7

# --- 4. Expiration rules
sec(40, "4. Expiration Status rules  (time left on the contract)",
    "Checked top to bottom; the FIRST rule that matches wins. A rule matches when the category fits AND (days to expiration <= its limit OR days to Start-By <= its limit). Both limits blank = always matches.")
hdr(42, {"B": "Order", "C": "Status Label", "D": "Applies when category is", "E": "Days to expiration at or under",
         "F": "Days to Start-By date at or under", "G": "Colour", "H": "Severity (1 = most urgent)", "I": "What it means / why"})
EXP = [("CLOSED", "Closed", None, None, "Grey", 9, "Renewal Stage counts as Closed. Kept for the record; excluded from the to-do list."),
       ("LAPSED", "Any", -1, None, "Dark Red", 1, "Past the end date and not closed. Department may be buying without a contract. Top priority."),
       ("IN PROCESS", "In Process", None, None, "Blue", 5, "Renewal work is submitted and moving. Watch it, but no action from you right now."),
       ("RED", "Any", 180, 0, "Red", 2, "Six months or less left, OR the Start-By date has passed with no action submitted. Act now."),
       ("YELLOW", "Any", 365, 90, "Amber", 3, "A year or less left, OR the Start-By date is within 90 days. Start preparing."),
       ("GREEN", "Any", None, None, "Green", 8, "More runway than the rules above. No action needed yet.")]
E0, E1 = 43, 52
for i in range(10):
    r = E0 + i
    ctl[f"B{r}"] = i + 1; ctl[f"B{r}"].font = font(color=GREY_HDR); ctl[f"B{r}"].border = BOX
    lab, app, d1, d2, col, sev, n = EXP[i] if i < len(EXP) else (None,) * 7
    inp(f"C{r}", lab); inp(f"D{r}", app); inp(f"E{r}", d1, "0"); inp(f"F{r}", d2, "0"); inp(f"G{r}", col); inp(f"H{r}", sev, "0"); note(f"I{r}", n)
dv(ctl, "=" + APPLIES_LIST, f"D{E0}:D{E1}")
dv(ctl, "=" + COLOR_REF, f"G{E0}:G{E1}")

# --- 5. Funds rules
sec(54, "5. Funds Flag rules  (money spent)",
    "First match wins. A rule matches when % expended >= its limit OR spending is ahead of time by >= its limit. Both blank = always matches. Needs Total Contract Value entered.")
hdr(56, {"B": "Order", "C": "Flag Label", "D": "% Expended at or over", "E": "Spending ahead of time by at least (% pts)",
         "G": "Colour", "H": "Severity (1 = most urgent)", "I": "What it means / why"})
FUN = [("RED", 0.75, 0.20, "Red", 2, "75%+ spent, or spending is 20+ points ahead of the time used. A funding increase should already be moving."),
       ("YELLOW", 0.50, 0.10, "Amber", 4, "50%+ spent, or 10+ points ahead of time. Check usage with the bureau and get ahead of it."),
       ("GREEN", None, None, "Green", 8, "Spending is in line with time.")]
U0, U1 = 57, 64
for i in range(8):
    r = U0 + i
    ctl[f"B{r}"] = i + 1; ctl[f"B{r}"].font = font(color=GREY_HDR); ctl[f"B{r}"].border = BOX
    lab, p, g, col, sev, n = FUN[i] if i < len(FUN) else (None,) * 6
    inp(f"C{r}", lab); inp(f"D{r}", p, "0%"); inp(f"E{r}", g, "0%"); ctl[f"F{r}"].border = BOX
    inp(f"G{r}", col); inp(f"H{r}", sev, "0"); note(f"I{r}", n)
dv(ctl, "=" + COLOR_REF, f"G{U0}:G{U1}")

# --- 6. Next action rules
sec(66, "6. Next Action Flag rules  (RETIRED 11/2026 - the Next Action columns were removed; nothing uses these rules)",
    "First match wins. Matches when days until the Next Action Due Date <= its limit (blank = always). No flag when there is no due date or the action is marked done.")
hdr(68, {"B": "Order", "C": "Flag Label", "E": "Days until due at or under", "G": "Colour", "H": "Severity (1 = most urgent)", "I": "What it means / why"})
NXT = [("OVERDUE", -1, "Red", 2, "Due date has passed and the action is not marked done."),
       ("DUE SOON", 14, "Amber", 3, "Due within two weeks."),
       ("UPCOMING", 30, "Light Amber", 6, "Due within 30 days. Put it on this month's plan.")]
N0, N1 = 69, 76
for i in range(8):
    r = N0 + i
    ctl[f"B{r}"] = i + 1; ctl[f"B{r}"].font = font(color=GREY_HDR); ctl[f"B{r}"].border = BOX
    lab, d, col, sev, n = NXT[i] if i < len(NXT) else (None,) * 5
    inp(f"C{r}", lab); ctl[f"D{r}"].border = BOX; inp(f"E{r}", d, "0"); ctl[f"F{r}"].border = BOX
    inp(f"G{r}", col); inp(f"H{r}", sev, "0"); note(f"I{r}", n)
dv(ctl, "=" + COLOR_REF, f"G{N0}:G{N1}")

# --- 7. Colour key
sec(80, "7. Colour key", "Pick these names in the Colour column of sections 4-6. Text always shows the status word too, so colour is never the only signal.")
hdr(83, {"B": "#", "C": "Colour name", "D": "Sample", "I": "Suggested use"})
USE = {"Dark Red": "Past due / lapsed. Reserve for the single most urgent state.", "Red": "Act now.",
       "Amber": "Prepare / watch closely.", "Light Amber": "Coming up; plan it.", "Green": "On track.",
       "Blue": "In process with someone else.", "Grey": "Closed / inactive.", "Purple": "Spare for a new status your team adds."}
for i, (nm, fc, tc) in enumerate(PALETTE):
    r = 84 + i
    ctl[f"B{r}"] = i + 1; ctl[f"B{r}"].font = font(color=GREY_HDR); ctl[f"B{r}"].border = BOX
    ctl[f"C{r}"] = nm; ctl[f"C{r}"].font = font(bold=True); ctl[f"C{r}"].border = BOX
    ctl[f"D{r}"] = "STATUS"; ctl[f"D{r}"].fill = fill(fc); ctl[f"D{r}"].font = font(bold=True, color=tc)
    ctl[f"D{r}"].alignment = Alignment(horizontal="center"); ctl[f"D{r}"].border = BOX
    note(f"I{r}", USE[nm])

# --- 8. Summary lists
sec(94, "8. Summary groupings", "The Summary tab breaks counts down by these lanes and owners. Edit to match your team.")
hdr(96, {"C": "Lanes", "D": "Owners"})
lanes = ["Chemical", "Meters", "Pipes", "Valves", "Water/Sewer", "Facility", "No lane marked"]
owners = sorted({str(old[f"Q{r}"].value).strip() for r in range(2, 88) if old[f"Q{r}"].value})
for i in range(14):
    inp(f"C{97+i}", lanes[i] if i < len(lanes) else None)
    inp(f"D{97+i}", owners[i] if i < len(owners) else None)
LANES0, OWN0, NGRP = 97, 97, 14

# colour previews on the rule tables
def colour_cf(ws, rng, test):
    for nm, fc, tc in PALETTE:
        ws.conditional_formatting.add(rng, FormulaRule(formula=[test.replace("@", nm)],
                                      fill=fill(fc), font=Font(color=tc, bold=True), stopIfTrue=True))
for a, b in [(E0, E1), (U0, U1), (N0, N1)]:
    colour_cf(ctl, f"C{a}:C{b}", f'$G{a}="@"')
    colour_cf(ctl, f"G{a}:G{b}", f'$G{a}="@"')
ctl.freeze_panes = "A4"

# =====================================================================================
# 2. PORTFOLIO TRACKER (rebuilt; columns A-Q and notes copied unchanged)
# =====================================================================================
new = wb.create_sheet("Portfolio Tracker NEW")
COPY = {c: c for c in "ABCDEFGHIJKLMNOPQ"}
COPY.update({"R": "W", "S": "X", "Y": "AM", "Z": "AN", "AA": "AO", "AB": "AP", "AC": "AQ"})

for r in range(1, 88):
    for src, dst in COPY.items():
        s = old[f"{src}{r}"]; d = new[f"{dst}{r}"]
        v = s.value
        if isinstance(v, str) and v.startswith("="):   # broken stray formula (old S7) -> leave input blank
            v = None
        d.value = v
        if s.has_style:
            d.font = copy(s.font); d.fill = copy(s.fill); d.border = copy(s.border)
            d.alignment = copy(s.alignment); d.number_format = s.number_format
new["AP1"] = "Additional Req No."   # was an unlabeled column
for src, dst in COPY.items():
    w = old.column_dimensions[src].width if src in old.column_dimensions else None
    if w: new.column_dimensions[dst].width = w
    if src in old.column_dimensions and old.column_dimensions[src].hidden:
        new.column_dimensions[dst].hidden = True

exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "contract_tracker_shared_part_JHK3.py")).read())

NEWCOLS = [  # (col, header, kind, width)
    ("R", "Renewal Path", "in", 22), ("S", "Renewal Stage", "in", 22), ("T", "Next Action", "in", 32),
    ("U", "Next Action Due Date", "in", 12), ("V", "Next Action Done?", "in", 9),
    ("W", "Total Contract Value ($)", "in", 15), ("X", "Expended to Date ($)", "in", 15),
    ("Y", "Remaining Balance ($)", "auto", 15), ("Z", "% Expended", "auto", 10), ("AA", "% of Term Used", "auto", 10),
    ("AB", "Spending Ahead of Time (pts)", "auto", 11), ("AC", "Funds Flag", "auto", 11),
    ("AD", "Days Until Expiration", "auto", 11), ("AE", "Lead Time (days)", "auto", 9), ("AF", "Start-By Date", "auto", 12),
    ("AG", "Days Until Start-By", "auto", 10), ("AH", "Action Category", "auto", 12), ("AI", "Expiration Status", "auto", 13),
    ("AJ", "Days Until Next Action", "auto", 10), ("AK", "Next Action Flag", "auto", 12), ("AL", "Priority (1 = most urgent)", "auto", 10)]
for col, h, kind, w in NEWCOLS:
    c = new[f"{col}1"]; c.value = h + (" (auto)" if kind == "auto" else "")
    c.font = font(bold=True, color="FFFFFF"); c.fill = fill(BLUE_HDR if kind == "in" else GREY_HDR)
    c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center"); c.border = BOX
    new.column_dimensions[col].width = w
for col in ["AM", "AN", "AO", "AP", "AQ"]:
    new[f"{col}1"].font = font(bold=True, color="FFFFFF"); new[f"{col}1"].fill = fill(BLUE_HDR)
    new[f"{col}1"].alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
new.column_dimensions["AM"].width = 30; new.column_dimensions["AN"].width = 45; new.column_dimensions["AO"].width = 18
new.column_dimensions["AP"].width = 12; new.column_dimensions["AQ"].width = 30
new.row_dimensions[1].height = 45

def slot_chain(slots, cond, label_col, fallback='""'):
    expr = fallback
    for r in reversed(slots):
        expr = f'IF({cond(r)},{AC}!${label_col}${r},{expr})'
    return expr

def exp_cond(rr):
    def c(s):
        E, Fd = f"{AC}!$E${s}", f"{AC}!$F${s}"
        return (f'AND({AC}!$C${s}<>"",OR({AC}!$D${s}="Any",{AC}!$D${s}=AH{rr}),'
                f'IF(AND({E}="",{Fd}=""),TRUE,OR(AND({E}<>"",AD{rr}<={E}),AND({Fd}<>"",AG{rr}<={Fd}))))')
    return c

def fund_cond(rr):
    def c(s):
        P, G = f"{AC}!$D${s}", f"{AC}!$E${s}"
        return (f'AND({AC}!$C${s}<>"",IF(AND({P}="",{G}=""),TRUE,'
                f'OR(AND({P}<>"",Z{rr}>={P}),AND({G}<>"",ISNUMBER(AB{rr}),AB{rr}>={G}))))')
    return c

def next_cond(rr):
    def c(s):
        D = f"{AC}!$E${s}"
        return f'AND({AC}!$C${s}<>"",IF({D}="",TRUE,AJ{rr}<={D}))'
    return c

def sev(cell, a, b):
    return f'IFERROR(INDEX({AC}!$H${a}:$H${b},MATCH({cell},{AC}!$C${a}:$C${b},0)),99)'

for r in range(2, LAST + 1):
    f = {}
    f["Y"] = f'=IF(W{r}="","",W{r}-X{r})'
    f["Z"] = f'=IF(OR(W{r}="",W{r}=0),"",N(X{r})/W{r})'
    f["AA"] = (f'=IF(D{r}="","",IF(NOT(ISNUMBER(C{r})),"Check start date",'
               f'IFERROR(MAX(0,MIN(1,(TODAY()-C{r})/(D{r}-C{r}))),"")))')
    f["AB"] = f'=IF(AND(ISNUMBER(Z{r}),ISNUMBER(AA{r})),Z{r}-AA{r},"")'
    f["AC"] = f'=IF(Z{r}="","",' + slot_chain(range(U0, U1 + 1), fund_cond(r), "C") + ")"
    f["AD"] = f'=IF(D{r}="","",D{r}-TODAY())'
    f["AE"] = (f'=IF(D{r}="","",IF(R{r}="",{DEFAULT_LEAD},'
               f'IFERROR(INDEX({PATH_LEAD},MATCH(R{r},{PATH_NAME},0)),{DEFAULT_LEAD})))')
    f["AF"] = f'=IF(AE{r}="","",D{r}-AE{r})'
    f["AG"] = f'=IF(AF{r}="","",AF{r}-TODAY())'
    f["AH"] = f'=IF(D{r}="","",IF(S{r}="","Normal",IFERROR(INDEX({STAT_CAT},MATCH(S{r},{STAT_NAME},0)),"Normal")))'
    f["AI"] = f'=IF(D{r}="","",' + slot_chain(range(E0, E1 + 1), exp_cond(r), "C") + ")"
    f["AJ"] = f'=IF(OR(NOT(ISNUMBER(U{r})),V{r}={DONE_VAL}),"",U{r}-TODAY())'
    f["AK"] = f'=IF(AJ{r}="","",' + slot_chain(range(N0, N1 + 1), next_cond(r), "C") + ")"
    f["AL"] = (f'=IF(A{r}="","",IF(MIN({sev(f"AI{r}",E0,E1)},{sev(f"AC{r}",U0,U1)})=99,"",'
               f'MIN({sev(f"AI{r}",E0,E1)},{sev(f"AC{r}",U0,U1)})))')
    for col, frm in f.items():
        c = new[f"{col}{r}"]; c.value = frm
        c.font = font(); c.fill = fill("F2F2F2"); c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="top")
    for col in ["R", "S", "T", "U", "V", "W", "X"]:
        c = new[f"{col}{r}"]; c.border = BOX; c.font = font(); c.alignment = Alignment(vertical="top", wrap_text=(col == "T"))
    new[f"U{r}"].number_format = "mm/dd/yyyy"; new[f"AF{r}"].number_format = "mm/dd/yyyy"
    for col in ["W", "X", "Y"]: new[f"{col}{r}"].number_format = '$#,##0;($#,##0);"-"'
    for col in ["Z", "AA"]: new[f"{col}{r}"].number_format = "0%"
    new[f"AB{r}"].number_format = '+0%;-0%;0%'
    for col in ["AD", "AE", "AG", "AJ", "AL"]: new[f"{col}{r}"].number_format = "0"
    for col in ["AI", "AC", "AK"]: new[f"{col}{r}"].font = font(bold=True)

dv(new, "=" + PATH_NAME, f"R2:R{LAST}")
dv(new, "=" + STAT_NAME, f"S2:S{LAST}")
dv(new, '"Yes,No"', f"V2:V{LAST}")
v = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True, showErrorMessage=True, error="Enter a date (mm/dd/yyyy).")
new.add_data_validation(v); v.add(f"U2:U{LAST}")
v = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True, showErrorMessage=True, error="Enter a dollar amount.")
new.add_data_validation(v); v.add(f"W2:X{LAST}")

def tracker_cf(col, a, b):
    rng = f"{col}2:{col}{LAST}"
    for nm, fc, tc in PALETTE:
        t = f'IFERROR(INDEX({AC}!$G${a}:$G${b},MATCH(${col}2,{AC}!$C${a}:$C${b},0)),"")="{nm}"'
        new.conditional_formatting.add(rng, FormulaRule(formula=[t], fill=fill(fc), font=Font(color=tc, bold=True), stopIfTrue=True))
tracker_cf("AI", E0, E1); tracker_cf("AC", U0, U1); tracker_cf("AK", N0, N1)
new.conditional_formatting.add(f"AL2:AL{LAST}", FormulaRule(formula=['AND($AL2<>"",$AL2<=1)'], fill=fill("C00000"), font=Font(color="FFFFFF", bold=True)))
new.conditional_formatting.add(f"AL2:AL{LAST}", FormulaRule(formula=['AND($AL2<>"",$AL2<=2)'], fill=fill("F4A6A6"), font=Font(color="7A0000", bold=True)))
new.conditional_formatting.add(f"AA2:AA{LAST}", FormulaRule(formula=['$AA2="Check start date"'], font=Font(color="C00000", italic=True)))

new.freeze_panes = "C2"
new.auto_filter.ref = f"A1:AQ{LAST}"
new.sheet_view.zoomScale = 90

# swap sheets
wb.remove(old)
new.title = "Portfolio Tracker"
wb.remove(wb["Settings"])

# =====================================================================================
# 3. SUMMARY (rebuilt; labels follow Alert Controls)
# =====================================================================================
wb.remove(wb["Summary"])
sm = wb.create_sheet("Summary")
sm.sheet_view.showGridLines = False
PT = lambda c: f"'Portfolio Tracker'!${c}$2:${c}${LAST}"
sm["A1"] = "Portfolio Summary"; sm["A1"].font = font(size=16, bold=True, color=NAVY)
sm["A2"] = '="As of "&TEXT(TODAY(),"mmmm d, yyyy")&"   |   Labels and colours come from the Alert Controls tab."'
sm["A2"].font = font(size=9, italic=True, color=GREY_HDR)
sm.column_dimensions["A"].width = 30
for i in range(2, 13): sm.column_dimensions[L(i)].width = 12

def shead(cell, text):
    sm[cell] = text; sm[cell].font = font(bold=True, size=12, color=NAVY)
def th(cell, text):
    c = sm[cell]; c.value = text; c.font = font(bold=True, color="FFFFFF"); c.fill = fill(BLUE_HDR); c.border = BOX
    c.alignment = Alignment(horizontal="center", wrap_text=True)
def td(cell, val, fmt="0", bold=False):
    c = sm[cell]; c.value = val; c.font = font(bold=bold); c.border = BOX; c.number_format = fmt

shead("A4", "At a glance")
th("A5", "Metric"); th("B5", "Count")
td("A6", "Contracts tracked"); td("B6", f'=COUNTA({PT("A")})', bold=True)
td("A7", "Need action now (priority 1-2)"); td("B7", f'=COUNTIF({PT("AL")},"<=2")', bold=True)
td("A8", "No Renewal Path chosen yet"); td("B8", f'=COUNTIFS({PT("A")},"<>",{PT("R")},"")')
td("A9", "No spending entered yet"); td("B9", f'=B6-COUNT({PT("X")})')
td("A10", "Start date needs checking"); td("B10", f'=COUNTIF({PT("AA")},"Check start date")')

def status_block(top, col0, title, a, b, rng, total_label=True):
    shead(f"{col0}{top}", title)
    c1 = L(CI(col0) + 1); c2 = L(CI(col0) + 2)
    th(f"{col0}{top+1}", "Status"); th(f"{c1}{top+1}", "Count"); th(f"{c2}{top+1}", "% of contracts")
    for i, s in enumerate(range(a, b + 1)):
        r = top + 2 + i
        td(f"{col0}{r}", f'=IF({AC}!$C${s}="","",{AC}!$C${s})', "@", bold=True)
        td(f"{c1}{r}", f'=IF({col0}{r}="","",COUNTIF({rng},{col0}{r}))')
        td(f"{c2}{r}", f'=IF({col0}{r}="","",IFERROR({c1}{r}/$B$6,0))', "0%")
        for nm, fc, tc in PALETTE:
            sm.conditional_formatting.add(f"{col0}{r}", FormulaRule(
                formula=[f'{AC}!$G${s}="{nm}"'], fill=fill(fc), font=Font(color=tc, bold=True), stopIfTrue=True))
    return top + 2 + (b - a + 1)

nxt = status_block(12, "A", "Expiration Status", E0, E1, PT("AI"))
nxt = status_block(nxt + 1, "A", "Funds Flag", U0, U1, PT("AC"))

def matrix(top, title, grp0, grp_rng):
    shead(f"A{top}", title)
    th(f"A{top+1}", "Group"); th(f"B{top+1}", "Contracts")
    for j, s in enumerate(range(E0, E1 + 1)):
        col = L(3 + j)
        c = sm[f"{col}{top+1}"]; c.value = f'=IF({AC}!$C${s}="","",{AC}!$C${s})'
        c.font = font(bold=True, color="FFFFFF"); c.fill = fill(BLUE_HDR); c.border = BOX
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        for nm, fc, tc in PALETTE:
            sm.conditional_formatting.add(f"{col}{top+1}", FormulaRule(
                formula=[f'{AC}!$G${s}="{nm}"'], fill=fill(fc), font=Font(color=tc, bold=True), stopIfTrue=True))
    for i in range(NGRP):
        r = top + 2 + i; src = f"{AC}!$C${grp0+i}" if grp_rng == "P" else f"{AC}!$D${grp0+i}"
        td(f"A{r}", f'=IF({src}="","",{src})', "@")
        td(f"B{r}", f'=IF(A{r}="","",COUNTIF({PT(grp_rng)},A{r}))', bold=True)
        for j in range(E1 - E0 + 1):
            col = L(3 + j)
            td(f"{col}{r}", f'=IF(OR($A{r}="",{col}${top+1}=""),"",COUNTIFS({PT(grp_rng)},$A{r},{PT("AI")},{col}${top+1}))')
    return top + 2 + NGRP

nxt = matrix(nxt + 1, "By Lane x Expiration Status", LANES0, "P")
nxt = matrix(nxt + 1, "By Owner x Expiration Status", OWN0, "Q")

shead(f"A{nxt+1}", "Money (fills in as dollar amounts are entered)")
r = nxt + 2
for lab, frm in [("Total contract value", f'=SUM({PT("W")})'), ("Expended to date", f'=SUM({PT("X")})'),
                 ("Remaining balance", f'=SUM({PT("Y")})')]:
    td(f"A{r}", lab); td(f"B{r}", frm, '$#,##0;($#,##0);"-"', bold=True); r += 1
sm.column_dimensions["B"].width = 16

# =====================================================================================
# 4. HOW TO USE
# =====================================================================================
hw = wb.create_sheet("How To Use")
hw.sheet_view.showGridLines = False
hw.column_dimensions["A"].width = 2; hw.column_dimensions["B"].width = 30; hw.column_dimensions["C"].width = 110
hw["B1"] = "How To Use This Tracker"; hw["B1"].font = font(size=16, bold=True, color=NAVY)
hw["B2"] = "Training guide for staff, and the reasoning behind the alert logic for leadership."
hw["B2"].font = font(size=9, italic=True, color=GREY_HDR)
row = [4]
def H(t):
    r = row[0] + 1
    c = hw[f"B{r}"]; c.value = t; c.font = font(bold=True, size=12, color="FFFFFF"); c.fill = fill(NAVY)
    hw[f"C{r}"].fill = fill(NAVY); row[0] = r + 1
def P(a, b="", bold_a=True):
    r = row[0]
    hw[f"B{r}"] = a; hw[f"C{r}"] = b
    hw[f"B{r}"].font = font(bold=bold_a); hw[f"C{r}"].font = font()
    for c in "BC":
        hw[f"{c}{r}"].alignment = Alignment(wrap_text=True, vertical="top"); hw[f"{c}{r}"].border = BOX
    row[0] = r + 1
def GAP(): row[0] += 1

H("1. What this workbook does")
P("Purpose", "Tracks every Goods & Services contract in the portfolio and raises three alerts: time left before the contract expires, how fast money is being spent, and whether the next action on the contract is due. The Priority column combines them so the most urgent contracts sort to the top.")
P("Who uses it", "Contract staff update it; the portfolio lead reviews it weekly; leadership sees the Summary tab.")

H("2. The tabs")
P("How To Use", "This guide.")
P("Column Guide", "Every column on the Portfolio Tracker and Task Order Tracker: who fills it in (original list, you, or a formula), what it means, what to enter or how it is figured, an example, and which Alert Controls section changes it.")
P("Dashboard", "One-page view for the portfolio owner and leadership: headline numbers, status mix, expirations by quarter, status by lane, workload by owner, and the 10 most urgent contracts. All formulas; nothing to type.")
P("Summary", "Detailed counts by status, lane and owner, plus money totals. All formulas; nothing to type.")
P("Portfolio Tracker", "One row per contract. Blue headers = you type. Grey headers marked (auto) = formulas, do not type over them.")
P("Task Order How To / Task Order Dashboard", "Instructions and charts for task orders and job order releases (teal tabs). See the Task Order How To tab for the full guide.")
P("Task Order Tracker", "One row per task order / job order release, linked to its master contract by PO. Shows each master's ceiling, total committed and available balance, and flags releases that exceed the ceiling, outlast the master, are overpaid, or are ending.")
P("Alert Controls", "Every list, threshold, rule and colour the alerts use. The only place to change the logic.")
P("Original (as received)", "The list exactly as it was received. Never edited; used to check against the source.")
P("Shared List (as received)", "The BWS shared-contracts file exactly as received. Never edited.")
P("DPS 365 Report", "The monthly DPS report of every department contract ending within a year (including ones managed by other procurement managers, for visibility). Grey columns show whether each contract is on your tracker, its responsibility group, your tracker status, next action and req #, and a suggested comment. Procurement managers fill in the DPS columns and return the report to DPS by the first of the month.")
P("DPS 365 (as received)", "The DPS 365 report exactly as received. Never edited.")
P("Data Check", "Every change made while loading the shared-contracts file: text dates and PO numbers converted, year typos fixed, contract-type spellings made consistent, and where the shared file disagreed with this tracker. Fixed cells are shaded light yellow on the Portfolio Tracker.")

H("3. Weekly routine for staff")
P("Step 1", "Sort the Portfolio Tracker by Priority (1 = most urgent). Work top down.")
P("Step 2", "For each contract you touched: update Renewal Stage and add a dated line at the top of Status & Notes.")
P("Step 3", "When a next action is finished, set 'Next Action Done?' to Yes, then enter the next one (or clear the date).")
P("Step 4", "Monthly: update Expended to Date from the financial system so the Funds Flag stays current.")
P("Step 5", "If a contract has no Renewal Path yet, pick one. Until then it is treated as the longest path (see Alert Controls, setting 1).")

H("4. Columns you fill in")
P("Renewal Path", "How this contract will be continued: extension, competitive rebid, sole source, or let expire. Sets the lead time.")
P("Renewal Stage", "Where the renewal work stands (formerly Action Status). Each stage counts as Normal, In Process or Closed (Alert Controls section 3). This is what the In Process / movement KPIs count.")
P("Status & Notes", "The one notes column, right after the contract details. Your notes come first, then the BWS shared-list notes and the DPS 365 comment, each labelled. Add new notes at the top with a date. The removed note columns are kept on the Archive - Removed Columns tab.")
P("Next Action / Due Date / Done?", "The single next step, when it is due, and Yes when finished.")
P("Total Contract Value / Expended to Date", "Dollar amounts. Needed for the Funds Flag.")

P("Contract Type (Portfolio Tracker)", "Standard, Master Service or Construction (JOC). Types marked 'Takes releases' on Alert Controls appear as masters on the Task Order Tracker and Dashboard.")
P("Task Order Tracker", "Master PO (dropdown), release number, type, scope, bureau, issue/start/end dates, release amount, paid to date, status, optional M/WBE goal and actual. Enter a release change order as its own row with the change amount (negative for a decrease).")
P("Owning Dept Code / Ownership", "Owning Dept Code is the PO Dept Code of the department that owns the contract. Ownership turns it into Owned (your code, set on Alert Controls section 15) or Shared (another department owns it, your department draws on it).")
P("Spec Number, PO Contract Type, BWS File Notes", "Carried over from the shared-contracts file. BWS File Notes keeps that file's notes separate from this tracker's Status / Notes.")
P("On DPS 365 Report? / DPS 365 Comments / DPS Procurement Mgr", "Formulas on the Portfolio Tracker. They show whether the PO is on the DPS 365 report, the comment and procurement manager entered there, and flag 'Missing - check with DPS' when your contract ends within a year but is not on the report. Tracker dates are never changed by the report.")
P("Reviewed (column A)", "Your own check-off column. Type anything (e.g. y or a date) when you have reviewed the row. Not used by any formula.")
P("Supply Coverage / Coverage PO / Note / Coverage Through", "For lapsed or soon-to-lapse contracts: how DWM keeps getting materials (pick from the list on Alert Controls section 18), the PO or note that proves it, and when that coverage ends. Values shaded yellow were suggested from your notes; confirm or change them.")
H("5. How the alerts work")
P("Start-By Date", "End Date minus the Lead Time for the contract's Renewal Path. This is the last day to start renewal work and still finish before the contract ends.")
P("Expiration Status", "Rules in Alert Controls section 4 are checked top to bottom; the first one that fits is shown. Default order: CLOSED (closed category) > LAPSED (past end date) > IN PROCESS (work submitted) > RED (180 days or less left, or Start-By date passed) > YELLOW (365 days or less, or Start-By within 90 days) > GREEN.")
P("Funds Flag", "% Expended = Expended / Total Value. % of Term Used = days elapsed / total contract days. 'Spending ahead of time' = % Expended minus % of Term Used. RED at 75% spent or 20+ points ahead of time; YELLOW at 50% or 10+ points ahead.")
P("Next Action Flag", "Days until the Next Action Due Date: OVERDUE when past due, DUE SOON within 14 days, UPCOMING within 30 days. Clears when Done? = Yes.")
P("Burn Index / Projected Exhaustion Date / Funds Run Out Before End?", "Burn Index = % spent divided by % of term used (1.00 = on pace; amber at 1.10, red at 1.25). Blank until 10% of the term has passed and money has been spent. Projected Exhaustion Date = when the money runs out at the current pace. Funds Run Out Before End? says Yes - N days early when that is before the end date: start a vendor limit increase.")
P("Supply Status / Supply Risk", "Supply Status is never blank: Active contract, your Supply Coverage choice, Not yet planned (ending within 90 days, no coverage entered) or Unknown - confirm (lapsed, no coverage entered). Supply Risk turns it into Covered / At risk / Not covered for the Dashboard.")
P("Action Needed and row colours", "The whole row turns red (Act now) or amber (Plan) when the contract needs your action: next action overdue, supply not covered, lapsed with supply at risk, or Priority 1-2 (Act now) / 3 (Plan) without a renewal in process. A renewal in process with supply covered is NOT highlighted. Levels on Alert Controls section 19.")
P("Priority", "The lowest (most urgent) Severity number among the three flags. Severity for each status is set in Alert Controls.")
P("Worked example", "Competitive rebid, end date 400 days away. Lead time 540 days -> Start-By date was 140 days ago. Days to Start-By = -140, which is <= 0, so RED even though a year is left. With a plain 365-day rule it would have shown GREEN. If Action Status is set to 'Submitted to DPS' it shows IN PROCESS (blue) instead.")

P("Shared contracts", "Shared contracts get the same alerts as owned ones. Pick the Renewal Path 'Owning dept renews' (365-day lead time = when your input is due) and use the owning-department Action Statuses to show where the other department stands.")
H("6. Why the logic is designed this way")
P("Start from the deadline and work backward", "Expiration risk depends on how long the renewal takes, not just on days left. A rebid needs about 18 months; an extension far less. Backward scheduling from the end date (end date minus lead time) shows when work must start, so alerts fire early enough to act.")
P("Separate 'needs me' from 'in process'", "Management by exception: RED should mean someone on our team must act. Contracts already submitted show blue IN PROCESS so they are watched but don't hide the ones nobody has started.")
P("Separate LAPSED from CLOSED", "An expired contract that is still needed is a compliance risk; one that was voided or replaced is not. Mixing them inflates the 'expired' count and hides the real problems.")
P("Spending compared with time", "75% spent means different things at month 6 and month 58. Comparing % spent to % of term used catches contracts burning money too fast, while the % spent limit still flags contracts close to their ceiling.")
P("Rules in a table, not in formulas", "Thresholds and statuses are team decisions that will change. Keeping them on one tab means changes are visible, consistent across every row, and do not require editing formulas.")
P("Colours", "Soft fills with dark text for readability, and the status word is always shown in the cell so the meaning never depends on colour alone (colour-blind readers, black-and-white printouts).")

P("Shared contracts on one list", "Your department depends on shared contracts as much as on its own. Keeping them on the same list means one set of alerts, one priority order and one dashboard; the Ownership column separates them whenever you need to.")
H("7. How to change the rules (after staff alignment)")
P("Change a number", "Alert Controls: edit the blue cell (lead times, day limits, % limits, severity). Everything updates.")
P("Add / rename a Renewal Path or Renewal Stage", "Type it into an empty slot (or over an existing one) in section 2 or 3. It appears in the tracker dropdown. For a new Action Status, pick its 'Counts As' category. Renaming does not update rows that already use the old name; re-pick them.")
P("Add a new category", "Section 3, column K: type it under Closed. It can then be used in 'Counts As' and in the rule tables' 'Applies when' column.")
P("Add or change a status rule", "Sections 4-6: fill an empty slot. Order matters: the first matching rule wins, so put specific rules above general ones and keep a catch-all (blank limits) last. Pick its Colour and Severity.")
P("Change a colour", "Pick a different name in the Colour column. Available colours are listed in section 7. Adding a brand-new colour needs a new conditional formatting rule (ask whoever maintains the file).")
P("Master contracts", "Ceiling = Total Contract Value on the Portfolio Tracker. Committed = sum of releases whose status counts as committed (Alert Controls section 12). Available = Ceiling - Committed. % Committed drives the Commitment Flag (section 13). Each release gets one Release Flag (section 14).")
P("Change task order settings", "Alert Controls sections 10-14: contract types and whether they take releases, release types, release statuses (committed? still open?), commitment thresholds, and release flag rules (label, condition, days, colour, severity, order). The list of release flag conditions is fixed; ask for a new one if you need it.")
P("Dashboard settings", "Alert Controls section 9: the 'expiring soon' window used on the Dashboard tiles and which priorities appear in the Top 10 list.")
P("Dashboard charts", "Chart colours are fixed to the starting statuses (CLOSED grey, LAPSED dark red, IN PROCESS blue, RED, YELLOW, GREEN). If you add a new expiration rule or lane, the tables update but the charts need their data range widened to show it.")
P("DPS 365 report", "Alert Controls section 16: the report window (365 days) and which procurement manager initials count as your portfolio or visibility only.")
P("Monthly DPS 365 routine", "1) Attach the new DPS report and your latest tracker, starting with 'Contract Tracker update: new 365 report'. 2) The DPS 365 Report tab is refreshed; your entries are kept. 3) For your rows, use the Suggested Comment to fill Additional Comments. 4) Check 'Missing from report' and 'End dates differ' on the Dashboard. 5) Send the DPS columns back to DPS by the first of the month.")
P("Lapsed contracts for leadership", "Dashboard rows 9-19: how many lapsed contracts have renewals in process (% with movement), how many DWM can still buy on (supply covered), what is at risk, lapsed 90+ days with no action, and average days lapsed, plus a one-sentence summary. A lapsed contract keeps the LAPSED status (it is a compliance fact); the movement and supply show underneath it.")
P("Spending pace, supply and row highlight settings", "Alert Controls sections 17 (Burn Index levels), 18 (Supply Coverage choices and what each counts as; add your own in the empty slots) and 19 (row highlight levels, supply planning window, coverage warning).")
P("Department code", "Alert Controls section 15: your department code (88) and the list of owning-department codes shown on the Dashboard.")
P("Capacity", "12 supply coverage choices, 10 renewal paths, 12 action statuses, 10 expiration rules, 8 funds rules, 8 next-action rules, 300 contract rows.")
P("Log the change", "Record every rule change in the change log below so leadership can see what changed and why.")

H("8. Sources and basis for the alert logic")
P("Thresholds and lead times", "Set by the contract administration team from working experience with the DPS process (the original Settings tab: 180/365-day stoplight, 540-day competitive and 365-day sole-source/NCRB lead times, 75%/50% funds limits). The 240-day extension lead time comes from the team note 'start TE 8 months ahead'. These are internal operating assumptions, not requirements from an outside standard; confirm them with DPS and adjust here.")
P("City procurement rules", "Municipal Code of Chicago, Title 2, Chapter 2-92 (Contracts) and the Department of Procurement Services (DPS) policies and procedures (chicago.gov/dps). Consult DPS for current solicitation, extension and NCRB requirements and timelines.")
P("Stoplight (RAG) status reporting", "Common project and portfolio management practice for status reporting (e.g., Project Management Institute, PMBOK Guide). The specific day and percent cut-offs are ours, not PMI's.")
P("Spending compared with time", "A simplified form of the earned value idea (planned vs. actual progress) described in the PMBOK Guide and the EIA-748 Earned Value Management standard. This tracker compares spending with elapsed time only; it is not a full earned value analysis.")
P("Backward scheduling", "Standard planning method: required completion date minus lead time = latest start date. Used here to set the Start-By Date.")
P("Colour and accessibility", "W3C Web Content Accessibility Guidelines (WCAG) 2.1: 1.4.1 Use of Color (don't rely on colour alone) and 1.4.3 Contrast. Red-green colour vision deficiency affects roughly 1 in 12 men (National Eye Institute).")
P("Contract data", "Columns A-O come from the contract list as received (see Original tab), re-sorted by lane and end date. Lane is whichever of the Facility / Water-Sewer / Pipes / Valves / Meters / Chemical columns has an x. Proposed Owner follows the 10/21 transition meeting. Four start dates on the original list were typed as text with year typos (e.g. 11/22/0024); they were converted to real dates on 11/02/2026 and are listed on the Data Check tab. The Original tab keeps them as received.")

H("9. Change log")
r = row[0]
for j, h in enumerate(["Date", "Change / reason"]):
    c = hw[f"{'BC'[j]}{r}"]; c.value = h; c.font = font(bold=True, color="FFFFFF"); c.fill = fill(BLUE_HDR); c.border = BOX
row[0] += 1
P("10/31/2026", "Alert logic rebuilt: renewal-path lead times and Start-By date, Action Status categories, LAPSED / IN PROCESS / CLOSED statuses, burn-rate funds check, Next Action flag, Priority column, all rules moved to Alert Controls. Colours softened. Settings tab merged into Alert Controls. Broken formulas in the old row 7 removed.", bold_a=False)
P("10/31/2026", "Dashboard tab added for portfolio owner and leadership reporting; Alert Controls section 9 added for dashboard settings.", bold_a=False)
P("11/01/2026", "Task Order Tracker tab added for task orders and job order releases against master and construction (JOC) contracts; Contract Type column added to the Portfolio Tracker; Alert Controls sections 10-14 added; Dashboard tiles added. Task Order Dashboard and Task Order How To tabs added; tabs colour-coded (navy = contracts, teal = task orders, gold = controls).", bold_a=False)
P("11/01/2026", "Column Guide tab added: every tracker column with its type, plain-language meaning, how it is figured and an example. Examples kept off the trackers so they never reach the dashboards.", bold_a=False)
P("11/02/2026", "Shared contracts added: 53 contracts from the BWS shared-contracts file added to the Portfolio Tracker and 17 existing rows enriched. New columns Owning Dept Code, Ownership, Spec Number, PO Contract Type and BWS File Notes; Alert Controls section 15; renewal path and three action statuses for shared contracts; Dashboard and Summary shared-contract sections; Data Check and Shared List (as received) tabs.", bold_a=False)
P("11/02/2026", "Four original-list start dates typed as text with year typos (11/22/0024, 07/11/0007, 02/08/0024, 05/05/0023) converted to real dates; listed on the Data Check tab.", bold_a=False)
P("11/03/2026", "DPS 365 Report tab added (124 department contracts ending within a year, linked to the tracker, with suggested comments); DPS 365 (as received) tab; Portfolio Tracker columns On DPS 365 Report?, DPS 365 Comments and DPS Procurement Mgr; Alert Controls section 16; Dashboard DPS 365 section. Tracker dates unchanged.", bold_a=False)
P("11/08/2026", "Added Burn Index, Projected Exhaustion Date, Funds Run Out Before End?, Supply Coverage / Coverage PO / Note / Coverage Through, Supply Status, Supply Risk and Action Needed (whole-row red / amber highlight). Dashboard: lapsed renewal x supply KPIs, grid, chart and summary sentence; Funds run out early tile. Alert Controls sections 17-19. Your Reviewed column (A) and all your entries kept.", bold_a=False)
for _ in range(8): P("", "", bold_a=False)

exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "contract_tracker_dashboard_part_JHK3.py")).read())
exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "contract_tracker_taskorder_part_JHK3.py")).read())
exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "contract_tracker_shared_report_part_JHK3.py")).read())
exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "contract_tracker_dps365_part_JHK3.py")).read())
exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "contract_tracker_supply_part_JHK3.py")).read())
exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "contract_tracker_column_guide_part_JHK3.py")).read())
order = ["How To Use", "Column Guide", "Dashboard", "Summary", "Portfolio Tracker", "DPS 365 Report", "Task Order How To", "Task Order Dashboard", "Task Order Tracker", "Alert Controls", "Original (as received)"] + [n for n in ("Shared List (as received)", "DPS 365 (as received)", "Data Check", "Archive - Removed Columns") if n in wb.sheetnames] + ["Dashboard Data"]
for _n in ["How To Use", "Column Guide", "Dashboard", "Summary", "Portfolio Tracker", "DPS 365 Report"]: wb[_n].sheet_properties.tabColor = "1F3864"
for _n in ["Task Order How To", "Task Order Dashboard", "Task Order Tracker"]: wb[_n].sheet_properties.tabColor = "0E5E5A"
wb["Alert Controls"].sheet_properties.tabColor = "BF9000"
wb._sheets = [wb[n] for n in order]
wb.active = 0
wb.properties.title = "Contract Tracker"
# Carry a working copy's typed entries into this build. The working copy is the source of truth for every input
# (non-formula) cell in the areas people type in; formulas, headers and guide text come from this build.
# REVIEWED_COL=1: the working copy has a "Reviewed" column in A on the Portfolio Tracker (everything shifted one
# column right). Its values are parked in column ZZ here; insert_reviewed_column_JHK3.py moves them back to A.
REVIEWED = os.environ.get("REVIEWED_COL") == "1"
if os.environ.get("CARRY_USER"):
    _u = openpyxl.load_workbook(os.environ["CARRY_USER"])
    _n = 0
    def _same(a, b):
        if a == b: return True
        if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool) and not isinstance(b, bool):
            return abs(a - b) < 1e-9
        return a in (None, "") and b in (None, "")
    for _name, _r0, _cmax in (("Portfolio Tracker", 2, None), ("Task Order Tracker", 5, None),
                              ("DPS 365 Report", 6, 11), ("Alert Controls", 1, None)):
        if _name not in _u.sheetnames or _name not in wb.sheetnames:
            continue
        _uw, _tw = _u[_name], wb[_name]
        _sh = 1 if (REVIEWED and _name == "Portfolio Tracker") else 0
        for _r in range(_r0, _uw.max_row + 1):
            for _uc in range(1 + _sh, _uw.max_column + 1):
                _tc = _uc - _sh
                if _cmax and _tc > _cmax:
                    continue
                _uv = _uw.cell(_r, _uc).value
                if isinstance(_uv, str) and _uv.startswith("="):
                    continue
                _t = _tw.cell(_r, _tc)
                if isinstance(_t.value, str) and str(_t.value).startswith("="):
                    continue
                if _name == "Alert Controls" and not (_t.font and _t.font.color and str(_t.font.color.rgb).endswith("0000FF")):
                    continue           # only the blue input cells on Alert Controls
                if not _same(_t.value, _uv):
                    _t.value = _uv; _n += 1
    if REVIEWED:
        _uw = _u["Portfolio Tracker"]
        for _r in range(1, _uw.max_row + 1):
            wb["Portfolio Tracker"][f"ZZ{_r}"].value = _uw.cell(_r, 1).value
    print("carried", _n, "input cells from", os.environ["CARRY_USER"], "(Reviewed column parked in ZZ)" if REVIEWED else "")
# The Reviewed column always exists (parked in ZZ until finalize_layout_JHK3.py moves it to A).
_pz = wb["Portfolio Tracker"]
_pz["ZZ1"] = "Reviewed"
_pz["ZZ1"].font = font(bold=True, color="FFFFFF"); _pz["ZZ1"].fill = fill(BLUE_HDR); _pz["ZZ1"].border = BOX
_pz["ZZ1"].alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
_pz.column_dimensions["ZZ"].width = 10

# Supply Coverage suggestions from the notes, for lapsed / soon-to-lapse contracts with no coverage entered.
import re as _sre, datetime as _sdt
_today = _sdt.datetime.now()
_pt = wb["Portfolio Tracker"]; _ac = wb["Alert Controls"]
_cat_of = {_ac[f"C{_r}"].value: _ac[f"D{_r}"].value for _r in range(27, 39) if _ac[f"C{_r}"].value}
_win = _ac["D249"].value or 90
_sugg = []
for _r in range(2, LAST + 1):
    if _pt[f"A{_r}"].value in (None, "") or _pt[f"BD{_r}"].value not in (None, ""):
        continue
    _end = _pt[f"D{_r}"].value
    if not isinstance(_end, _sdt.datetime) or (_end - _today).days > _win:
        continue
    if _cat_of.get(_pt[f"S{_r}"].value) == "Closed":
        continue
    _txt = " | ".join(str(_pt[f"{_c}{_r}"].value) for _c in ("AM", "AN", "AW") if _pt[f"{_c}{_r}"].value)
    _low = _txt.lower(); _val = _note = None
    if _sre.search(r"(can we\s*(we\s*)?use|find out if|is)\s+(exh|exhibit|ex)\.?\s*b\b[^.?]*\?|find out if\s+(exh|exhibit|ex)\.?\s*b", _low):
        _val = "Unconfirmed"; _note = "Exhibit B not yet confirmed (from notes)"
    elif _sre.search(r"\b(exh|exhibit|ex)\.?\s*b\b\s*(is\s+)?(allowed|can be used|could be used)|use exhibit b", _low):
        _val = "Exhibit B"; _note = "Exhibit B noted as allowed (from notes)"
    else:
        _m = _sre.search(r"(?:use|new(?: contract)?)\s+po\s*[#:]?\s*(\d{5,6})", _low)
        if _m:
            _val = "Alternate PO / contract"; _note = f"PO {_m.group(1)} (from notes)"
    if _val:
        _pt[f"BD{_r}"] = _val; _pt[f"BD{_r}"].fill = fill("FFF2CC")
        _sugg.append((_r, _pt[f"A{_r}"].value, _val, _note, _txt[:160]))
if _sugg and "Data Check" in wb.sheetnames:
    _dc = wb["Data Check"]; _r0 = _dc.max_row + 3
    _dc[f"A{_r0}"] = f"Supply Coverage suggested from notes ({_today:%m/%d/%Y}) - shaded yellow on the tracker; confirm or change"
    _dc[f"A{_r0}"].font = font(bold=True, color=NAVY, size=12)
    for _j, _h in enumerate(["#", "Tracker row", "PO", "Suggested coverage", "Coverage note", "Based on this note text"]):
        _c = _dc.cell(_r0 + 1, 1 + _j, _h); _c.font = font(bold=True, color="FFFFFF"); _c.fill = fill(NAVY); _c.border = BOX
    for _i, (_r, _po, _val, _note, _txt) in enumerate(_sugg, 1):
        for _j, _v in enumerate([_i, _r, _po, _val, _note, _txt]):
            _c = _dc.cell(_r0 + 1 + _i, 1 + _j, _v); _c.font = font(); _c.border = BOX
            _c.alignment = Alignment(wrap_text=True, vertical="top")
    print("supply coverage suggested on", len(_sugg), "rows")

# Status & Notes: combine the old note columns (only where Status & Notes is still empty) and archive them.
_dps = wb["DPS 365 Report"] if "DPS 365 Report" in wb.sheetnames else None
_dps_note = {}
if _dps is not None:
    for _r in range(6, 306):
        if _dps[f"D{_r}"].value not in (None, "") and _dps[f"J{_r}"].value not in (None, ""):
            _dps_note[_dps[f"D{_r}"].value] = _dps[f"J{_r}"].value
def _txt(v):
    if v is None: return ""
    if isinstance(v, _sdt.datetime): return v.strftime("%m/%d/%Y")
    return str(v).strip()
_ARCH = [("T", "Next Action"), ("U", "Next Action Due Date"), ("V", "Next Action Done?"), ("AM", "Notes / Next Action"),
         ("AN", "Status /Notes"), ("AW", "BWS File Notes"), (None, "DPS 365 Comment"), ("BE", "Coverage PO / Note"),
         ("BF", "Coverage Through")]
_arch_rows = []; _merged = 0
for _r in range(2, LAST + 1):
    _po = _pt[f"A{_r}"].value
    if _po in (None, ""):
        continue
    _vals = {h: (_dps_note.get(_po) if c is None else _pt[f"{c}{_r}"].value) for c, h in _ARCH}
    if any(v not in (None, "") for v in _vals.values()):
        _arch_rows.append((_r, _po, _pt[f"B{_r}"].value, _vals))
    if _pt[f"BJ{_r}"].value in (None, ""):
        _parts = []
        if _txt(_vals["Status /Notes"]): _parts.append("Status: " + _txt(_vals["Status /Notes"]))
        if _txt(_vals["Notes / Next Action"]): _parts.append("Notes: " + _txt(_vals["Notes / Next Action"]))
        if _txt(_vals["Next Action"]):
            _nx = "Next: " + _txt(_vals["Next Action"])
            if _txt(_vals["Next Action Due Date"]): _nx += " (due " + _txt(_vals["Next Action Due Date"]) + ")"
            if _txt(_vals["Next Action Done?"]): _nx += " [done: " + _txt(_vals["Next Action Done?"]) + "]"
            _parts.append(_nx)
        if _txt(_vals["BWS File Notes"]): _parts.append("BWS: " + _txt(_vals["BWS File Notes"]))
        if _txt(_vals["DPS 365 Comment"]): _parts.append("DPS 365: " + _txt(_vals["DPS 365 Comment"]))
        _cv = [_txt(_vals["Coverage PO / Note"]), _txt(_vals["Coverage Through"])]
        if _cv[0] and "(from notes)" not in _cv[0]: _parts.append("Coverage: " + _cv[0] + (" through " + _cv[1] if _cv[1] else ""))
        if _parts:
            _pt[f"BJ{_r}"] = "\n".join(_parts); _merged += 1
_ar = wb.create_sheet("Archive - Removed Columns")
_ar.sheet_view.showGridLines = False
_ar["A1"] = "Archive - Removed Columns"; _ar["A1"].font = font(size=16, bold=True, color=NAVY)
_ar["A2"] = ("Values from the Portfolio Tracker columns removed on " + _today.strftime("%m/%d/%Y") +
             " (combined into Status & Notes). Kept for reference; nothing here feeds a formula.")
_ar["A2"].font = font(size=9, italic=True, color=GREY_HDR)
_hd = ["Tracker row", "PO Number", "PO Description"] + [h for _, h in _ARCH]
for _j, _h in enumerate(_hd):
    _c = _ar.cell(4, 1 + _j, _h); _c.font = font(bold=True, color="FFFFFF"); _c.fill = fill(NAVY); _c.border = BOX
    _c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
for _i, (_r, _po, _desc, _vals) in enumerate(_arch_rows):
    for _j, _v in enumerate([_r, _po, _desc] + [_vals[h] for _, h in _ARCH]):
        _c = _ar.cell(5 + _i, 1 + _j, _v); _c.font = font(size=9); _c.border = BOX
        _c.alignment = Alignment(wrap_text=True, vertical="top")
        if isinstance(_v, _sdt.datetime): _c.number_format = "mm/dd/yyyy"
for _col, _w in zip("ABCDEFGHIJKL", [8, 10, 30, 30, 11, 8, 40, 40, 40, 40, 20, 11]):
    _ar.column_dimensions[_col].width = _w
_ar.freeze_panes = "D5"; _ar.auto_filter.ref = f"A4:L{4 + len(_arch_rows)}"
wb._sheets.remove(_ar); wb._sheets.insert(wb._sheets.index(wb["Dashboard Data"]), _ar)
print("Status & Notes combined on", _merged, "rows; archived", len(_arch_rows), "rows")

if os.environ.get("TEST_TO"):
    exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "contract_tracker_testdata_part_JHK3.py")).read())
if os.environ.get("RENDER_DASH"):
    _keep = os.environ["RENDER_DASH"] if os.environ["RENDER_DASH"] != "1" else "Dashboard"
    wb.active = wb.sheetnames.index(_keep)
    for w in wb.worksheets:
        if w.title != _keep: w.sheet_state = "hidden"
wb.save(OUT)
print("saved", OUT)
