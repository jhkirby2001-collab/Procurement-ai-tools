# Shared contracts: load the BWS shared-contracts list into the Portfolio Tracker.
# Exec'd inside build_contract_tracker_JHK3.py right after the original rows are copied (uses its globals).
# Shared contracts are owned (PO Dept Code) by another department but the department draws from them.
import re as _re
import datetime as _dt

SH_COLS = {"AS": ("Owning Dept Code", "in", 11), "AT": ("Ownership", "auto", 11), "AU": ("Spec Number", "in", 12),
           "AV": ("PO Contract Type", "in", 24), "AW": ("BWS File Notes", "in", 50)}
YELLOW = "FFF2CC"

# ---- Alert Controls section 15
sec(181, "15. Shared contracts  (owned by another department, drawn on by yours)",
    "Ownership = Owned when a contract's Owning Dept Code equals your department code, otherwise Shared. The code list feeds the Dashboard table.")
hdr(183, {"B": "#", "C": "Setting", "D": "Value", "I": "What it does"})
ctl["B184"] = 1; ctl["B184"].font = font(); ctl["B184"].border = BOX
ctl["C184"] = "My department code"; ctl["C184"].font = font(); ctl["C184"].border = BOX
inp("D184", 88, "0"); ctl["D184"].fill = fill("FFFF00")
note("I184", "Your department's PO Dept Code. Every contract with this code shows as Owned; any other code shows as Shared.")
ctl.merge_cells("E184:H184")
MY_CODE = f"{AC}!$D$184"
hdr(186, {"B": "Slot", "C": "Owning Dept Codes (for the Dashboard table)", "I": "Notes"})
CODE0, CODE1 = 187, 198

# ---- read the shared file
SHARED_SRC = sys.argv[3] if len(sys.argv) > 3 else None
shared_ws = openpyxl.load_workbook(SHARED_SRC).worksheets[0] if SHARED_SRC else None
fixes = []          # (shared_row, po, field, before, after, note, tracker_row)
TYPE_MAP = {
    "COMMODITIES": "COMMODITIES", "Commodities": "COMMODITIES",
    "WORK SERVICES / FACILITIES MAINT.": "WORK SERVICES / FACILITIES MAINT.",
    "Work Services / FACILITIES MAINT.": "WORK SERVICES / FACILITIES MAINT.",
    "Work Serices": "WORK SERVICES", "Work Services ": "WORK SERVICES", "Work Services": "WORK SERVICES",
    "WORK SERV-AVIATION": "WORK SERVICES - AVIATION", "COMMODITIES-AVIATION": "COMMODITIES - AVIATION",
    "COMMODITIES-SMALL ORDERS": "COMMODITIES - SMALL ORDERS",
    "PRO SERV CONSULTING $250,000orABOVE": "PROFESSIONAL SERVICES - CONSULTING $250,000 OR ABOVE",
}

def _parse_date(v):
    """Real date from a text date. 'X - new end date Y' -> Y. Two-digit-century typos like 0016 -> 2016."""
    if isinstance(v, _dt.datetime) or v is None:
        return v, None
    s = str(v)
    m = _re.search(r"new end date\s*(\d{1,2})/+(\d{1,2})/+(\d{2,4})", s, _re.I)
    note_ = "Used the new end date written in the cell" if m else None
    if not m:
        m = _re.search(r"(\d{1,2})/+(\d{1,2})/+(\d{2,4})", s)
    if not m:
        return v, "Could not read a date"
    mo, d, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if y < 100:
        y += 2000
    elif y < 1000:
        y = 2000 + (y % 100); note_ = (note_ + "; " if note_ else "") + "Fixed year typo"
    return _dt.datetime(y, mo, d), note_

def _norm_po(v):
    return _re.sub(r"\D", "", str(v)) if v is not None else ""

tracker_rows = {}
for r in range(2, 88):
    tracker_rows.setdefault(_norm_po(new[f"A{r}"].value), r)

def _style_fixed(cell):
    cell.fill = fill(YELLOW)

next_row = 88
n_over = 0
codes_seen = set()
if shared_ws is not None:
    # untouched copy of the file
    asis = wb.create_sheet("Shared List (as received)")
    for row in shared_ws.iter_rows():
        for c in row:
            d = asis.cell(c.row, c.column, c.value)
            if c.has_style:
                d.font = copy(c.font); d.fill = copy(c.fill); d.border = copy(c.border)
                d.alignment = copy(c.alignment); d.number_format = c.number_format
    for k, dim in shared_ws.column_dimensions.items():
        asis.column_dimensions[k].width = dim.width
    asis.freeze_panes = "A2"

    for sr in range(2, shared_ws.max_row + 1):
        vals = [shared_ws.cell(sr, c).value for c in range(1, 13)]
        if all(v in (None, "") for v in vals):
            continue
        end_raw, start_raw, spec, po_raw, desc, limit, ctype, unit, vendor, code, analyst, notes = vals
        # PO number
        po, po_note = po_raw, None
        if isinstance(po_raw, str):
            nums = _re.findall(r"\d+", po_raw)
            if _re.match(r"\s*\(\d+\)", po_raw) and len(nums) == 2:
                po, po_note = int(nums[1]), f"Formerly PO {nums[0]}"
            elif len(nums) == 1:
                po = int(nums[0])
        po_fix = (sr, po, "PO Number", po_raw, po, po_note or "Stored as text; converted to a number") if po != po_raw else None
        end, end_note = _parse_date(end_raw)
        start, start_note = _parse_date(start_raw)
        ctype_clean = TYPE_MAP.get(ctype, str(ctype).strip() if ctype else ctype)
        if code is not None:
            codes_seen.add(code)
        notes_all = notes
        if po_note:
            notes_all = f"{po_note}. {notes}" if notes else po_note

        key = _norm_po(po)
        if key in tracker_rows and tracker_rows[key] < 88:
            # already on the tracker: keep its dates, owners and notes; add the shared-file fields
            tr = tracker_rows[key]; n_over += 1
            if po_fix:
                fixes.append(po_fix + (tr,))
            new[f"AS{tr}"] = code; new[f"AU{tr}"] = spec; new[f"AV{tr}"] = ctype_clean; new[f"AW{tr}"] = notes_all
            if new[f"W{tr}"].value in (None, ""):
                new[f"W{tr}"] = limit
            if end is not None and isinstance(end, _dt.datetime) and new[f"D{tr}"].value != end:
                fixes.append((sr, po, "End Date", end_raw, new[f"D{tr}"].value,
                              "Already on the tracker; kept the tracker's end date (shared file differs)", tr))
            if code != 88 and str(new[f"G{tr}"].value).strip() == "88":
                fixes.append((sr, po, "Owning Dept Code", "Your list: Dept Code 88", code,
                              "Shared file says another department owns this PO, so it shows as Shared. Change the code to 88 if it is yours.", tr))
            if ctype_clean != ctype:
                fixes.append((sr, po, "PO Contract Type", ctype, ctype_clean, "Spelling made consistent", tr))
            continue

        tr = next_row; next_row += 1
        new[f"A{tr}"] = po; new[f"B{tr}"] = desc; new[f"C{tr}"] = start; new[f"D{tr}"] = end
        new[f"E{tr}"] = vendor; new[f"H{tr}"] = analyst; new[f"W{tr}"] = limit
        new[f"AS{tr}"] = code; new[f"AU{tr}"] = spec; new[f"AV{tr}"] = ctype_clean; new[f"AW{tr}"] = notes_all
        for col in ("C", "D"):
            new[f"{col}{tr}"].number_format = "mm/dd/yyyy"
        if po_fix:
            fixes.append(po_fix + (tr,)); _style_fixed(new[f"A{tr}"])
        if end is not end_raw:
            fixes.append((sr, po, "End Date", end_raw, end, end_note or "Text date converted to a real date", tr)); _style_fixed(new[f"D{tr}"])
        if start is not start_raw:
            fixes.append((sr, po, "Start Date", start_raw, start, start_note or "Text date converted to a real date", tr)); _style_fixed(new[f"C{tr}"])
        if ctype_clean != ctype:
            fixes.append((sr, po, "PO Contract Type", ctype, ctype_clean, "Spelling made consistent", tr)); _style_fixed(new[f"AV{tr}"])
        tracker_rows.setdefault(key, tr)

# Original list: start dates typed as text with a year typo (e.g. 11/22/0024) -> real dates
start_fixed = 0
for r in range(2, 88):
    v = new[f"C{r}"].value
    if v is not None and not isinstance(v, _dt.datetime):
        d, why = _parse_date(v)
        if isinstance(d, _dt.datetime):
            new[f"C{r}"] = d; new[f"C{r}"].number_format = "mm/dd/yyyy"; _style_fixed(new[f"C{r}"])
            fixes.append(("Original list", new[f"A{r}"].value, "Start Date", v, d,
                          "Text date with a year typo; converted to a real date" + (" (2007 confirmed)" if str(v).endswith("0007") else ""), r))
            start_fixed += 1

# Contracts only on the original list: owning department = their Dept Code (all 88 in the original list)
set_from_dept = 0
for r in range(2, 88):
    if new[f"A{r}"].value is not None and new[f"AS{r}"].value in (None, ""):
        g = new[f"G{r}"].value
        try:
            new[f"AS{r}"] = int(str(g).strip()); set_from_dept += 1
        except (TypeError, ValueError):
            pass
SHARED_LAST_ROW = next_row - 1

# code list on Alert Controls
codes = sorted(codes_seen | {88})
for i, rr in enumerate(range(CODE0, CODE1 + 1)):
    ctl[f"B{rr}"] = i + 1; ctl[f"B{rr}"].font = font(color=GREY_HDR); ctl[f"B{rr}"].border = BOX
    inp(f"C{rr}", codes[i] if i < len(codes) else None, "0")
    note(f"I{rr}", "Your department" if i < len(codes) and codes[i] == 88 else None)
    ctl.merge_cells(f"D{rr}:H{rr}")
CODE_LIST = f"{AC}!$C${CODE0}:$C${CODE1}"

# ---- tracker columns AS-AW
for col, (h, kind, w) in SH_COLS.items():
    c = new[f"{col}1"]; c.value = h + (" (auto)" if kind == "auto" else "")
    c.font = font(bold=True, color="FFFFFF"); c.fill = fill(BLUE_HDR if kind == "in" else GREY_HDR)
    c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center"); c.border = BOX
    new.column_dimensions[col].width = w
for r in range(2, LAST + 1):
    for col in SH_COLS:
        x = new[f"{col}{r}"]; x.border = BOX; x.font = font()
        x.alignment = Alignment(vertical="top", wrap_text=(col == "AW"))
    t = new[f"AT{r}"]
    t.value = f'=IF(A{r}="","",IF(AS{r}="","Not set",IF(AS{r}&""={MY_CODE}&"","Owned","Shared")))'
    t.fill = fill("F2F2F2"); t.font = font(bold=True); t.alignment = Alignment(horizontal="center", vertical="top")
new.conditional_formatting.add(f"AT2:AT{LAST}", FormulaRule(formula=['$AT2="Shared"'], fill=fill("E1D5F0"), font=Font(color="4B2C73", bold=True)))
new.conditional_formatting.add(f"AT2:AT{LAST}", FormulaRule(formula=['$AT2="Not set"'], font=Font(color="C00000", italic=True)))

# ---- Data Check tab
dc = wb.create_sheet("Data Check")
dc.sheet_view.showGridLines = False
dc["A1"] = "Data Check: fixes made while loading"; dc["A1"].font = font(size=16, bold=True, color=NAVY)
dc["A2"] = ("Every change made while loading the BWS shared-contracts file and fixing the original list's start dates. Fixed cells on the Portfolio Tracker are shaded light yellow. "
            "The untouched file is on the 'Shared List (as received)' tab.")
dc["A2"].font = font(size=9, italic=True, color=GREY_HDR)
n_new = SHARED_LAST_ROW - 87
summary = [
    ("Shared-file contracts added as new rows", n_new),
    ("Shared-file contracts already on your tracker (dates, owners and notes kept; code, spec, type, notes and value added)", n_over),
    ("Original-list contracts with no shared-file row: Owning Dept Code set from their Dept Code", set_from_dept),
    ("Original-list start dates fixed (text dates with a year typo)", start_fixed),
    ("Changes listed below", len(fixes)),
]
rr = 4
for lab, v in summary:
    dc[f"A{rr}"] = v; dc[f"A{rr}"].font = font(bold=True); dc[f"A{rr}"].alignment = Alignment(horizontal="center")
    dc[f"B{rr}"] = lab; dc[f"B{rr}"].font = font(); rr += 1
dc[f"A{rr}"] = ("PO 87372 appears twice in the shared file (two descriptions). Both rows were kept as you asked; "
                "the duplicate PO stays highlighted red on the tracker until you decide.")
dc[f"A{rr}"].font = font(size=9, color="C00000"); rr += 2
heads = ["#", "Source row", "PO", "Field", "As received", "Changed to", "Why", "Tracker row"]
for j, h in enumerate(heads):
    c = dc.cell(rr, 1 + j, h); c.font = font(bold=True, color="FFFFFF"); c.fill = fill(NAVY); c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
for i, f in enumerate(fixes, 1):
    vals = [i] + list(f)
    for j, v in enumerate(vals):
        c = dc.cell(rr + i, 1 + j, v); c.font = font(); c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="center" if j in (0, 1, 2, 7) else "left")
        if isinstance(v, _dt.datetime):
            c.number_format = "mm/dd/yyyy"
for col, w in zip("ABCDEFGH", [5, 10, 10, 18, 34, 22, 60, 10]):
    dc.column_dimensions[col].width = w
dc.freeze_panes = dc.cell(rr + 1, 1)
dc.auto_filter.ref = f"A{rr}:H{rr + len(fixes)}"
dc.page_setup.orientation = "landscape"; dc.page_setup.fitToWidth = 1; dc.page_setup.fitToHeight = 0
dc.sheet_properties.pageSetUpPr.fitToPage = True
