"""Column Guide tab shared by the Contract Tracker and Project Tracker builds.

add_column_guide(wb, title, intro, types, rows, index) writes one table: a row per tracker column with its
type (who fills it), what it means, what to enter or how it is calculated, an example, and where its
settings live. Examples live here, not as sample rows on the trackers, so they never reach a dashboard.
"""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

_F = "Arial"
_thin = Side(style="thin", color="BFBFBF")
_BOX = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
_HEAD = "1F3864"


def _font(**k):
    k.setdefault("name", _F); k.setdefault("size", 10)
    return Font(**k)


def _fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def add_column_guide(wb, title, intro, types, rows, index=1, sheet_name="Column Guide"):
    """types: {type_name: (fill_hex, text_hex, legend_text)}.
    rows: (tab, column_letter, column_name, type_name, meaning, how, example, settings)."""
    ws = wb.create_sheet(sheet_name, index)
    ws.sheet_view.showGridLines = False
    widths = {"A": 2, "B": 20, "C": 9, "D": 26, "E": 22, "F": 44, "G": 50, "H": 24, "I": 24}
    for c, w in widths.items():
        ws.column_dimensions[c].width = w
    ws["B1"] = title; ws["B1"].font = _font(size=16, bold=True, color=_HEAD)
    ws["B2"] = intro; ws["B2"].font = _font(size=9, italic=True, color="595959")
    ws.merge_cells("B2:I2"); ws["B2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 30

    r = 4
    ws[f"B{r}"] = "How to read the Type column"; ws[f"B{r}"].font = _font(size=11, bold=True, color=_HEAD)
    for name, (fc, tc, text) in types.items():
        r += 1
        c = ws[f"B{r}"]; c.value = name; c.fill = _fill(fc); c.font = _font(bold=True, color=tc); c.border = _BOX
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(f"C{r}:I{r}")
        d = ws[f"C{r}"]; d.value = text; d.font = _font(); d.alignment = Alignment(vertical="center", wrap_text=True)
    r += 1
    ws[f"B{r}"] = "Tip: use the filter arrows to show one tab or one type at a time."
    ws[f"B{r}"].font = _font(size=9, italic=True, color="595959")

    hr = r + 2
    heads = ["Tab", "Column", "Column Name", "Type", "What it means", "What to enter / how it's figured",
             "Example", "Settings that change it"]
    for j, h in enumerate(heads):
        c = ws.cell(hr, 2 + j, h)
        c.font = _font(bold=True, color="FFFFFF"); c.fill = _fill(_HEAD); c.border = _BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[hr].height = 28
    for i, row in enumerate(rows):
        rr = hr + 1 + i
        for j, v in enumerate(row):
            c = ws.cell(rr, 2 + j, v)
            c.border = _BOX; c.font = _font(bold=(j in (2,)))
            c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="center" if j in (1, 3) else "left")
        fc, tc, _ = types[row[3]]
        t = ws.cell(rr, 5); t.fill = _fill(fc); t.font = _font(bold=True, color=tc)
        ws.cell(rr, 3).fill = _fill("F2F2F2")
    ws.freeze_panes = ws.cell(hr + 1, 3)
    ws.auto_filter.ref = f"B{hr}:I{hr + len(rows)}"
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{hr}:{hr}"
    return ws
