from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter as L
import sys

OUT = sys.argv[1]
F = "Arial"
NAVY, GREY, LIGHT, AUTO = "1F3864", "595959", "D9E1F2", "F2F2F2"
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
RED, AMBER, GREEN = "F8CBAD", "FFE699", "C6EFCE"

def font(**k):
    k.setdefault("name", F); k.setdefault("size", 10)
    return Font(**k)

def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)

def title(ws, text, sub):
    ws["A1"] = text; ws["A1"].font = font(size=16, bold=True, color=NAVY)
    ws["A2"] = sub; ws["A2"].font = font(size=9, italic=True, color=GREY)
    ws.sheet_view.showGridLines = False

wb = Workbook()
readme = wb.active; readme.title = "Read Me"
dash = wb.create_sheet("Dashboard")
reg = wb.create_sheet("Contract Register")
task = wb.create_sheet("Task Tracker")
iss = wb.create_sheet("Issue Log")
mat = wb.create_sheet("Priority Matrix")
guide = wb.create_sheet("Guide")
lists = wb.create_sheet("Lists")
# ---------------------------------------------------------------- tracker builder
def tracker(ws, ttl, sub, cols, nrows, freeze):
    """cols: list of (header, width, kind, extra) kind in input/auto/date/money/pct/int; extra=list key or formula fn"""
    title(ws, ttl, sub)
    ws["A3"] = "White columns = you type.   Grey columns marked (auto) = formulas, do not overwrite."
    ws["A3"].font = font(size=9, color=GREY)
    first, last = 5, 4 + nrows
    for j, (h, w, kind, extra) in enumerate(cols, 1):
        c = L(j)
        cell = ws[f"{c}4"]
        cell.value = h + (" (auto)" if kind.startswith("auto") else "")
        cell.font = font(bold=True, color="FFFFFF")
        cell.fill = fill(GREY if kind.startswith("auto") else NAVY)
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = BOX
        ws.column_dimensions[c].width = w
        fmt = {"date": "mm/dd/yyyy", "money": '$#,##0;($#,##0);"-"', "pct": "0.0%",
               "auto_money": '$#,##0;($#,##0);"-"', "auto_pct": "0.0%", "auto_date": "mm/dd/yyyy",
               "auto_int": "0;-0;0", "int": "0"}.get(kind)
        for r in range(first, last + 1):
            x = ws[f"{c}{r}"]
            x.border = BOX
            x.font = font()
            x.alignment = Alignment(vertical="top", wrap_text=kind in ("text",))
            if fmt: x.number_format = fmt
            if kind.startswith("auto"):
                x.fill = fill(AUTO)
                x.value = extra(r)
                x.alignment = Alignment(horizontal="center", vertical="top")
        if kind == "list":
            dv = DataValidation(type="list", formula1="=" + REF[extra], allow_blank=True,
                                showErrorMessage=True, errorStyle="warning",
                                error="Value not in the Lists tab. Add it there to keep reports accurate.")
            ws.add_data_validation(dv); dv.add(f"{c}{first}:{c}{last}")
        elif kind == "listref":
            dv = DataValidation(type="list", formula1=extra[0], allow_blank=True, showErrorMessage=True,
                                errorStyle="warning", error=extra[1])
            ws.add_data_validation(dv); dv.add(f"{c}{first}:{c}{last}")
        elif kind == "date":
            dv = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                                showErrorMessage=True, error="Enter a valid date (mm/dd/yyyy).")
            ws.add_data_validation(dv); dv.add(f"{c}{first}:{c}{last}")
    ws.row_dimensions[4].height = 42
    ws.freeze_panes = freeze
    ws.auto_filter.ref = f"A4:{L(len(cols))}{last}"
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "4:4"
    return first, last

def cf(ws, rng, formula, color):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=fill(color), stopIfTrue=False))


def cf(ws, rng, formula, color):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=fill(color), stopIfTrue=False))
