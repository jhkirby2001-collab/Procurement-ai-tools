"""
Reusable spend-report engine for the NIGP-Sourced Procurement Category Mapper.

Takes a spend file, classifies its descriptions with the mapper's rules
(vectorized, no AI, no API key), then produces a deterministic, leadership-grade
spend report: an executive summary, summary tiles, spend-by-category, Pareto
80/20, top vendors, spend-by-department, and vendor consolidation — rendered as a
professionally formatted, brand-colored Excel workbook with charts.

Pure pandas/openpyxl. Same core is used by the Streamlit "Spend Report" page and
can be run standalone for testing:

    python spend_report_JHK3.py "data/raw/spen data1.csv" \
        --desc "Item Description" --amount "Amount Ordered" \
        --vendor "Vendor Name" --dept "Release Dept Code"
"""
from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Optional

import pandas as pd

import sys
sys.path.insert(0, str(Path(__file__).resolve().parent))
from classifier_JHK3 import (  # noqa: E402
    load_keyword_rules,
    CONFIDENCE_BY_LEVEL,
    CONFIDENCE_LABEL,
    LOWINFO_RE,
)

UNCLASSIFIED_NO_RULE = "Unclassified — No Rule Match"
UNCLASSIFIED_NO_DESC = "Unclassified — No Description"

AMOUNT_HINTS = ["amount", "spend", "cost", "price", "total", "value", "paid",
                "billed", "ordered", "award", "expenditure", "payment"]
VENDOR_HINTS = ["vendor", "supplier", "payee", "merchant", "company", "contractor"]
DEPT_HINTS = ["department name", "dept name", "using area", "department", "dept",
              "agency", "division", "bureau", "office", "unit"]
DESC_HINTS = ["description", "item", "purpose", "commodity", "detail", "line",
              "service", "narrative"]
DATE_HINTS = ["date", "period", "fiscal", "year"]

# City of Chicago brand palette (hex, no '#')
NAVY = "002F6C"
BLUE = "41B6E6"
RED = "DA291C"
LT_BLUE = "D6EEF9"
GRAY = "595959"
WHITE = "FFFFFF"


# ---------------------------------------------------------------------------
# Column detection + currency cleaning
# ---------------------------------------------------------------------------
def _best_col(columns, hints) -> Optional[str]:
    cols = list(columns)
    low = {c: str(c).strip().lower() for c in cols}
    for h in hints:
        for c in cols:
            if low[c] == h:
                return c
    for h in hints:
        for c in cols:
            if h in low[c]:
                return c
    return None


def detect_columns(df: pd.DataFrame) -> dict:
    return {
        "desc": _best_col(df.columns, DESC_HINTS),
        "amount": _best_col(df.columns, AMOUNT_HINTS),
        "vendor": _best_col(df.columns, VENDOR_HINTS),
        "dept": _best_col(df.columns, DEPT_HINTS),
        "date": _best_col(df.columns, DATE_HINTS),
    }


def clean_amount(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return series.astype(float)
    s = series.astype(str).str.strip()
    neg = s.str.startswith("(") & s.str.endswith(")")
    s = s.str.replace(r"[,$\s]", "", regex=True)
    s = s.str.replace(r"^\((.*)\)$", r"\1", regex=True)
    out = pd.to_numeric(s, errors="coerce")
    out = out.mask(neg, -out.abs())
    return out


# ---------------------------------------------------------------------------
# Vectorized classification (Tier 1 keyword rules; no AI, no account codes)
# ---------------------------------------------------------------------------
def classify_series(desc: pd.Series, rules_df=None) -> pd.DataFrame:
    if rules_df is None:
        rules_df = load_keyword_rules()
    desc = desc.fillna("").astype(str).str.strip()
    desc_u = desc.str.upper()
    n = len(desc)
    out = pd.DataFrame({
        "Business_Category": [""] * n,
        "NIGP_Class_3digit": [""] * n,
        "NIGP_Item_5digit": [""] * n,
        "NIGP_Code_Assigned": [""] * n,
        "Classification_Method": [""] * n,
        "Classification_Confidence": [""] * n,
        "Classification_Reason": [""] * n,
    }, index=desc.index)

    unclassified = pd.Series(True, index=desc.index)
    for _, rule in rules_df.iterrows():
        if not unclassified.any():
            break
        pat_u = rule["pattern_upper"]
        mt = rule["match_type"]
        if mt == "exact":
            mask = unclassified & (desc_u == pat_u)
        elif mt == "starts_with":
            mask = unclassified & desc_u.str.startswith(pat_u, na=False)
        else:
            mask = unclassified & desc_u.str.contains(re.escape(pat_u), regex=True, na=False)
        if not mask.any():
            continue
        level = rule["nigp_match_level"]
        score = CONFIDENCE_BY_LEVEL.get(level, 0.50)
        code = str(rule["nigp_item_5digit"]) or str(rule["nigp_class_3digit"])
        out.loc[mask, "Business_Category"] = rule["business_category"]
        out.loc[mask, "NIGP_Class_3digit"] = str(rule["nigp_class_3digit"])
        out.loc[mask, "NIGP_Item_5digit"] = str(rule["nigp_item_5digit"])
        out.loc[mask, "NIGP_Code_Assigned"] = code
        out.loc[mask, "Classification_Method"] = "keyword_rule"
        out.loc[mask, "Classification_Confidence"] = CONFIDENCE_LABEL(score)
        out.loc[mask, "Classification_Reason"] = f"matched '{rule['pattern']}' ({mt})"
        unclassified &= ~mask

    no_desc = unclassified & (desc == "")
    no_rule = unclassified & (desc != "")
    out.loc[no_rule, "Business_Category"] = UNCLASSIFIED_NO_RULE
    out.loc[no_rule, "Classification_Method"] = "no_rule"
    out.loc[no_desc, "Business_Category"] = UNCLASSIFIED_NO_DESC
    out.loc[no_desc, "Classification_Method"] = "no_description"
    return out


# ---------------------------------------------------------------------------
# Analyzers (pure pandas, column-parameterized)
# ---------------------------------------------------------------------------
def _is_classified(cat: pd.Series) -> pd.Series:
    return ~cat.isin([UNCLASSIFIED_NO_RULE, UNCLASSIFIED_NO_DESC])


def summary_tiles(df, category_col, amount_col=None, vendor_col=None,
                  date_col=None) -> dict:
    tiles = {"transactions": int(len(df)),
             "categories": int(df[category_col].nunique())}
    if amount_col and amount_col in df:
        amt = clean_amount(df[amount_col])
        tiles["total_spend"] = float(amt.sum(skipna=True))
        tiles["has_amount"] = True
    else:
        tiles["total_spend"] = None
        tiles["has_amount"] = False
    tiles["vendors"] = int(df[vendor_col].nunique()) if vendor_col and vendor_col in df else None
    if date_col and date_col in df:
        d = pd.to_datetime(df[date_col], errors="coerce")
        if d.notna().any():
            tiles["date_min"] = str(d.min().date())
            tiles["date_max"] = str(d.max().date())
    return tiles


def _rollup(df, group_col, amount_col):
    work = df.copy()
    if amount_col and amount_col in work:
        work["_amt"] = clean_amount(work[amount_col]).fillna(0.0)
        g = work.groupby(group_col)["_amt"].agg(["sum", "count"])
        g.columns = ["Spend", "Transactions"]
        g = g.sort_values("Spend", ascending=False)
        measure = "Spend"
    else:
        g = work.groupby(group_col).size().to_frame("Transactions")
        g = g.sort_values("Transactions", ascending=False)
        measure = "Transactions"
    total = g[measure].sum()
    g["% of Total"] = (g[measure] / total * 100).round(1) if total else 0.0
    return g.reset_index(), measure


def spend_by_category(df, category_col, amount_col=None):
    g, _ = _rollup(df, category_col, amount_col)
    return g.rename(columns={category_col: "Business Category"})


def spend_by_department(df, dept_col, amount_col=None, top_n=25):
    if not dept_col or dept_col not in df:
        return None
    g, _ = _rollup(df, dept_col, amount_col)
    g = g.rename(columns={dept_col: "Department"})
    g["Department"] = g["Department"].astype(str)
    return g.head(top_n) if top_n else g


def pareto(df, group_col, amount_col=None, top_n=None):
    g, measure = _rollup(df, group_col, amount_col)
    g["Cumulative %"] = g["% of Total"].cumsum().round(1)
    n80 = int((g["Cumulative %"] < 80).sum() + 1) if len(g) else 0
    n80 = min(n80, len(g))
    if top_n:
        g = g.head(top_n)
    return g.rename(columns={group_col: "Group"}), measure, n80


def top_vendors(df, vendor_col, amount_col=None, n=15):
    if not vendor_col or vendor_col not in df:
        return None
    g, _ = _rollup(df, vendor_col, amount_col)
    return g.rename(columns={vendor_col: "Vendor"}).head(n)


def consolidation_finder(df, category_col, vendor_col, amount_col=None,
                         dept_col=None, min_vendors=2):
    if not vendor_col or vendor_col not in df:
        return None
    work = df[_is_classified(df[category_col])].copy()
    if amount_col and amount_col in work:
        work["_amt"] = clean_amount(work[amount_col]).fillna(0.0)
        g = work.groupby(category_col).agg(**{
            "Spend": ("_amt", "sum"),
            "Vendors": (vendor_col, "nunique"),
            "Transactions": (vendor_col, "size"),
        })
    else:
        g = work.groupby(category_col).agg(**{
            "Vendors": (vendor_col, "nunique"),
            "Transactions": (vendor_col, "size"),
        })
    if dept_col and dept_col in work:
        g["Departments"] = work.groupby(category_col)[dept_col].nunique()
    g = g[g["Vendors"] >= min_vendors]
    sort_key = "Spend" if "Spend" in g.columns else "Transactions"
    g = g.sort_values(sort_key, ascending=False)
    return g.reset_index().rename(columns={category_col: "Business Category"})


def coverage(df, category_col, method_col=None):
    total = len(df)
    classified = int(_is_classified(df[category_col]).sum())
    no_rule = int((df[category_col] == UNCLASSIFIED_NO_RULE).sum())
    no_desc = int((df[category_col] == UNCLASSIFIED_NO_DESC).sum())
    return {
        "total": total,
        "classified": classified,
        "classified_pct": round(classified / total * 100, 1) if total else 0.0,
        "no_rule": no_rule,
        "no_description": no_desc,
    }


# ---------------------------------------------------------------------------
# Executive summary (deterministic, data-driven)
# ---------------------------------------------------------------------------
def _money_text(v):
    try:
        return f"${float(v):,.0f}"
    except (TypeError, ValueError):
        return "n/a"


def executive_summary(tiles, cat, n80, vend, cons, cov, dept_tbl=None) -> dict:
    has_amt = tiles.get("has_amount")
    unit = "spend" if has_amt else "transactions"
    real = cat[~cat["Business Category"].isin([UNCLASSIFIED_NO_RULE, UNCLASSIFIED_NO_DESC])]
    lines, steps = [], []

    parts = []
    if has_amt:
        parts.append(f"{_money_text(tiles.get('total_spend'))} in spend")
    parts.append(f"{tiles['transactions']:,} transactions")
    if tiles.get("vendors"):
        parts.append(f"{tiles['vendors']:,} vendors")
    if dept_tbl is not None:
        parts.append(f"{len(dept_tbl):,} departments")
    span = f", covering {tiles['date_min']} to {tiles['date_max']}" if "date_min" in tiles else ""
    lines.append("This report analyzes " + ", ".join(parts) + span + ".")

    unclass_pct = round(100 - cov["classified_pct"], 1)
    lines.append(
        f"{cov['classified_pct']}% of rows were automatically classified into commodity "
        f"categories; the remaining {unclass_pct}% could not be matched to a rule and are grouped "
        "as “Unclassified.” Classifying that remainder would widen spend visibility.")

    if len(real):
        top = real.iloc[0]
        val = _money_text(top["Spend"]) if has_amt else f"{int(top['Transactions']):,} transactions"
        lines.append(
            f"Classified {unit} is concentrated: the top {n80} categor"
            f"{'y' if n80 == 1 else 'ies'} drive about 80%. The largest is "
            f"“{top['Business Category']}” at {val} ({top['% of Total']}%).")

    if cons is not None and len(cons):
        frag = cons.head(3)
        names = ", ".join(f"{r['Business Category']} ({int(r['Vendors'])} vendors)"
                          for _, r in frag.iterrows())
        lines.append(
            f"The most fragmented categories — bought from many vendors — are {names}. "
            "These are the clearest consolidation opportunities.")
        steps.append(
            "Target the most fragmented high-spend categories for consolidation ("
            + ", ".join(frag["Business Category"].tolist())
            + "): aggregate demand and reduce the vendor count.")

    if vend is not None and len(vend):
        tv = vend.iloc[0]
        val = _money_text(tv["Spend"]) if has_amt else f"{int(tv['Transactions']):,} transactions"
        lines.append(
            f"The largest single vendor is {tv['Vendor']} at {val} ({tv['% of Total']}% of total).")
        steps.append("Review the top vendors for volume-based pricing, contract compliance, and overlap.")

    if cov["classified_pct"] < 90:
        steps.append(
            f"Reduce the {unclass_pct}% unclassified share by adding keyword rules for the most "
            "common unmatched descriptions — this widens visibility with no new tools.")
    if dept_tbl is not None and len(dept_tbl) > 1:
        steps.append(
            "Coordinate purchasing across departments buying the same commodity to capture "
            "cross-department leverage.")

    return {"lines": lines, "steps": steps}


# ---------------------------------------------------------------------------
# Excel report (professional, brand-colored, charts)
# ---------------------------------------------------------------------------
def _style_data_sheet(ws, title, df, *, money_cols=(), pct_cols=(), int_cols=(),
                      total_cols=(), add_total=True):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ncols = df.shape[1]
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    tcell = ws.cell(row=1, column=1, value=title)
    tcell.font = Font(name="Calibri", size=14, bold=True, color=WHITE)
    tcell.fill = PatternFill("solid", fgColor=NAVY)
    tcell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 24

    header_row, data_first = 2, 3
    data_last = header_row + len(df)
    for c in range(1, ncols + 1):
        cell = ws.cell(row=header_row, column=c)
        cell.font = Font(name="Calibri", bold=True, color=WHITE, size=11)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = border
    for r in range(data_first, data_last + 1):
        band = (r - data_first) % 2 == 1
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            if band:
                cell.fill = PatternFill("solid", fgColor=LT_BLUE)
            colname = df.columns[c - 1]
            if colname in money_cols:
                cell.number_format = '$#,##0'
            elif colname in pct_cols:
                cell.number_format = '0.0"%"'
            elif colname in int_cols:
                cell.number_format = '#,##0'

    if add_total and len(df):
        tr = data_last + 1
        for c in range(1, ncols + 1):
            cell = ws.cell(row=tr, column=c)
            cell.font = Font(bold=True, color=WHITE)
            cell.fill = PatternFill("solid", fgColor=RED)
            cell.border = border
            colname = df.columns[c - 1]
            if c == 1:
                cell.value = "TOTAL"
            elif colname in total_cols:
                col_letter = cell.column_letter
                cell.value = f"=SUM({col_letter}{data_first}:{col_letter}{data_last})"
                if colname in money_cols:
                    cell.number_format = '$#,##0'
                elif colname in int_cols:
                    cell.number_format = '#,##0'
            elif colname in pct_cols:
                cell.value = 100.0
                cell.number_format = '0.0"%"'

    for c in range(1, ncols + 1):
        maxlen = max([len(str(df.columns[c - 1]))] +
                     [len(str(df.iloc[r, c - 1])) for r in range(len(df))], default=12)
        ws.column_dimensions[ws.cell(row=header_row, column=c).column_letter].width = min(max(maxlen + 3, 12), 46)
    ws.freeze_panes = ws.cell(row=data_first, column=1)
    ws.sheet_view.showGridLines = False
    return header_row, data_first, data_last


def _bar_chart(ws, title, cat_col, val_col, header_row, data_first, data_last, anchor):
    from openpyxl.chart import BarChart, Reference
    ch = BarChart()
    ch.type = "col"
    ch.title = title
    ch.height = 8.5
    ch.width = 17
    ch.style = 10
    data = Reference(ws, min_col=val_col, min_row=header_row, max_row=data_last)
    cats = Reference(ws, min_col=cat_col, min_row=data_first, max_row=data_last)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    ch.legend = None
    ws.add_chart(ch, anchor)


def _pareto_chart(ws, cat_col, spend_col, cum_col, header_row, data_first, data_last, anchor):
    from openpyxl.chart import BarChart, LineChart, Reference
    bar = BarChart()
    bar.type = "col"
    bar.title = "Pareto — spend and cumulative %"
    bar.height = 9
    bar.width = 18
    bar.style = 10
    d = Reference(ws, min_col=spend_col, min_row=header_row, max_row=data_last)
    cats = Reference(ws, min_col=cat_col, min_row=data_first, max_row=data_last)
    bar.add_data(d, titles_from_data=True)
    bar.set_categories(cats)
    bar.legend = None
    line = LineChart()
    dl = Reference(ws, min_col=cum_col, min_row=header_row, max_row=data_last)
    line.add_data(dl, titles_from_data=True)
    line.y_axis.axId = 200
    line.y_axis.title = "Cumulative %"
    line.y_axis.crosses = "max"
    bar += line
    ws.add_chart(bar, anchor)


def _write_exec_summary(ws, title, tiles, cov, es):
    from openpyxl.styles import Font, PatternFill, Alignment
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2.5
    for col in ["B", "C", "D", "E", "F", "G"]:
        ws.column_dimensions[col].width = 19

    ws.merge_cells("B2:G2")
    t = ws["B2"]
    t.value = title
    t.font = Font(size=18, bold=True, color=WHITE)
    t.fill = PatternFill("solid", fgColor=NAVY)
    t.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[2].height = 34
    ws.merge_cells("B3:G3")
    sub = ws["B3"]
    subtxt = "Procurement Spend Analysis"
    if "date_min" in tiles:
        subtxt += f"   |   {tiles['date_min']} to {tiles['date_max']}"
    sub.value = subtxt
    sub.font = Font(size=11, italic=True, color=NAVY)

    metrics = []
    if tiles.get("has_amount"):
        metrics.append(("Total Spend", _money_text(tiles.get("total_spend"))))
    metrics.append(("Transactions", f"{tiles['transactions']:,}"))
    metrics.append(("Categories", f"{tiles['categories']:,}"))
    if tiles.get("vendors"):
        metrics.append(("Vendors", f"{tiles['vendors']:,}"))
    metrics.append(("Classified", f"{cov['classified_pct']}%"))
    col = 2
    for label, val in metrics:
        v = ws.cell(row=5, column=col, value=val)
        v.font = Font(size=15, bold=True, color=RED)
        v.alignment = Alignment(horizontal="center")
        v.fill = PatternFill("solid", fgColor=LT_BLUE)
        d = ws.cell(row=6, column=col, value=label)
        d.font = Font(size=10, color=NAVY)
        d.alignment = Alignment(horizontal="center")
        d.fill = PatternFill("solid", fgColor=LT_BLUE)
        col += 1

    row = 8
    h = ws.cell(row=row, column=2, value="What the data shows")
    h.font = Font(size=13, bold=True, color=NAVY)
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=7)
    row += 1
    for line in (es["lines"] if es else []):
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=7)
        c = ws.cell(row=row, column=2, value="•  " + line)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.font = Font(size=11)
        ws.row_dimensions[row].height = 32
        row += 1

    row += 1
    h2 = ws.cell(row=row, column=2, value="Recommended first steps")
    h2.font = Font(size=13, bold=True, color=NAVY)
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=7)
    row += 1
    for i, step in enumerate(es["steps"] if es else [], 1):
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=7)
        c = ws.cell(row=row, column=2, value=f"{i}.  {step}")
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.font = Font(size=11)
        ws.row_dimensions[row].height = 32
        row += 1

    row += 1
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=7)
    fn = ws.cell(row=row, column=2, value=(
        "Note: “Unclassified — No Rule Match” means the description did not match any "
        "keyword rule, so no category was assigned (not an error). Adding rules for recurring "
        "items reduces it. All figures are deterministic — no AI is used in the analysis."))
    fn.font = Font(size=9, italic=True, color=GRAY)
    fn.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[row].height = 46


def build_excel_report(path, *, tiles, cat_tbl, pareto_tbl, vendors_tbl,
                       consolidation_tbl, cov, exec_summary=None, dept_tbl=None,
                       report_title="Procurement Spend Analysis") -> str:
    has_amt = tiles.get("has_amount")
    measure = "Spend" if has_amt else "Transactions"
    money = ["Spend"] if has_amt else []

    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        cat_tbl.to_excel(xl, sheet_name="Spend by Category", index=False, startrow=1)
        pareto_tbl.to_excel(xl, sheet_name="Pareto 80-20", index=False, startrow=1)
        if vendors_tbl is not None:
            vendors_tbl.to_excel(xl, sheet_name="Top Vendors", index=False, startrow=1)
        if dept_tbl is not None:
            dept_tbl.to_excel(xl, sheet_name="Spend by Department", index=False, startrow=1)
        if consolidation_tbl is not None:
            consolidation_tbl.to_excel(xl, sheet_name="Consolidation", index=False, startrow=1)
        wb = xl.book

        es = wb.create_sheet("Executive Summary", 0)
        _write_exec_summary(es, report_title, tiles, cov, exec_summary)

        def col_idx(df, name):
            return list(df.columns).index(name) + 1

        # Spend by Category
        ws = wb["Spend by Category"]
        hr, df1, dl = _style_data_sheet(
            ws, "Spend by Business Category", cat_tbl,
            money_cols=money, pct_cols=["% of Total"], int_cols=["Transactions"],
            total_cols=[measure] + (["Transactions"] if has_amt else []))
        _bar_chart(ws, f"{measure} by Category", 1, col_idx(cat_tbl, measure),
                   hr, df1, dl, ws.cell(row=2, column=cat_tbl.shape[1] + 2).coordinate)

        # Pareto
        ws = wb["Pareto 80-20"]
        hr, df1, dl = _style_data_sheet(
            ws, "Pareto 80/20 (classified spend)", pareto_tbl,
            money_cols=money, pct_cols=["% of Total", "Cumulative %"],
            int_cols=["Transactions"], add_total=False)
        _pareto_chart(ws, 1, col_idx(pareto_tbl, measure), col_idx(pareto_tbl, "Cumulative %"),
                      hr, df1, dl, ws.cell(row=2, column=pareto_tbl.shape[1] + 2).coordinate)

        # Top Vendors
        if vendors_tbl is not None:
            ws = wb["Top Vendors"]
            hr, df1, dl = _style_data_sheet(
                ws, "Top Vendors by Spend", vendors_tbl,
                money_cols=money, pct_cols=["% of Total"], int_cols=["Transactions"],
                total_cols=[measure] + (["Transactions"] if has_amt else []))
            _bar_chart(ws, f"Top Vendors by {measure}", 1, col_idx(vendors_tbl, measure),
                       hr, df1, dl, ws.cell(row=2, column=vendors_tbl.shape[1] + 2).coordinate)

        # Spend by Department
        if dept_tbl is not None:
            ws = wb["Spend by Department"]
            hr, df1, dl = _style_data_sheet(
                ws, "Spend by Department", dept_tbl,
                money_cols=money, pct_cols=["% of Total"], int_cols=["Transactions"],
                total_cols=[measure] + (["Transactions"] if has_amt else []))
            _bar_chart(ws, f"{measure} by Department", 1, col_idx(dept_tbl, measure),
                       hr, df1, dl, ws.cell(row=2, column=dept_tbl.shape[1] + 2).coordinate)

        # Consolidation
        if consolidation_tbl is not None:
            ws = wb["Consolidation"]
            int_c = [c for c in ["Vendors", "Transactions", "Departments"] if c in consolidation_tbl.columns]
            _style_data_sheet(
                ws, "Vendor Consolidation / Fragmentation", consolidation_tbl,
                money_cols=money, int_cols=int_c,
                total_cols=(["Spend", "Transactions"] if has_amt else ["Transactions"]))

    return path


# ---------------------------------------------------------------------------
# Standalone runner (for testing)
# ---------------------------------------------------------------------------
def run(path, desc=None, amount=None, vendor=None, dept=None, out=None, sample=None):
    if str(path).lower().endswith(".csv"):
        try:
            df = pd.read_csv(path, low_memory=False)
        except UnicodeDecodeError:
            df = pd.read_csv(path, low_memory=False, encoding="latin-1")
    else:
        df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]
    if sample:
        df = df.head(int(sample))
    guess = detect_columns(df)
    desc = desc or guess["desc"]
    amount = amount or guess["amount"]
    vendor = vendor or guess["vendor"]
    dept = dept or guess["dept"]
    print(f"Columns -> desc={desc!r} amount={amount!r} vendor={vendor!r} dept={dept!r}")
    print(f"Classifying {len(df):,} rows ...")
    cls = classify_series(df[desc])
    df = pd.concat([df.reset_index(drop=True), cls.reset_index(drop=True)], axis=1)

    tiles = summary_tiles(df, "Business_Category", amount, vendor, guess["date"])
    cov = coverage(df, "Business_Category")
    cat = spend_by_category(df, "Business_Category", amount)
    classified_df = df[_is_classified(df["Business_Category"])]
    par, measure, n80 = pareto(classified_df, "Business_Category", amount)
    vend = top_vendors(df, vendor, amount)
    deptt = spend_by_department(df, dept, amount)
    cons = consolidation_finder(df, "Business_Category", vendor, amount, dept)
    es = executive_summary(tiles, cat, n80, vend, cons, cov, deptt)

    print("\n=== EXECUTIVE SUMMARY ===")
    for ln in es["lines"]:
        print("  • " + ln)
    print("  Recommended steps:")
    for i, s in enumerate(es["steps"], 1):
        print(f"   {i}. {s}")
    print(f"\n=== COVERAGE === classified {cov['classified']:,}/{cov['total']:,} "
          f"({cov['classified_pct']}%)")
    print(f"\n=== SPEND BY CATEGORY (top 8) ===\n{cat.head(8).to_string(index=False)}")

    out = out or (Path(path).with_suffix("").as_posix() + "_Spend_Report_JHK3.xlsx")
    build_excel_report(out, tiles=tiles, cat_tbl=cat, pareto_tbl=par, vendors_tbl=vend,
                       consolidation_tbl=cons, cov=cov, exec_summary=es, dept_tbl=deptt)
    print(f"\nWrote Excel report: {out}")
    return out


def main():
    p = argparse.ArgumentParser()
    p.add_argument("path")
    p.add_argument("--desc"); p.add_argument("--amount")
    p.add_argument("--vendor"); p.add_argument("--dept")
    p.add_argument("--out"); p.add_argument("--sample", type=int)
    a = p.parse_args()
    run(a.path, a.desc, a.amount, a.vendor, a.dept, a.out, a.sample)


if __name__ == "__main__":
    main()
