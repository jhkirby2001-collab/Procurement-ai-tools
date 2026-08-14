"""
Reusable spend-report engine for the NIGP-Sourced Procurement Category Mapper.

Takes a spend file, classifies its descriptions with the mapper's rules
(vectorized, no AI, no API key), then produces a deterministic spend report:
summary tiles, spend-by-category, Pareto 80/20, top vendors, vendor
consolidation/fragmentation, and a coverage note — plus a formatted Excel
workbook.

Pure pandas/openpyxl. Same core is used by the Streamlit "Spend Report" page
and can be run standalone for testing:

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

# Reuse the mapper's rule loader + confidence helpers so classification here
# matches the classifier exactly.
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

# Column auto-detection hints (case-insensitive substring match on column names)
AMOUNT_HINTS = ["amount", "spend", "cost", "price", "total", "value", "paid",
                "billed", "ordered", "award", "expenditure", "payment"]
VENDOR_HINTS = ["vendor", "supplier", "payee", "merchant", "company", "contractor"]
DEPT_HINTS = ["dept", "department", "agency", "division", "using area", "bureau",
              "office", "unit"]
DESC_HINTS = ["description", "item", "purpose", "commodity", "detail", "line",
              "service", "narrative"]
DATE_HINTS = ["date", "period", "fiscal", "year"]


# ---------------------------------------------------------------------------
# Column detection + currency cleaning
# ---------------------------------------------------------------------------
def _best_col(columns, hints) -> Optional[str]:
    cols = list(columns)
    low = {c: str(c).strip().lower() for c in cols}
    # exact-ish first, then substring by hint order
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
    """Best-guess description/amount/vendor/dept/date columns from names."""
    return {
        "desc": _best_col(df.columns, DESC_HINTS),
        "amount": _best_col(df.columns, AMOUNT_HINTS),
        "vendor": _best_col(df.columns, VENDOR_HINTS),
        "dept": _best_col(df.columns, DEPT_HINTS),
        "date": _best_col(df.columns, DATE_HINTS),
    }


def clean_amount(series: pd.Series) -> pd.Series:
    """Coerce a currency-ish column to float. Handles $, commas, and
    parenthesized negatives. Non-parseable values become NaN."""
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
    """Classify a Series of descriptions using the mapper's keyword rules.

    Vectorized: iterates rules (already sorted exact>starts_with>contains,
    curated before AI) and assigns each rule only to still-unclassified rows,
    so first-match-wins matches the production classifier. Returns a frame of
    classification columns aligned to `desc.index`.
    """
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
    lowinfo = desc.apply(lambda s: bool(s) and bool(LOWINFO_RE.match(s)))

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
    """Group -> (Spend, Transactions), sorted by the primary measure desc."""
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


def pareto(df, group_col, amount_col=None, top_n=None):
    """Returns (table with Cumulative %, n_groups_to_80). 80/20 view."""
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
    """Categories bought from >1 vendor — the fragmentation/consolidation view.
    Sorted by spend so the biggest consolidation opportunities surface first."""
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
# Excel report
# ---------------------------------------------------------------------------
NAVY = "002F6C"
LT_BLUE = "D6EEF9"


def build_excel_report(path, *, tiles, cat_tbl, pareto_tbl, vendors_tbl,
                       consolidation_tbl, cov) -> str:
    from openpyxl.styles import Font, PatternFill, Alignment
    # `path` may be a filename or an in-memory buffer (BytesIO) for web download.
    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        # Summary tab
        summ_rows = [
            ["Total spend", tiles.get("total_spend")],
            ["Transactions", tiles.get("transactions")],
            ["Business categories", tiles.get("categories")],
            ["Vendors", tiles.get("vendors")],
            ["Classified", f"{cov['classified']:,} of {cov['total']:,} ({cov['classified_pct']}%)"],
        ]
        if "date_min" in tiles:
            summ_rows.append(["Date range", f"{tiles['date_min']} to {tiles['date_max']}"])
        pd.DataFrame(summ_rows, columns=["Metric", "Value"]).to_excel(
            xl, sheet_name="Summary", index=False)
        cat_tbl.to_excel(xl, sheet_name="Spend by Category", index=False)
        pareto_tbl.to_excel(xl, sheet_name="Pareto 80-20", index=False)
        if vendors_tbl is not None:
            vendors_tbl.to_excel(xl, sheet_name="Top Vendors", index=False)
        if consolidation_tbl is not None:
            consolidation_tbl.to_excel(xl, sheet_name="Consolidation", index=False)

        navy = PatternFill("solid", fgColor=NAVY)
        head_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        for ws in xl.book.worksheets:
            for cell in ws[1]:
                cell.fill = navy
                cell.font = head_font
                cell.alignment = Alignment(vertical="center")
            for col in ws.columns:
                width = max((len(str(c.value)) for c in col if c.value is not None), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 12), 48)
            ws.freeze_panes = "A2"
    return path


# ---------------------------------------------------------------------------
# Standalone runner (for testing)
# ---------------------------------------------------------------------------
def run(path, desc=None, amount=None, vendor=None, dept=None, out=None,
        sample=None):
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
    cat = spend_by_category(df, "Business_Category", amount)
    classified_df = df[_is_classified(df["Business_Category"])]
    par, measure, n80 = pareto(classified_df, "Business_Category", amount)
    vend = top_vendors(df, vendor, amount)
    cons = consolidation_finder(df, "Business_Category", vendor, amount, dept)
    cov = coverage(df, "Business_Category")

    print("\n=== SUMMARY ===")
    for k, v in tiles.items():
        print(f"  {k:16s} {v}")
    print(f"\n=== COVERAGE === classified {cov['classified']:,}/{cov['total']:,} "
          f"({cov['classified_pct']}%), no-rule {cov['no_rule']:,}, no-desc {cov['no_description']:,}")
    print(f"\n=== SPEND BY CATEGORY (top 10 of {len(cat)}) ===")
    print(cat.head(10).to_string(index=False))
    print(f"\n=== PARETO ({measure}): {n80} categories drive 80% ===")
    print(par.head(n80).to_string(index=False))
    if vend is not None:
        print(f"\n=== TOP VENDORS ===\n{vend.head(10).to_string(index=False)}")
    if cons is not None:
        print(f"\n=== CONSOLIDATION (categories bought from >1 vendor, top 10) ===")
        print(cons.head(10).to_string(index=False))

    out = out or (Path(path).with_suffix("").as_posix() + "_Spend_Report_JHK3.xlsx")
    build_excel_report(out, tiles=tiles, cat_tbl=cat, pareto_tbl=par,
                       vendors_tbl=vend, consolidation_tbl=cons, cov=cov)
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
