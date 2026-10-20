"""
Objective-driven analysis for the NIGP-Sourced Procurement Category Mapper.

The reader states an objective in plain language ("which treatment-chemical
contracts expire before June, and who holds them?"). The design keeps every
number defensible:

1. **Claude plans** - it reads the objective plus a *profile* of the data (column
   names, types, a few sample values, row count; never the full file) and returns
   a plan made only of the operations below, with its assumptions and any
   questions it needs answered.
2. **The reader approves** the plan, shown step by step in plain English.
3. **This module calculates** - each step is ordinary pandas, run in order. No AI
   touches a number, so the same plan on the same data always gives the same
   answer, and every figure traces to a shown step.
4. **Claude writes the summary** from the finished result table only.

When an objective needs an operation that isn't in the set, the plan says so in
`cannot_do` instead of improvising.

James is the only user and authorized AI use (2026-10-20). Classification itself
stays rules-only; this module needs ANTHROPIC_API_KEY only for steps 1 and 4.
"""
from __future__ import annotations

import io
import json
from datetime import date

import pandas as pd

import spend_report_JHK3 as sr

MODEL = "claude-opus-5-5"
FALLBACK_BETA = "server-side-fallback-2026-07-01"
SAMPLE_VALUES = 6
MAX_PROFILE_COLS = 80
RESULT_ROWS_FOR_SUMMARY = 40

# --------------------------------------------------------------------------
# The operation set. Every plan step uses one of these.
# --------------------------------------------------------------------------
OPS = {
    "filter": "Keep rows where `column` meets `operator` and `value`/`values`. Operators: "
              "equals, not_equals, contains, not_contains, in, not_in, gt, gte, lt, lte, between, "
              "is_blank, not_blank. Text comparisons ignore case. Numbers and dates are parsed "
              "automatically (dates as YYYY-MM-DD).",
    "days_until": "Add column `as_name` = whole days from `today` to the date in `column` "
                  "(negative = already past). Use for expirations.",
    "bucket": "Add column `as_name` sorting the number in `column` into ranges. `edges` are the "
              "boundaries (ascending); `labels` has one fewer item than edges.",
    "group_summarize": "Group by `columns` and compute `metrics`. Each metric names a source column, "
                       "an agg (sum, count, mean, min, max, nunique) and an output name.",
    "add_share": "Add column `as_name` = each row's share (%) of the total of `column`.",
    "sort": "Sort by `column` (descending when `descending` is true).",
    "top_n": "Keep the first `n` rows.",
    "select_columns": "Keep only `columns`, in that order.",
    "rename_columns": "Rename columns: `columns` are the old names, `values` the new names, in order.",
}

_STEP_SCHEMA = {
    "type": "object",
    "properties": {
        "op": {"type": "string", "enum": list(OPS)},
        "column": {"type": "string"},
        "columns": {"type": "array", "items": {"type": "string"}},
        "operator": {"type": "string"},
        "value": {"type": "string"},
        "values": {"type": "array", "items": {"type": "string"}},
        "as_name": {"type": "string"},
        "edges": {"type": "array", "items": {"type": "number"}},
        "labels": {"type": "array", "items": {"type": "string"}},
        "metrics": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "column": {"type": "string"},
                    "agg": {"type": "string", "enum": ["sum", "count", "mean", "min", "max", "nunique"]},
                    "as_name": {"type": "string"},
                },
                "required": ["column", "agg", "as_name"],
                "additionalProperties": False,
            },
        },
        "descending": {"type": "boolean"},
        "n": {"type": "integer"},
        "why": {"type": "string"},
    },
    "required": ["op", "column", "columns", "operator", "value", "values", "as_name", "edges",
                 "labels", "metrics", "descending", "n", "why"],
    "additionalProperties": False,
}

PLAN_SCHEMA = {
    "type": "object",
    "properties": {
        "restated_objective": {"type": "string"},
        "assumptions": {"type": "array", "items": {"type": "string"}},
        "questions": {"type": "array", "items": {"type": "string"}},
        "cannot_do": {"type": "array", "items": {"type": "string"}},
        "steps": {"type": "array", "items": _STEP_SCHEMA},
        "result_title": {"type": "string"},
    },
    "required": ["restated_objective", "assumptions", "questions", "cannot_do", "steps", "result_title"],
    "additionalProperties": False,
}

PLANNER_SYSTEM = """You plan analyses of procurement data for a public-sector contracts administrator \
(a city water department). You never calculate anything yourself: you write a plan that a separate, \
deterministic program runs, and the administrator must be able to defend every number to leadership.

Rules:
- Use only the operations listed. Use only column names that appear in the data profile, or names \
created by an earlier step's as_name / metric as_name.
- Every step must fill every field. Unused text fields are "", unused lists are [], unused numbers \
are 0, unused booleans are false. Give each step a short plain-English `why`.
- Restate the objective in one sentence as you understood it.
- List every assumption you made (for example which column holds the end date, what "this year" \
means, what counts as a chemical). Prefer columns over guesses.
- If the objective is ambiguous in a way that would change the answer, put the question in \
`questions` and still give your best plan.
- If part of the objective needs an operation, a column or a second file that isn't available, \
say so plainly in `cannot_do` and plan the part that is possible. Never invent data.
- End with a table a leader can read: few columns, clear names, sorted with the most important \
rows first.
- Columns Business_Category, NIGP_Class_3digit, NIGP_Item_5digit, NAICS_Code and NAICS_Title, \
when present, come from the mapper's rules-based classification of the description column."""

SUMMARY_SYSTEM = """You write the opening of a briefing for a public-sector contracts administrator. \
You are given the objective, the steps that were run, and the result table. Lead with the direct \
answer in one or two sentences, using numbers from the table exactly as they appear. Then give at \
most three short supporting points. Do not add numbers that are not in the table, do not speculate \
about causes, and mention any caveat listed. Plain language, no headings, no jargon."""


# --------------------------------------------------------------------------
# Data profile (what the planner sees instead of the data)
# --------------------------------------------------------------------------
def profile_for_planner(df: pd.DataFrame) -> dict:
    cols = []
    for c in list(df.columns)[:MAX_PROFILE_COLS]:
        s = df[c]
        nn = s.dropna()
        if pd.api.types.is_numeric_dtype(s):
            kind = "number"
        elif sr._date_frac(s) >= 0.7:
            kind = "date"
        elif sr._num_frac(s) >= 0.8:
            kind = "number (stored as text)"
        else:
            kind = "text"
        samples = [str(v)[:60] for v in nn.astype(str).drop_duplicates().head(SAMPLE_VALUES)]
        cols.append({"name": str(c), "kind": kind, "non_blank": int(len(nn)),
                     "distinct": int(nn.nunique()), "examples": samples})
    return {"rows": int(len(df)), "columns": cols, "today": date.today().isoformat()}


# --------------------------------------------------------------------------
# Claude calls
# --------------------------------------------------------------------------
class AIUnavailable(RuntimeError):
    pass


def _client(api_key: str):
    if not api_key:
        raise AIUnavailable("No Anthropic API key is set. Add ANTHROPIC_API_KEY to the app's secrets.")
    import anthropic
    return anthropic.Anthropic(api_key=api_key)


def _text_of(response) -> str:
    if response.stop_reason == "refusal":
        raise AIUnavailable("The request was declined. Rephrase the objective and try again.")
    return next((b.text for b in response.content if b.type == "text"), "")


def plan_objective(objective: str, profile: dict, api_key: str, answers: str = "") -> dict:
    """Ask Claude for a plan. Returns the plan dict (validated by the API's schema)."""
    client = _client(api_key)
    ops = "\n".join(f"- {k}: {v}" for k, v in OPS.items())
    user = (f"Operations available:\n{ops}\n\nData profile (JSON):\n{json.dumps(profile, indent=1)}\n\n"
            f"Objective:\n{objective.strip()}")
    if answers.strip():
        user += f"\n\nThe administrator answered your earlier questions:\n{answers.strip()}"
    resp = client.beta.messages.create(
        model=MODEL,
        max_tokens=16000,
        betas=[FALLBACK_BETA],
        fallbacks="default",
        output_config={"effort": "high", "format": {"type": "json_schema", "schema": PLAN_SCHEMA}},
        system=PLANNER_SYSTEM,
        messages=[{"role": "user", "content": user}],
    )
    return json.loads(_text_of(resp))


def summarize_result(objective: str, plan: dict, result: pd.DataFrame, caveats: list,
                     api_key: str) -> str:
    client = _client(api_key)
    table = result.head(RESULT_ROWS_FOR_SUMMARY).to_csv(index=False)
    more = f"\n(The full result has {len(result):,} rows; the first {RESULT_ROWS_FOR_SUMMARY} are shown.)" \
        if len(result) > RESULT_ROWS_FOR_SUMMARY else ""
    user = (f"Objective: {objective}\n\nSteps run:\n" + "\n".join(f"- {s}" for s in describe_steps(plan))
            + "\n\nCaveats:\n" + ("\n".join(f"- {c}" for c in caveats) or "- none")
            + f"\n\nResult table (CSV):\n{table}{more}")
    resp = client.beta.messages.create(
        model=MODEL,
        max_tokens=4000,
        betas=[FALLBACK_BETA],
        fallbacks="default",
        output_config={"effort": "low"},
        system=SUMMARY_SYSTEM,
        messages=[{"role": "user", "content": user}],
    )
    return _text_of(resp).strip()


# --------------------------------------------------------------------------
# Plain-English description of a plan
# --------------------------------------------------------------------------
_OP_WORDS = {"equals": "is", "not_equals": "is not", "contains": "contains", "not_contains": "does not contain",
             "in": "is one of", "not_in": "is not one of", "gt": "is more than", "gte": "is at least",
             "lt": "is less than", "lte": "is at most", "between": "is between",
             "is_blank": "is blank", "not_blank": "is not blank"}


def describe_step(s: dict) -> str:
    op = s.get("op")
    if op == "filter":
        val = " and ".join(s.get("values") or []) if s.get("operator") == "between" \
            else ", ".join(s.get("values") or []) or s.get("value", "")
        cond = _OP_WORDS.get(s.get("operator"), s.get("operator"))
        return f"Keep rows where **{s['column']}** {cond}" + (f" **{val}**" if val and "blank" not in s.get("operator", "") else "")
    if op == "days_until":
        return f"Count days from today to **{s['column']}** → new column **{s['as_name']}**"
    if op == "bucket":
        return f"Sort **{s['column']}** into ranges {', '.join(s.get('labels') or [])} → **{s['as_name']}**"
    if op == "group_summarize":
        ms = "; ".join(f"{m['agg']} of {m['column']} as **{m['as_name']}**" for m in s.get("metrics") or [])
        return f"Group by **{', '.join(s.get('columns') or [])}** and compute {ms}"
    if op == "add_share":
        return f"Add each row's % of total **{s['column']}** → **{s['as_name']}**"
    if op == "sort":
        return f"Sort by **{s['column']}** ({'largest first' if s.get('descending') else 'smallest first'})"
    if op == "top_n":
        return f"Keep the top **{s.get('n')}** rows"
    if op == "select_columns":
        return f"Show columns **{', '.join(s.get('columns') or [])}**"
    if op == "rename_columns":
        return "Rename " + ", ".join(f"{a} → {b}" for a, b in zip(s.get("columns") or [], s.get("values") or []))
    return f"Unknown step {op}"


def describe_steps(plan: dict) -> list:
    return [describe_step(s) for s in plan.get("steps", [])]


# --------------------------------------------------------------------------
# Deterministic execution
# --------------------------------------------------------------------------
class PlanError(ValueError):
    pass


def _need(work, col, step_no):
    if col not in work.columns:
        raise PlanError(f"Step {step_no}: there is no column named '{col}'. "
                        f"Available: {', '.join(map(str, work.columns))}")


def _as_number(series):
    return sr.clean_amount(series)


def _as_date(series):
    return pd.to_datetime(series, errors="coerce")


def _is_dateish(series):
    return pd.api.types.is_datetime64_any_dtype(series) or sr._date_frac(series) >= 0.7


def _compare(series, operator, value, values):
    """Build a boolean mask; compares as numbers, dates or case-insensitive text."""
    if operator == "is_blank":
        return series.isna() | (series.astype(str).str.strip() == "")
    if operator == "not_blank":
        return ~(series.isna() | (series.astype(str).str.strip() == ""))
    text = series.astype(str).str.strip().str.lower()
    if operator in ("contains", "not_contains"):
        m = text.str.contains(str(value).strip().lower(), regex=False, na=False)
        return m if operator == "contains" else ~m
    if operator in ("in", "not_in"):
        opts = {str(v).strip().lower() for v in (values or [value])}
        m = text.isin(opts)
        return m if operator == "in" else ~m
    if operator in ("equals", "not_equals") and not _looks_numeric_or_date(value):
        m = text == str(value).strip().lower()
        return m if operator == "equals" else ~m
    # numeric or date comparison
    if _is_dateish(series) and _looks_date(value if operator != "between" else (values or [""])[0]):
        left = _as_date(series)
        conv = lambda v: pd.Timestamp(v)  # noqa: E731
    else:
        left = _as_number(series)
        conv = lambda v: float(sr.clean_amount(pd.Series([v])).iloc[0])  # noqa: E731
    if operator == "between":
        if not values or len(values) != 2:
            raise PlanError("A 'between' filter needs exactly two values.")
        lo, hi = conv(values[0]), conv(values[1])
        return (left >= lo) & (left <= hi)
    v = conv(value)
    return {"equals": left == v, "not_equals": left != v, "gt": left > v, "gte": left >= v,
            "lt": left < v, "lte": left <= v}[operator]


def _looks_date(v):
    try:
        pd.Timestamp(str(v))
        return len(str(v)) >= 8 and "-" in str(v)
    except (ValueError, TypeError):
        return False


def _looks_numeric_or_date(v):
    if _looks_date(v):
        return True
    return pd.notna(sr.clean_amount(pd.Series([v])).iloc[0])


def execute_plan(df: pd.DataFrame, plan: dict, today: date | None = None):
    """Run the plan. Returns (result DataFrame, list of caveats)."""
    today = pd.Timestamp(today or date.today())
    work = df.copy()
    caveats = []
    for i, s in enumerate(plan.get("steps", []), 1):
        op = s.get("op")
        before = len(work)
        if op == "filter":
            _need(work, s["column"], i)
            if s["operator"] not in _OP_WORDS:
                raise PlanError(f"Step {i}: unknown filter operator '{s['operator']}'.")
            col = work[s["column"]]
            work = work[_compare(col, s["operator"], s.get("value", ""), s.get("values"))]
            if s["operator"] in ("gt", "gte", "lt", "lte", "between"):
                parsed = _as_date(col) if _is_dateish(col) else _as_number(col)
                blanks = int(parsed.isna().sum())
                if blanks:
                    caveats.append(f"{blanks:,} rows had no usable value in '{s['column']}' "
                                   f"and were left out at step {i}.")
        elif op == "days_until":
            _need(work, s["column"], i)
            d = _as_date(work[s["column"]])
            work[s["as_name"]] = (d - today).dt.days
            if d.isna().any():
                caveats.append(f"{int(d.isna().sum()):,} rows have no usable date in '{s['column']}'.")
        elif op == "bucket":
            _need(work, s["column"], i)
            edges, labels = s.get("edges") or [], s.get("labels") or []
            if len(labels) != len(edges) - 1:
                raise PlanError(f"Step {i}: a bucket needs one fewer label than edges.")
            work[s["as_name"]] = pd.cut(_as_number(work[s["column"]]), bins=edges, labels=labels,
                                        include_lowest=True).astype(str).replace("nan", "")
        elif op == "group_summarize":
            by = s.get("columns") or []
            for c in by:
                _need(work, c, i)
            g = work.copy()
            aggs = {}
            for m in s.get("metrics") or []:
                src = m["column"]
                if m["agg"] == "count" and src not in g.columns:
                    src = by[0] if by else g.columns[0]
                _need(g, src, i)
                tmp = f"__{m['as_name']}"
                g[tmp] = g[src] if m["agg"] in ("count", "nunique") else _as_number(g[src])
                aggs[m["as_name"]] = (tmp, m["agg"])
            if not aggs:
                aggs = {"Rows": (by[0], "count")}
            work = (g.groupby(by, dropna=False).agg(**aggs).reset_index() if by
                    else g.agg(**aggs).T.reset_index(drop=True))
            for name, (_, agg) in aggs.items():
                if agg == "sum":
                    work[name] = work[name].round(2)
        elif op == "add_share":
            _need(work, s["column"], i)
            v = _as_number(work[s["column"]])
            total = v.sum()
            work[s["as_name"]] = (v / total * 100).round(1) if total else 0.0
        elif op == "sort":
            _need(work, s["column"], i)
            key = work[s["column"]]
            if _is_dateish(key) and not pd.api.types.is_numeric_dtype(key):
                key = _as_date(key)
            elif sr._num_frac(key) >= 0.8:
                key = _as_number(key)
            work = work.loc[key.sort_values(ascending=not s.get("descending"), na_position="last").index]
        elif op == "top_n":
            work = work.head(max(int(s.get("n") or 0), 0))
        elif op == "select_columns":
            for c in s.get("columns") or []:
                _need(work, c, i)
            work = work[list(s["columns"])]
        elif op == "rename_columns":
            work = work.rename(columns=dict(zip(s.get("columns") or [], s.get("values") or [])))
        else:
            raise PlanError(f"Step {i}: '{op}' is not an available operation.")
        if op == "filter" and len(work) == 0 and before:
            caveats.append(f"Step {i} left no rows — check the condition.")
    return work.reset_index(drop=True), caveats


# --------------------------------------------------------------------------
# Excel export
# --------------------------------------------------------------------------
def excel_bytes(objective: str, plan: dict, result: pd.DataFrame, caveats: list, summary: str = "",
                source_name: str = "") -> bytes:
    from openpyxl.styles import Alignment, Font, PatternFill
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        pd.DataFrame({"x": []}).to_excel(xl, sheet_name="Answer", index=False)
        ws = xl.sheets["Answer"]
        ws.delete_rows(1)
        ws.column_dimensions["A"].width = 24
        ws.column_dimensions["B"].width = 100
        rows = [("Objective", objective), ("As understood", plan.get("restated_objective", "")),
                ("Data", source_name), ("Answer", summary or "(summary not generated)")]
        rows += [("Assumption", a) for a in plan.get("assumptions", [])]
        rows += [(f"Step {n}", t.replace("**", "")) for n, t in enumerate(describe_steps(plan), 1)]
        rows += [("Caveat", c) for c in caveats]
        rows += [("Not covered", c) for c in plan.get("cannot_do", [])]
        rows += [("Method", "Steps were planned with AI from the objective and the column list, approved, "
                            "then calculated with ordinary formulas. No AI produced any number. "
                            "AI-assisted draft — verify before relying on it.")]
        for r, (k, v) in enumerate(rows, 1):
            a, b = ws.cell(row=r, column=1, value=k), ws.cell(row=r, column=2, value=v)
            a.font = Font(bold=True, color="FFFFFF" if r == 1 else sr.NAVY)
            if r == 1:
                a.fill = PatternFill("solid", fgColor=sr.NAVY)
            b.alignment = Alignment(wrap_text=True, vertical="top")
        title = (plan.get("result_title") or "Result")[:31]
        for bad in '[]:*?/\\':
            title = title.replace(bad, " ")
        result.to_excel(xl, sheet_name=title or "Result", index=False)
        rs = xl.sheets[title or "Result"]
        for c in rs[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=sr.NAVY)
        for col in rs.columns:
            width = max(len(str(col[0].value or "")), *(len(str(c.value or "")) for c in col[1:200])) + 2
            rs.column_dimensions[col[0].column_letter].width = min(max(width, 10), 60)
        rs.freeze_panes = "A2"
    return buf.getvalue()
