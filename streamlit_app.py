"""
Streamlit web app — NIGP-Sourced Procurement Category Mapper.

Multi-page browser app wrapping classifier_JHK3.py. Authorized users
sign in once, then navigate via the sidebar between:
  - Classify (single description + recent-history list)
  - Bulk Classify (paste-many or CSV upload, downloadable result)
  - Procurement Taxonomy Logic (three-level hierarchy explainer)
  - Methodology (4-tier classifier model + confidence-badge teaching)
  - Business Categories (the 17-category reference)
  - Rule Lookup (search the curated + AI-mined rule catalog)

Author: James H. Kirby III, CSCP, MS-SCM

DEPLOYMENT (Streamlit Community Cloud):
  1. Visit https://share.streamlit.io and sign in with GitHub.
  2. New app → this repo, branch `main`, main file `streamlit_app.py`.
  3. Advanced settings → Secrets → set:
         password = "YOUR-CHOSEN-PASSWORD"
  4. Deploy. App goes live at <subdomain>.streamlit.app within ~2 min.
"""
from __future__ import annotations

import io
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "spend-analysis" / "scripts"))
from classifier_JHK3 import classify_one, load_keyword_rules  # noqa: E402
import spend_report_JHK3 as sr  # noqa: E402

# -- Brand palette ---------------------------------------------------------
CHI_NAVY = "#002F6C"
CHI_BLUE = "#41B6E6"
CHI_RED = "#DA291C"
CHI_LT_BLUE = "#D6EEF9"
CHI_GREEN = "#0E7C3A"
CHI_GRAY = "#595959"

AUTHOR_NAME = "James H. Kirby III, CSCP, MS-SCM"
AUTHOR_TITLE_LINE = "Independent Public-Sector Procurement Tool"

# -- Page config -----------------------------------------------------------
st.set_page_config(
    page_title="NIGP-Sourced Procurement Category Mapper",
    layout="wide",
    initial_sidebar_state="expanded",
)


def apply_css() -> None:
    st.markdown(
        f"""
        <style>
          .stApp {{ background-color: #FAFCFE; }}
          h1 {{
            color: {CHI_NAVY};
            border-bottom: 3px solid {CHI_RED};
            padding-bottom: 0.3em;
          }}
          h2, h3, h4 {{ color: {CHI_BLUE}; }}
          .stButton > button {{
            background-color: {CHI_NAVY};
            color: white;
            border: 0;
            font-weight: 600;
            padding: 0.5em 1.5em;
          }}
          .stButton > button:hover {{
            background-color: {CHI_RED};
            color: white;
          }}
          .stDownloadButton > button {{
            background-color: {CHI_BLUE};
            color: white;
            border: 0;
            font-weight: 600;
          }}
          .stDownloadButton > button:hover {{
            background-color: {CHI_NAVY};
            color: white;
          }}
          .chi-badge {{
            display: inline-block;
            padding: 2px 10px;
            border-radius: 3px;
            font-weight: 700;
            font-size: 11px;
            letter-spacing: 0.5px;
            color: white;
          }}
          .chi-result-box {{
            background: {CHI_LT_BLUE};
            padding: 16px 18px;
            border-left: 6px solid {CHI_NAVY};
            border-radius: 4px;
            margin-top: 8px;
          }}
          .chi-result-category {{
            color: {CHI_NAVY};
            font-size: 22px;
            font-weight: 700;
            margin-top: 6px;
          }}
          .chi-callout {{
            background: {CHI_LT_BLUE};
            border-left: 6px solid {CHI_RED};
            padding: 14px 18px;
            border-radius: 4px;
            margin: 14px 0;
          }}
          .chi-callout-label {{
            color: {CHI_RED};
            font-weight: 700;
            font-size: 12px;
            letter-spacing: 0.5px;
            text-transform: uppercase;
          }}
          .chi-footer {{
            color: {CHI_GRAY};
            font-size: 11px;
            text-align: center;
            margin-top: 32px;
            padding-top: 12px;
            border-top: 1px solid #ddd;
          }}
          .chi-author-card {{
            background: white;
            border: 1px solid {CHI_BLUE};
            border-radius: 4px;
            padding: 12px 14px;
            margin: 16px 0;
          }}
          [data-testid="stSidebar"] {{ background-color: #F2F8FC; }}
          [data-testid="stSidebar"] h3 {{ color: {CHI_NAVY}; }}
          [data-testid="stSidebar"] [role="radiogroup"] label {{
            color: {CHI_NAVY} !important;
            font-weight: 500;
            font-size: 15px;
            padding: 4px 0;
          }}
          [data-testid="stSidebar"] [role="radiogroup"] label p {{
            color: {CHI_NAVY} !important;
            font-size: 15px !important;
            margin: 0 !important;
          }}
        </style>
        """,
        unsafe_allow_html=True,
    )


# -- Password gate ---------------------------------------------------------
def check_password() -> bool:
    if st.session_state.get("password_correct"):
        return True

    st.title("NIGP-Sourced Procurement Category Mapper")
    st.markdown(
        f"**Authorized access only.**  "
        f"Enter the access password to continue."
    )
    pw = st.text_input("Password", type="password", key="pw_input")
    if st.button("Sign in"):
        try:
            expected = st.secrets["password"]
        except (KeyError, FileNotFoundError):
            st.error(
                "This server is missing the access-password configuration. "
                f"Contact the administrator ({AUTHOR_NAME}) to set the `password` secret."
            )
            return False
        if pw == expected:
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.markdown(
        f"<div class='chi-footer'>{AUTHOR_TITLE_LINE}  •  NIGP Classifier v1.1  •  "
        f"Created by {AUTHOR_NAME}</div>",
        unsafe_allow_html=True,
    )
    return False


# -- Cached data loaders ---------------------------------------------------
@st.cache_resource(show_spinner="Loading classification rules...")
def get_rules() -> pd.DataFrame:
    return load_keyword_rules()


@st.cache_data
def get_categories_summary() -> pd.DataFrame:
    path = ROOT / "spend-analysis" / "data" / "reference" / "business_categories_summary_JHK3.csv"
    return pd.read_csv(path)


# -- Session state init ----------------------------------------------------
def init_session_state() -> None:
    if "history" not in st.session_state:
        st.session_state.history = []


# -- Helpers ---------------------------------------------------------------
def classify_with_history(description: str) -> dict:
    """Run classifier and append to in-session history (max 10 entries)."""
    rules = get_rules()
    result = classify_one(description, rules)
    d = result.as_dict()
    entry = {
        "Description": description.strip(),
        "Business_Category": d["Business_Category"] or "—",
        "NIGP_Code": d["NIGP_Code_Assigned"] or "—",
        "Confidence": d["Classification_Confidence"] or "—",
        "Review_Flag": d["Review_Flag"],
    }
    history = st.session_state.history
    history.insert(0, entry)
    st.session_state.history = history[:10]
    return d


def render_result_box(d: dict) -> None:
    if d["Review_Flag"] == "Yes" and d["Business_Category"] == "":
        badge_color = CHI_RED
        badge_text = "NO RULE MATCHED"
        category_display = "Add a rule for this description"
    elif d["Review_Flag"] == "Yes":
        badge_color = CHI_RED
        badge_text = "REVIEW RECOMMENDED"
        category_display = d["Business_Category"]
    elif d["Classification_Confidence"] == "High":
        badge_color = CHI_GREEN
        badge_text = "HIGH CONFIDENCE"
        category_display = d["Business_Category"]
    else:
        badge_color = CHI_BLUE
        badge_text = f"{d['Classification_Confidence'].upper()} CONFIDENCE"
        category_display = d["Business_Category"]

    # NIGP code line — always shown so the user always knows what happened
    if d["NIGP_Code_Assigned"]:
        level = d["NIGP_Match_Level"] or ""
        level_label = {
            "exact": "5-digit Item",
            "broad": "3-digit Class",
            "review": "broader Class (review)",
        }.get(level, level)
        nigp_line_html = (
            f"<div style='color:{CHI_NAVY}; font-size:15px; font-weight:600; "
            f"margin-top:8px;'>NIGP Code:  "
            f"<span style='color:{CHI_RED};'>{d['NIGP_Code_Assigned']}</span>"
            f"  <span style='color:{CHI_GRAY}; font-weight:400; font-size:13px;'>"
            f"({level_label})</span></div>"
        )
    elif d["Business_Category"] == "Grants & Pass-Through Funding":
        nigp_line_html = (
            f"<div style='color:{CHI_GRAY}; font-size:13px; font-style:italic; "
            f"margin-top:8px;'>"
            f"NIGP Code:  <em>does not apply</em> — subgrant disbursements are "
            f"financial transfers, not commodity purchases."
            f"</div>"
        )
    elif d["Business_Category"]:
        nigp_line_html = (
            f"<div style='color:{CHI_GRAY}; font-size:13px; font-style:italic; "
            f"margin-top:8px;'>"
            f"NIGP Code:  <em>not yet mapped for this rule</em> — Business "
            f"Category is assigned, but the matching rule has no NIGP class "
            f"on file. Add a code to the rule in "
            f"<code>keyword_rules_DRAFT_JHK3.csv</code> to populate this field."
            f"</div>"
        )
    else:
        nigp_line_html = (
            f"<div style='color:{CHI_GRAY}; font-size:13px; font-style:italic; "
            f"margin-top:8px;'>"
            f"NIGP Code:  <em>none assigned</em> — no curated or AI-mined rule fired. "
            f"Add a rule for this description in <code>keyword_rules_DRAFT_JHK3.csv</code> "
            f"to teach the classifier."
            f"</div>"
        )

    st.markdown(
        f"""
        <div class='chi-result-box'>
          <span class='chi-badge' style='background:{badge_color};'>{badge_text}</span>
          <div class='chi-result-category'>{category_display}</div>
          {nigp_line_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_detail_table(d: dict) -> None:
    rows = [
        ("Business Category", d["Business_Category"] or "—"),
        ("NIGP 3-digit Class", d["NIGP_Class_3digit"] or "—"),
        ("NIGP 5-digit Item", d["NIGP_Item_5digit"] or "—"),
        ("NIGP Code Assigned", d["NIGP_Code_Assigned"] or "—"),
        ("Match Level", d["NIGP_Match_Level"] or "—"),
        (
            "Confidence",
            f"{d['Classification_Confidence']}  ({d['Confidence_Score']:.2f})"
            if d["Classification_Confidence"]
            else "—",
        ),
        ("Method", d["Classification_Method"] or "—"),
        ("Review Flag", d["Review_Flag"]),
        ("Reason", d["Classification_Reason"] or "—"),
    ]
    st.table(pd.DataFrame(rows, columns=["Field", "Value"]))


# =========================================================================
# PAGE: Classify (single description + history)
# =========================================================================
def page_classify() -> None:
    st.title("Classify a Description")
    st.markdown(
        "Paste a purchase order, requisition, or invoice description below. "
        "The classifier returns a Business Category, NIGP code, confidence "
        "score, and the rule that fired."
    )

    with st.form("classify_form"):
        description = st.text_area(
            "Description to classify",
            placeholder="e.g.,  RENTAL OF HEAVY EQUIPMENT FOR AIRPORT TAXIWAY REPAIR",
            height=120,
        )
        submitted = st.form_submit_button("Classify")

    if submitted:
        if not description.strip():
            st.warning("Please enter a description.")
        else:
            d = classify_with_history(description)
            st.markdown("### Result")
            render_result_box(d)
            st.markdown("#### Classification detail")
            render_detail_table(d)
            with st.expander("What does \"REVIEW RECOMMENDED\" mean?"):
                st.markdown(
                    "When a description doesn't confidently match a rule, the classifier "
                    "**does not guess.** It flags the record for procurement-staff review "
                    "instead. About 14% of all historical records are in this queue today. "
                    "Procurement staff triage them on a monthly or quarterly cadence, and "
                    "their decisions become new rules — the queue shrinks over time. See "
                    "the **Methodology** page for the full explanation."
                )

    if st.session_state.history:
        st.markdown("---")
        st.markdown("### Recent classifications (this session)")
        st.caption("The last 10 descriptions you classified during this browser session.")
        hist_df = pd.DataFrame(st.session_state.history)
        st.dataframe(hist_df, use_container_width=True, hide_index=True)
        if st.button("Clear history"):
            st.session_state.history = []
            st.rerun()


# =========================================================================
# PAGE: Bulk Classify (paste-many or CSV upload)
# =========================================================================
def read_uploaded_table(uploaded) -> pd.DataFrame:
    """Read an uploaded CSV/Excel file robustly, tolerating non-UTF-8 encodings
    (Excel exports are often cp1252/latin-1). Tries UTF-8 first, then falls back
    to cp1252, then latin-1 (which decodes any byte)."""
    name = uploaded.name.lower()
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded)
    raw = uploaded.getvalue()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(raw), encoding=enc, low_memory=False)
        except UnicodeDecodeError:
            continue
    return pd.read_csv(io.BytesIO(raw), encoding="latin-1",
                       encoding_errors="replace", low_memory=False)


def page_bulk() -> None:
    st.title("Bulk Classify")
    st.markdown(
        "Classify many descriptions at once. Two modes — pick whichever fits "
        "the file in front of you."
    )

    mode = st.radio(
        "Mode",
        ["Paste many descriptions (one per line)", "Upload a CSV file"],
        horizontal=False,
    )
    rules = get_rules()

    if mode.startswith("Paste"):
        text = st.text_area(
            "One description per line",
            height=200,
            placeholder=(
                "RENTAL OF HEAVY EQUIPMENT\n"
                "ANNUAL HVAC PREVENTIVE MAINTENANCE\n"
                "DFSS-2024-VIOLENCE-PREVENTION\n"
                "POLICE UNIFORMS — CLASS A"
            ),
        )
        if st.button("Classify all"):
            lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
            if not lines:
                st.warning("No descriptions to classify.")
            else:
                results = []
                progress = st.progress(0.0, text="Classifying...")
                for i, ln in enumerate(lines):
                    r = classify_one(ln, rules).as_dict()
                    r["Description"] = ln
                    results.append(r)
                    progress.progress((i + 1) / len(lines))
                progress.empty()
                df = pd.DataFrame(results)
                cols = [
                    "Description",
                    "Business_Category",
                    "NIGP_Code_Assigned",
                    "NIGP_Match_Level",
                    "Classification_Confidence",
                    "Confidence_Score",
                    "Review_Flag",
                    "Classification_Method",
                    "Classification_Reason",
                ]
                df = df[cols]
                _render_bulk_result(df, "paste_results")
        return

    # CSV upload mode
    uploaded = st.file_uploader(
        "Drop a CSV file with at least one column containing descriptions.",
        type=["csv"],
        accept_multiple_files=False,
    )
    if uploaded is None:
        st.info("Awaiting file upload…")
        return

    try:
        df_in = read_uploaded_table(uploaded)
    except Exception as e:
        st.error(f"Could not read the file: {e}")
        return

    st.success(f"Loaded {len(df_in):,} rows × {len(df_in.columns)} columns.")
    st.markdown("**Preview (first 5 rows):**")
    st.dataframe(df_in.head(), use_container_width=True, hide_index=True)

    desc_col = st.selectbox(
        "Which column contains the description text?",
        options=list(df_in.columns),
    )
    if st.button("Classify all rows"):
        results = []
        progress = st.progress(0.0, text=f"Classifying {len(df_in):,} rows...")
        for i, row in df_in.iterrows():
            desc = str(row[desc_col]) if pd.notna(row[desc_col]) else ""
            r = classify_one(desc, rules).as_dict()
            results.append(r)
            if (i + 1) % max(1, len(df_in) // 100) == 0:
                progress.progress((i + 1) / len(df_in))
        progress.empty()

        result_df = pd.DataFrame(results)
        out_df = pd.concat([df_in.reset_index(drop=True), result_df], axis=1)
        st.markdown(f"**Classified {len(out_df):,} rows.** Preview:")
        st.dataframe(out_df.head(20), use_container_width=True, hide_index=True)
        _render_bulk_result(out_df, "csv_results")


def _render_bulk_result(df: pd.DataFrame, key: str) -> None:
    st.markdown("### Results")
    st.dataframe(df, use_container_width=True, hide_index=True)
    csv_bytes = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Download results as CSV",
        data=csv_bytes,
        file_name=f"NIGP_classification_{key}.csv",
        mime="text/csv",
    )


# =========================================================================
# PAGE: Methodology
# =========================================================================
def page_methodology() -> None:
    st.title("Methodology & Why You Can Trust It")
    st.markdown(f"**Created by:** {AUTHOR_NAME}")
    st.markdown("---")

    st.markdown(
        "This tool classifies procurement descriptions into one of **17 Business Categories** "
        "and assigns an **NIGP code** when one applies. It uses deterministic rules — no AI "
        "runs at classification time."
    )

    st.subheader("The six sidebar pages")
    st.markdown(
        "- **Classify** — paste a single description, get a category and NIGP code.\n"
        "- **Bulk Classify** — paste many descriptions or upload a CSV; download the classified file.\n"
        "- **Procurement Taxonomy Logic** — three-level explainer of how spend rolls up.\n"
        "- **Methodology** — this page.\n"
        "- **Business Categories** — reference list of all 17 categories with definitions.\n"
        "- **Rule Lookup** — search the rule catalog by keyword to see why a description routed where it did."
    )

    st.subheader("What this tool was built on")
    st.markdown(
        "The taxonomy and every rule were derived from a **784,556-row public-sector "
        "procurement extract** (87 columns; AP activity years **2017, 2020, 2021, 2023**). "
        "Every purchase-order and invoice line was classified **independently from its own "
        "description text** — not re-labeled from anyone's prior categorization. The source "
        "file already carried partial prior commodity codes on about 30% of rows; the "
        "classifier **does not use them as inputs**. They are preserved only for optional "
        "after-the-fact cross-checking (\"did our independent answer agree?\")."
    )

    st.subheader("What it deliberately ignores — and why")
    st.markdown(
        "Credibility comes as much from what the classifier refuses to use as from what it "
        "does. It reads **only** the description text and the agency's own account/object/fund "
        "codes. It deliberately excludes:\n"
        "- **Vendor name** — the same vendor sells across many categories on one contract, so "
        "\"who sold it\" is an unreliable guide to \"what it is.\" Using vendor would *introduce* "
        "misclassification, not reduce it.\n"
        "- **Spend dollars** — the tool answers *what was bought*, never *how much was spent*. "
        "Dollars are a downstream analysis that consumes this taxonomy — never an input to it.\n"
        "- **Prior consultant codes** — building an owned, defensible classification means "
        "classifying from scratch, not inheriting another party's work product."
    )

    st.subheader("How the 4-tier classifier model works")
    st.markdown(
        "The classifier validated against the 784,556-row historical dataset achieves **100% "
        "coverage — every row carries a Business Category, zero rows in a review queue.** It "
        "gets there through four deterministic tiers. Each row is resolved by exactly one tier, "
        "and the tier is recorded on the row so any classification can be filtered, audited, "
        "or traced back to its source."
    )
    st.markdown(
        f"""
        <div style="margin: 18px 0;">
          <table style="width:100%; border-collapse: collapse; font-size: 14px;">
            <thead>
              <tr style="background:{CHI_NAVY}; color:white;">
                <th style="padding:8px 10px; text-align:left;">Tier</th>
                <th style="padding:8px 10px; text-align:left;">Method</th>
                <th style="padding:8px 10px; text-align:right;">Rows</th>
                <th style="padding:8px 10px; text-align:right;">Share</th>
                <th style="padding:8px 10px; text-align:left;">What it does</th>
              </tr>
            </thead>
            <tbody>
              <tr style="background:{CHI_LT_BLUE};">
                <td style="padding:8px 10px;"><strong>1</strong></td>
                <td style="padding:8px 10px;">Keyword rule</td>
                <td style="padding:8px 10px; text-align:right;">688,044</td>
                <td style="padding:8px 10px; text-align:right;"><strong>87.7%</strong></td>
                <td style="padding:8px 10px;">293 curated + 6,766 AI-mined rules match the description text. Strongest evidence.</td>
              </tr>
              <tr>
                <td style="padding:8px 10px;"><strong>2</strong></td>
                <td style="padding:8px 10px;">Account-code pattern</td>
                <td style="padding:8px 10px; text-align:right;">2,028</td>
                <td style="padding:8px 10px; text-align:right;"><strong>0.3%</strong></td>
                <td style="padding:8px 10px;">No keyword rule fired. A source-system account code (e.g., the subgrant-disbursement account series) resolves it.</td>
              </tr>
              <tr style="background:{CHI_LT_BLUE};">
                <td style="padding:8px 10px;"><strong>3</strong></td>
                <td style="padding:8px 10px;">AI-assist resolver</td>
                <td style="padding:8px 10px; text-align:right;">93,518</td>
                <td style="padding:8px 10px; text-align:right;"><strong>11.9%</strong></td>
                <td style="padding:8px 10px;">No rule or account match. Resolver looks up the description in the saved 2026-04-30 AI mining output and applies the model's category. <strong>No new API call.</strong> Tagged <code>AI-high</code>, <code>AI-medium</code>, or <code>AI-low</code>.</td>
              </tr>
              <tr>
                <td style="padding:8px 10px;"><strong>4</strong></td>
                <td style="padding:8px 10px;">Unclassified — No Description</td>
                <td style="padding:8px 10px; text-align:right;">966</td>
                <td style="padding:8px 10px; text-align:right;"><strong>0.1%</strong></td>
                <td style="padding:8px 10px;">No usable text in any of the four description fields. Explicitly tagged rather than silently dropped.</td>
              </tr>
              <tr style="background:{CHI_NAVY}; color:white; font-weight:600;">
                <td style="padding:8px 10px;">—</td>
                <td style="padding:8px 10px;">Total mapped</td>
                <td style="padding:8px 10px; text-align:right;">784,556</td>
                <td style="padding:8px 10px; text-align:right;">100.0%</td>
                <td style="padding:8px 10px;">Zero rows in the review queue.</td>
              </tr>
            </tbody>
          </table>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        "**Single-record vs. batch:** This live tool runs **Tier 1 + Tier 2 only** for the "
        "description you paste — it does not call the resolver or any AI. The 11.9% AI-assist "
        "tier was applied once to the historical 784K dataset and is preserved in the "
        "deliverable file. If a live description doesn't match any rule, the result box says "
        "*NO RULE MATCHED* and points you to the curated rule file — the right next action "
        "is to add a rule, not to wait on AI."
    )
    st.markdown(
        f"<div style='color:{CHI_RED}; font-style:italic; font-size:14px;'>"
        "Note on transparency: 100% mapped is not the same statement as 0% misclassification "
        "risk. Tier 3 rows carry an <code>AI-medium</code> or <code>AI-low</code> confidence "
        "tag. The <code>Classification_Confidence</code> field on every row exposes the tier — "
        "leadership and auditors can filter or exclude any tier they choose."
        "</div>",
        unsafe_allow_html=True,
    )

    st.subheader("Why the one-time AI use is defensible")
    st.markdown(
        f"<div style='border-left: 4px solid {CHI_RED}; padding: 12px 16px; "
        f"background: {CHI_LT_BLUE}; margin: 6px 0 14px 0; font-size:14px; line-height:1.6;'>"
        "<strong>The production tool is rules-only. No AI model is called when you classify "
        "anything.</strong> AI was used <strong>once</strong>, at build time, to read the long "
        "tail of ~30,000 descriptions that no hand-written rule covered and propose patterns "
        "(total one-time cost about $27). Those proposals were frozen into rules and the model "
        "was never called again."
        "</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "1. **One-time, not ongoing.** Runs never call AI — no API key, no internet, no "
        "per-classification charge.\n"
        "2. **Bounded.** The model could only choose from the 17 approved categories and 138 "
        "NIGP classes — it cannot invent a code.\n"
        "3. **Transparent.** Every AI-mined rule records the model, confidence, source "
        "frequency, and reasoning in its `notes` field.\n"
        "4. **Reviewable & editable.** Staff can demote or delete any AI-mined rule; it has no "
        "special status in the classifier — it is matched by the same logic as a curated rule.\n"
        "5. **Replaceable.** Delete all AI content and the tool still runs on curated rules + "
        "account patterns. Coverage would drop; nothing would break.\n"
        "6. **Confidence-labeled, never laundered.** AI-derived answers are tagged "
        "`AI-high / AI-medium / AI-low` so leadership and auditors can filter or exclude them — "
        "they are never hidden inside a generic \"auto-classified\" bucket."
    )

    st.subheader("How descriptions get mapped to actual NIGP codes")
    st.markdown(
        "Believability rests on the chain of custody from a free-text description to a "
        "specific NIGP code. **The classifier does not invent codes.** Every rule has been "
        "pre-authored with an explicit NIGP class attached (and a 5-digit item when the "
        "description supports that specificity). Here is the chain, step by step:"
    )
    st.markdown(
        "1. **Description selection.** Each procurement row carries up to four description "
        "fields (PO Description, Line Description, Item Description, Vendor Item Description). "
        "The classifier picks the first non-blank, substantive description (more than 3 "
        "characters, not generic placeholders like \"Per contract\" or \"Misc\") across all "
        "four. The chosen text is preserved in <code>Description_Best</code> so every "
        "classification traces back to the exact text used.\n"
        "2. **Rule matching.** The chosen description is matched against the 7,059-rule "
        "catalog (293 curated + 6,766 AI-mined). Every rule was pre-authored with an "
        "explicit <code>nigp_class_3digit</code> and (where specificity allows) a "
        "<code>nigp_item_5digit</code>. First match wins; subsequent rules don't fire. "
        "The rule file is a CSV — open it in Excel to see the exact mapping for any pattern.\n"
        "3. **NIGP code assignment.** The matched rule's NIGP class (and 5-digit item, when "
        "present) is copied verbatim onto the row. The rule pattern that fired is recorded "
        "in <code>Classification_Reason</code>, so any auditor can reconstruct the decision "
        "from the row alone — no system access required.\n"
        "4. **Business Category rollup.** The 17-bucket Business Category field is set "
        "deterministically from a canonical NIGP-class → Business-Category map "
        "(<code>business_categories_JHK3.csv</code>, 138 rows with judgment notes). The "
        "mapping is fixed and version-controlled — categories never drift between batch runs.",
        unsafe_allow_html=True,
    )
    st.markdown(
        "**The NIGP codes themselves are not invented by this tool.** They come from the "
        "publicly-recognized NIGP Commodity/Services Code framework <strong>*</strong> — the "
        "same framework used by hundreds of U.S. state, county, and municipal governments "
        "for procurement classification, inter-agency benchmarking, and audit defensibility. "
        "The 138 classes, 470 5-digit items, and 4,592 10-digit codes used by this project "
        "were extracted from a real-world public-sector procurement transaction history "
        "<strong>**</strong> and validated against the public NIGP framework. **No code "
        "assigned by this classifier sits outside the NIGP standard.**",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div style='border-left: 4px solid {CHI_RED}; padding: 12px 16px; "
        f"background: {CHI_LT_BLUE}; margin: 14px 0; font-size:13px; line-height:1.6;'>"
        "<strong>Sources cited above</strong><br>"
        "<strong>*</strong> &nbsp; <strong>NIGP Commodity/Services Code framework</strong> "
        "— maintained by NIGP, The Institute for Public Procurement "
        "(<a href='https://www.nigp.org/' target='_blank'>https://www.nigp.org/</a>). "
        "The foundational public-standard taxonomy used across U.S. public-sector "
        "procurement.<br>"
        "<strong>**</strong> &nbsp; <strong>Project working catalog of NIGP codes</strong> "
        "— derived from a real-world public-sector procurement transaction history (a "
        "five-year AP dataset), validated against the public NIGP framework. "
        "Stored in this repo as <code>nigp_codes_3digit_JHK3.csv</code> (138 classes), "
        "<code>nigp_codes_5digit_JHK3.csv</code> (470 items), and "
        "<code>nigp_codes_10digit_JHK3.csv</code> (4,592 codes).<br>"
        "<strong>Hand-curated rules (293)</strong> — authored by James H. Kirby III, "
        "CSCP, MS-SCM, drawing on 28 years of public-sector procurement domain expertise.<br>"
        "<strong>AI-mined rules (6,766)</strong> — generated by Anthropic Claude Haiku 4.5 "
        "in a single bounded 2026-04-30 batch run. The model was constrained by JSON schema "
        "to output only codes within the 17-category / 138-class catalog already approved "
        "by procurement leadership. The AI cannot invent codes outside that catalog. "
        "Provenance metadata (model, confidence, source row count, reasoning) is preserved "
        "in every AI-mined rule's <code>notes</code> field for audit."
        "</div>",
        unsafe_allow_html=True,
    )

    st.subheader("Reading a result")
    st.markdown(
        "Each single-record classification shows a confidence badge:\n"
        "- **HIGH** — strong rule match, trust it.\n"
        "- **MEDIUM** — rule fired with caveats, spot-check.\n"
        "- **REVIEW RECOMMENDED** — a rule fired but is itself flagged for review.\n"
        "- **NO RULE MATCHED** — no curated or AI-mined rule fired. The live tool does not "
        "guess; the right action is to add a rule. (In the historical batch, the Tier 3 "
        "resolver would have caught this from the saved AI output.)"
    )

    st.subheader("How the confidence badge is calculated")
    st.markdown(
        "Every NIGP code assigned by the classifier carries a confidence badge. The badge "
        "reflects how specific the evidence was — not how the classifier \"felt.\" It is set "
        "by the rule that fired, not invented after the fact."
    )
    st.markdown(
        "- **HIGH** — The description matched a specific rule that maps directly to a precise "
        "NIGP 5-digit Item code (e.g., \"READY-MIX CONCRETE\" → NIGP class 750, item 75070). "
        "The NIGP assignment is exact, defensible, and traceable to a single rule.\n"
        "- **MEDIUM** — The description matched a broader rule that resolves to an NIGP 3-digit "
        "Class but not a specific 5-digit Item (e.g., generic \"PIPE FITTINGS\" lands in the "
        "plumbing class without pinning the exact item). The Business Category and 3-digit "
        "Class are reliable; the 5-digit Item is best-effort. Spot-check recommended for "
        "high-dollar transactions.\n"
        "- **AI-HIGH / AI-MEDIUM / AI-LOW** — In the historical batch file, rows that no "
        "keyword rule matched were resolved by the AI-assist tier from the saved 2026-04-30 "
        "mining output. The confidence sub-tier is the model's own self-rating. Filter on "
        "these tags to audit AI-tier classifications. **No new API call is made.**\n"
        "- **REVIEW RECOMMENDED / NO RULE MATCHED** — In the live tool, the matching rule was "
        "flagged for review, or no rule fired at all. The classifier does **not** assign a "
        "guess. The right next action is to add a rule to <code>keyword_rules_DRAFT_JHK3.csv</code>."
    )
    st.markdown(
        "Every classification is auditable. The exact rule that fired (or the resolver "
        "lookup that filled it) is recorded on every row, so any NIGP code can be traced "
        "back to its source."
    )

    st.subheader("How we know it's right — regression tested")
    st.markdown(
        "The classifier ships with a **plain-English regression test**: a fixed list of "
        "everyday procurement phrases, each paired with the Business Category it must return. "
        "The current suite passes **72 of 72 with zero mis-categorizations**. Every rule change "
        "is re-run against this suite before it goes live, so a fix in one place cannot silently "
        "break a classification somewhere else. This is the same discipline software teams use "
        "to keep a system trustworthy as it grows — the rule base can expand without eroding "
        "what already worked."
    )

    st.subheader("Audit cadence (what changed in May & August 2026)")
    st.markdown(
        "Through 2026-04-30 the classifier produced a 17.8% review queue — rows flagged "
        "for procurement-staff triage. The 2026-05-14 refresh eliminated that queue by "
        "(a) adding 98 curated rules, (b) deploying the Tier 3 AI-assist resolver to "
        "consume the latent medium/low confidence AI proposals at fill time, and "
        "(c) explicitly tagging the 966 no-description rows rather than silently dropping "
        "them. Staff time is now spent **auditing Tier 3** and promoting recurring "
        "high-quality AI-assist matches into the curated rule file — the rule base grows; "
        "Tier 3 shrinks over time."
    )
    st.markdown(
        "**August 2026 — interactive robustness pass.** A user report that a common word "
        "returned no result surfaced that the *live* single-record tool answers only when typed "
        "text matches a rule. In response, **34 curated rules** were added — each mapped to its "
        "correct NIGP class **verified against the reference table**, not auto-guessed — raising "
        "coverage on a 45-term common-word probe from **38% to 78%** and keeping the regression "
        "suite at **72/72**. The curated rule base grew **246 → 280**. Runtime stayed "
        "rules-only, no API key."
    )

    st.subheader("Need help?")
    st.markdown(f"Contact **{AUTHOR_NAME.split(',')[0]}**.")


# =========================================================================
# PAGE: Procurement Taxonomy Logic
# =========================================================================
def page_taxonomy_logic() -> None:
    st.title("Procurement Taxonomy Logic")
    st.markdown(
        "This procurement taxonomy organizes spend data into three hierarchical levels. "
        "Each level serves a different audience and answers a different question."
    )
    st.markdown("---")

    st.subheader("How the three levels fit together")
    st.markdown(
        f"""
        <div style="max-width: 520px; margin: 24px auto;">
          <div style="background:{CHI_NAVY}; color:white; padding:18px 20px; border-radius:8px; text-align:center;">
            <div style="font-size:12px; letter-spacing:1px; opacity:0.85;">LEVEL 1</div>
            <div style="font-size:20px; font-weight:700; margin-top:4px;">Executive View</div>
            <div style="font-size:14px; margin-top:6px;">17 Business Categories</div>
            <div style="font-size:12px; opacity:0.85; margin-top:4px;">Audience: Leadership &amp; Budget Decisions</div>
          </div>
          <div style="text-align:center; color:{CHI_GRAY}; font-size:24px; line-height:1; margin:6px 0;">▼</div>
          <div style="background:{CHI_BLUE}; color:{CHI_NAVY}; padding:18px 20px; border-radius:8px; text-align:center;">
            <div style="font-size:12px; letter-spacing:1px; opacity:0.9;">LEVEL 2</div>
            <div style="font-size:20px; font-weight:700; margin-top:4px;">Sourcing View</div>
            <div style="font-size:14px; margin-top:6px;">138 NIGP 3-digit Classes</div>
            <div style="font-size:12px; opacity:0.9; margin-top:4px;">Audience: Category Managers &amp; Sourcing Teams</div>
          </div>
          <div style="text-align:center; color:{CHI_GRAY}; font-size:24px; line-height:1; margin:6px 0;">▼</div>
          <div style="background:{CHI_LT_BLUE}; color:{CHI_NAVY}; padding:18px 20px; border-radius:8px; text-align:center; border:1px solid {CHI_BLUE};">
            <div style="font-size:12px; letter-spacing:1px;">LEVEL 3</div>
            <div style="font-size:20px; font-weight:700; margin-top:4px;">Audit View</div>
            <div style="font-size:14px; margin-top:6px;">784,556 Individual Transactions</div>
            <div style="font-size:11px; font-style:italic; color:#E57373; margin-top:4px; padding:0 6px;">PO/invoice line records in the historical dataset the classifier was validated against (AP activity years 2017, 2020, 2021, 2023).</div>
            <div style="font-size:12px; margin-top:6px;">Audience: Auditors, Analysts &amp; Anyone Asking "Why?"</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div style='font-size:13px; color:{CHI_GRAY}; font-style:italic; "
        f"text-align:center; margin: -8px 0 12px 0;'>"
        "Note: \"Level\" here refers to the taxonomy hierarchy (how spend rolls up). "
        "The Methodology page describes a separate 4-<em>tier</em> classification pipeline "
        "(how each row is assigned to a category). Two different concepts; do not confuse."
        "</div>",
        unsafe_allow_html=True,
    )

    st.subheader("Level 1 — Executive View")
    st.markdown(
        "**17 Business Categories** that show where procurement dollars go. This is the "
        "strategic level for leadership attention and budget decisions. Categories are "
        "**mutually exclusive** (no record fits two) and **collectively exhaustive** "
        "(every record fits one)."
    )

    st.subheader("Level 2 — Sourcing View")
    st.markdown(
        "**Subcategories within each Business Category**, aligned to NIGP 3-digit "
        "commodity classes. This is where consolidation opportunities become visible — "
        "when the same commodity is purchased across multiple departments from multiple "
        "vendors without coordination."
    )

    st.subheader("Level 3 — Audit View")
    st.markdown(
        "**Individual transactions** with full classification detail including the NIGP "
        "code, confidence level, and the exact rule that fired. Every classification is "
        "transparent and defensible — no black-box decisions."
    )

    st.markdown("---")
    st.subheader("Follow one purchase through all three levels")
    st.markdown(
        "To make the hierarchy concrete, here is a single real classification flowing from "
        "raw free text down through every level — the same record answers a different question "
        "for a different audience at each step:"
    )
    st.markdown(
        f"""
        <div style="max-width: 600px; margin: 16px 0;">
          <div style="border:1px dashed {CHI_GRAY}; padding:12px 16px; border-radius:8px; text-align:center; color:{CHI_NAVY};">
            <span style="font-size:11px; letter-spacing:1px; color:{CHI_GRAY};">FREE-TEXT DESCRIPTION (the raw input)</span><br>
            <strong style="font-size:16px;">"chain link fence repair"</strong>
          </div>
          <div style="text-align:center; color:{CHI_GRAY}; font-size:22px; line-height:1; margin:6px 0;">▼</div>
          <div style="background:{CHI_NAVY}; color:white; padding:14px 18px; border-radius:8px;">
            <span style="font-size:11px; letter-spacing:1px; opacity:0.85;">LEVEL 1 · EXECUTIVE VIEW</span><br>
            <strong style="font-size:17px;">Landscaping, Grounds &amp; Irrigation</strong><br>
            <span style="font-size:13px; opacity:0.9;">The category leadership sees on a budget rollup.</span>
          </div>
          <div style="text-align:center; color:{CHI_GRAY}; font-size:22px; line-height:1; margin:6px 0;">▼</div>
          <div style="background:{CHI_BLUE}; color:{CHI_NAVY}; padding:14px 18px; border-radius:8px;">
            <span style="font-size:11px; letter-spacing:1px;">LEVEL 2 · SOURCING VIEW</span><br>
            <strong style="font-size:17px;">NIGP 3-digit Class 988</strong><br>
            <span style="font-size:13px;">The commodity class category managers group on to find consolidation.</span>
          </div>
          <div style="text-align:center; color:{CHI_GRAY}; font-size:22px; line-height:1; margin:6px 0;">▼</div>
          <div style="background:{CHI_LT_BLUE}; color:{CHI_NAVY}; padding:14px 18px; border-radius:8px; border:1px solid {CHI_BLUE};">
            <span style="font-size:11px; letter-spacing:1px;">LEVEL 3 · AUDIT VIEW</span><br>
            <span style="font-size:13px;">The individual transaction, stamped with
            <code>Classification_Reason = matched keyword rule 'FENCE' (contains)</code> &mdash;
            so an auditor can reconstruct the decision from the row alone.</span>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.subheader("Why \"mutually exclusive, collectively exhaustive\" matters")
    st.markdown(
        "The 17 Business Categories are built to be **MECE**, and that property is what lets "
        "leadership trust a category total:\n"
        "- **Mutually exclusive** — no purchase lands in two categories. \"chain link fence "
        "repair\" is *Landscaping, Grounds & Irrigation* and nothing else, so spend is never "
        "double-counted across buckets.\n"
        "- **Collectively exhaustive** — every purchase has a home. Even subgrant "
        "disbursements (Category 17) and rows with no usable description (explicitly tagged "
        "*Unclassified — No Description*, never silently dropped) land somewhere.\n\n"
        "Together they guarantee the 17 categories sum to **100% of spend with no overlaps and "
        "no gaps** — the arithmetic property behind every rollup number in this tool."
    )

    st.markdown("---")
    st.subheader("How the NIGP codes were sourced")
    st.markdown(
        "Level 2 (the 138 NIGP 3-digit Classes) and Level 3 (the 470 NIGP 5-digit Class-Items, "
        "and the 4,592 underlying 10-digit codes) are not invented by this project. They "
        "come from two layered sources, both auditable:"
    )
    st.markdown(
        "1. **The public NIGP Commodity/Services Code framework <strong>*</strong>** — the "
        "inter-agency standard taxonomy maintained by NIGP, The Institute for Public "
        "Procurement. This is the same framework used by hundreds of U.S. state, county, and "
        "municipal procurement offices. Aligning to NIGP provides peer benchmarking, "
        "audit defensibility, and catalog portability.\n"
        "2. **A real-world public-sector procurement transaction history "
        "<strong>**</strong>** — the actual NIGP codes that appeared in the source "
        "dataset (a five-year AP dataset) were extracted and "
        "validated against the public framework. The working catalog used by this "
        "classifier (138 classes / 470 items / 4,592 codes) is the intersection: every "
        "code that appeared in real-world purchasing AND exists in the public NIGP "
        "framework.",
        unsafe_allow_html=True,
    )
    st.markdown(
        "**This tool does not invent NIGP codes.** Every code assigned by the classifier "
        "already existed in the working catalog before any rule was authored. When the "
        "description text doesn't support pinning a specific 5-digit Item, the classifier "
        "returns the broader 3-digit Class with `NIGP_Match_Level = broad` — but never "
        "invents a code. See the Methodology page for the full chain-of-custody from "
        "description text to assigned NIGP code."
    )
    st.markdown(
        f"<div style='border-left: 4px solid {CHI_RED}; padding: 12px 16px; "
        f"background: {CHI_LT_BLUE}; margin: 14px 0; font-size:13px; line-height:1.6;'>"
        "<strong>Sources</strong><br>"
        "<strong>*</strong> &nbsp; <strong>NIGP Commodity/Services Code framework</strong> "
        "— maintained by NIGP, The Institute for Public Procurement "
        "(<a href='https://www.nigp.org/' target='_blank'>https://www.nigp.org/</a>). "
        "The foundational public-standard taxonomy used across U.S. public-sector "
        "procurement.<br>"
        "<strong>**</strong> &nbsp; <strong>Working catalog of NIGP codes used by this "
        "project</strong> — derived from a real-world public-sector procurement "
        "transaction history (a five-year AP dataset), validated against "
        "the public NIGP framework. Lives in this repo as "
        "<code>nigp_codes_3digit_JHK3.csv</code> (138 classes), "
        "<code>nigp_codes_5digit_JHK3.csv</code> (470 items), "
        "<code>nigp_codes_10digit_JHK3.csv</code> (4,592 codes).<br>"
        "<strong>Future-state expansion</strong> — licensing the full ~9,000-code NIGP "
        "catalog from Periscope Holdings / OpenGov would widen Level 3 audit coverage to "
        "the complete NIGP standard. Not required for current operation."
        "</div>",
        unsafe_allow_html=True,
    )


# =========================================================================
# PAGE: Business Categories
# =========================================================================
def page_categories() -> None:
    st.title("17 Business Categories — Reference")
    st.markdown(
        "Every record is assigned to exactly one of these 17 categories. The categories "
        "are designed to be **mutually exclusive** (no record fits two) and **collectively "
        "exhaustive** (every public-sector purchase fits one)."
    )

    cats = get_categories_summary()
    for i, row in cats.iterrows():
        with st.container():
            st.markdown(
                f"""
                <div style='border-left: 4px solid {CHI_BLUE}; padding-left: 12px;
                            margin: 14px 0;'>
                  <div style='color:{CHI_NAVY}; font-size:18px; font-weight:700;'>
                    {i + 1}. {row['business_category']}
                  </div>
                  <div style='color:{CHI_GRAY}; font-size:14px; margin-top:4px;'>
                    {row['definition']}
                  </div>
                  <div style='color:{CHI_BLUE}; font-size:11px; margin-top:6px;'>
                    NIGP 3-digit classes covered: {row['classes_listed']}
                  </div>
                </div>
                """,
                unsafe_allow_html=True,
            )


# =========================================================================
# PAGE: Rule Lookup
# =========================================================================
def page_rule_lookup() -> None:
    st.title("Rule Lookup")
    st.markdown(
        "Type a keyword to see every rule whose pattern contains it. Useful for "
        "answering *\"why did this classification happen?\"* and *\"do we have a rule for X?\"*"
    )

    rules = get_rules()
    st.caption(
        f"Rule catalog: {len(rules):,} total rules — "
        f"{(rules['source'] == 'curated').sum()} hand-curated and "
        f"{(rules['source'] != 'curated').sum():,} AI-mined."
    )

    query = st.text_input(
        "Search rules (case-insensitive substring match on the pattern text)",
        placeholder="e.g.,  HVAC  /  RENTAL  /  DFSS",
    )

    if not query.strip():
        return

    q = query.strip().upper()
    hits = rules[rules["pattern_upper"].str.contains(q, na=False)].copy()
    if hits.empty:
        st.warning(f"No rules match '{query}'.")
        return

    st.success(f"Found {len(hits):,} matching rule(s).")
    display = hits[
        ["pattern", "match_type", "business_category", "nigp_class_3digit",
         "nigp_item_5digit", "nigp_match_level", "source", "notes"]
    ].rename(columns={
        "pattern": "Pattern",
        "match_type": "Match Type",
        "business_category": "Business Category",
        "nigp_class_3digit": "NIGP Class",
        "nigp_item_5digit": "NIGP Item",
        "nigp_match_level": "Match Level",
        "source": "Source",
        "notes": "Notes",
    })
    st.dataframe(display, use_container_width=True, hide_index=True)


# =========================================================================
# PAGE: Spend Report
# =========================================================================
def _fmt_money(v) -> str:
    try:
        return f"${float(v):,.0f}"
    except (TypeError, ValueError):
        return "—"


def _style_spend_df(df, money_all_but_first=False):
    """Return a Styler that shows $ + commas on spend, % on percent columns, and
    thousands separators on counts — for on-screen display in the app."""
    fmt = {}
    for i, col in enumerate(df.columns):
        cl = str(col)
        if money_all_but_first and i > 0:
            fmt[col] = "${:,.0f}"
        elif cl == "Spend":
            fmt[col] = "${:,.0f}"
        elif "%" in cl:
            fmt[col] = "{:.1f}%"
        elif cl in ("Transactions", "Vendors", "Departments"):
            fmt[col] = "{:,.0f}"
    try:
        return df.style.format(fmt, na_rep="—")
    except Exception:  # noqa: BLE001
        return df


def _render_spend_report(res: dict) -> None:
    """Render a previously computed report bundle. Reading from a stored bundle
    (not local variables) is what lets the report survive page navigation."""
    b = res["b"]
    excel_bytes = res["excel"]
    tiles, cov, cat, measure = b["tiles"], b["cov"], b["cat"], b["measure"]

    # Executive summary
    st.markdown("### Executive Summary")
    for ln in b["es"]["lines"]:
        st.markdown(f"- {ln}")
    if b["es"]["steps"]:
        st.markdown("**Recommended first steps:**")
        for i, s in enumerate(b["es"]["steps"], 1):
            st.markdown(f"{i}. {s}")
    if b["es"].get("consolidation_items"):
        st.markdown("**Items you can consolidate** (same item — multiple vendors / departments):")
        for it in b["es"]["consolidation_items"]:
            cat = f" · _{it['category']}_" if it.get("category") else ""
            st.markdown(f"- **{it['item']}** — {it['detail']}, {it['value']}{cat}")

    # Summary tiles
    st.markdown("### Summary")
    m = st.columns(4)
    m[0].metric("Total spend", _fmt_money(tiles["total_spend"]) if tiles["has_amount"] else "n/a")
    m[1].metric("Transactions", f"{tiles['transactions']:,}")
    m[2].metric("Categories", tiles["categories"])
    m[3].metric("Vendors", f"{tiles['vendors']:,}" if tiles["vendors"] is not None else "n/a")

    st.info(
        f"**Coverage:** {cov['classified']:,} of {cov['total']:,} rows "
        f"({cov['classified_pct']}%) mapped to a specific commodity category. The remaining "
        f"{cov['catchall']:,} ({cov['catchall_pct']}%) sit in **General & Other Procurement** — "
        "still counted and shown, with example descriptions so you can see what's in it. "
        "Add keyword rules for the most common of those descriptions to move that spend into "
        "specific categories.")

    # Spend by category
    st.markdown("### Spend by Business Category")
    st.caption("Every row lands in a real category. The **Examples** column shows representative "
               "descriptions from each one — including the General & Other catch-all.")
    st.dataframe(_style_spend_df(cat), use_container_width=True, hide_index=True)
    _chart_col = measure if measure in cat.columns else cat.columns[1]
    st.bar_chart(cat.set_index("Business Category")[_chart_col])

    # Spend trend
    if b["trend"] is not None:
        st.markdown("### Spend Trend Over Time")
        st.dataframe(_style_spend_df(b["trend"]), use_container_width=True, hide_index=True)
        if measure in b["trend"].columns:
            st.bar_chart(b["trend"].set_index("Year")[measure])

    # Pareto
    st.markdown("### Pareto 80/20")
    if len(b["par"]):
        st.caption(f"**{b['n80']}** categor{'y' if b['n80'] == 1 else 'ies'} drive 80% of "
                   f"{'spend' if tiles['has_amount'] else 'transactions'} (classified rows).")
        st.dataframe(_style_spend_df(b["par"]), use_container_width=True, hide_index=True)

    # Top vendors
    if b["vend"] is not None:
        st.markdown("### Top Vendors")
        st.dataframe(_style_spend_df(b["vend"]), use_container_width=True, hide_index=True)

    # Vendor analytics
    if b["concentration"] or b["tail"] or b["single_multi"]:
        st.markdown("### Vendor Analytics")
        vc = st.columns(3)
        if b["concentration"]:
            vc[0].metric("Top-10 vendor share", f"{b['concentration']['top10']}%")
            vc[1].metric("Total vendors", f"{b['concentration']['total_vendors']:,}")
            vc[2].metric("HHI (concentration)", f"{b['concentration']['hhi']:,}")
        if b["tail"]:
            t = b["tail"]
            st.caption(f"**Tail spend:** {t['tail_vendors']:,} vendors ({t['tail_pct_of_vendors']}%) "
                       f"account for only {t['tail_pct_of_value']}% of {t['measure'].lower()} — consolidation targets.")
        if b["single_multi"]:
            sm = b["single_multi"]
            st.caption(f"**Sourcing:** {sm['single_categories']} single-source vs "
                       f"{sm['multi_categories']} multi-source categories.")

    # Spend by department
    if b["dept_tbl"] is not None:
        st.markdown("### Spend by Department")
        st.dataframe(_style_spend_df(b["dept_tbl"]), use_container_width=True, hide_index=True)
        st.bar_chart(b["dept_tbl"].set_index("Department")[
            measure if measure in b["dept_tbl"].columns else b["dept_tbl"].columns[1]])

    # Category × department matrix
    if b["matrix"] is not None:
        st.markdown("### Category × Department Matrix")
        st.dataframe(_style_spend_df(b["matrix"], money_all_but_first=True),
                     use_container_width=True, hide_index=True)

    # Consolidation
    if b["cons"] is not None and len(b["cons"]):
        st.markdown("### Vendor Consolidation / Fragmentation")
        st.caption("Categories bought from more than one vendor — biggest opportunities first.")
        st.dataframe(_style_spend_df(b["cons"]), use_container_width=True, hide_index=True)

    # Same item across vendors / departments (line-level fragmentation) — always shown
    st.markdown("### Same Item — Multiple Vendors / Departments")
    ic = b.get("item_consol")
    if ic is not None and len(ic):
        st.caption("The **same item** (by description) bought from more than one vendor and/or "
                   "across more than one department — line-level fragmentation and the sharpest "
                   "consolidation targets. Runs on every upload. Sorted by spend.")
        st.dataframe(_style_spend_df(ic), use_container_width=True, hide_index=True)
    else:
        st.caption(
            "This check runs on every report. No single item here was bought from more than one "
            "vendor or across more than one department — so there's nothing to consolidate at the "
            "line level. If you expected results, make sure a **Vendor** or **Department** column "
            "is mapped above (the check needs at least one to detect fragmentation).")

    # Same commodity by NIGP code (groups differently-worded descriptions) — always shown
    st.markdown("### Same Commodity (NIGP code) — Multiple Vendors / Departments")
    nc = b.get("nigp_consol")
    if nc is not None and len(nc):
        st.caption("Groups by the **NIGP commodity code** (5-digit item where the rule assigned "
                   "one, otherwise 3-digit class) instead of the exact wording — so two "
                   "differently-worded descriptions of the *same commodity* roll together. "
                   "**Each department is its own column** (spend); the shaded blue cells show which "
                   "departments bought it. Rows shaded **red** in *Descriptions (variants)* are the "
                   "consolidation targets — the same commodity bought across two or more "
                   "departments. Sorted by spend.")
        st.dataframe(_style_nigp_matrix(nc), use_container_width=True, hide_index=True)
    else:
        st.caption("This view groups rows that carry a NIGP commodity code. None of the coded "
                   "commodities here were bought from more than one vendor or department. (Only "
                   "rows whose rule assigned a NIGP code appear here; the item view above covers "
                   "everything by description.)")

    # Dimension breakdowns
    for label, tbl in b["dimensions"]:
        st.markdown(f"### {label}")
        st.dataframe(_style_spend_df(tbl), use_container_width=True, hide_index=True)

    # Excel download — the full leadership workbook (bytes were built once, at generation)
    st.download_button(
        "⬇  Download full spend report (Excel)",
        data=excel_bytes,
        file_name="Spend_Report_JHK3.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _style_nigp_matrix(df):
    """Styler for the NIGP commodity × department matrix: money-format the department
    columns (zeros blank), highlight the departments that actually bought each
    commodity, and flag the multi-department consolidation targets."""
    meta = set(getattr(sr, "NIGP_META_COLS", []))
    variants_col = getattr(sr, "DESC_VARIANTS_COL", "Descriptions (variants)")
    dept_cols = [c for c in df.columns if c not in meta]
    money_cols = list(dept_cols) + [c for c in ["Total Spend"] if c in df.columns]

    def _money_blank_zero(v):
        try:
            return f"${float(v):,.0f}" if float(v) > 0 else ""
        except (TypeError, ValueError):
            return v

    fmt = {}
    for c in df.columns:
        if c in money_cols:
            fmt[c] = _money_blank_zero
        elif c in ("Vendors", "Departments (#)", "Total Transactions"):
            fmt[c] = "{:,.0f}"

    def _highlight(row):
        styles = [""] * len(row)
        nz = 0
        for i, c in enumerate(df.columns):
            if c in dept_cols:
                try:
                    v = float(row[c])
                except (TypeError, ValueError):
                    v = 0
                if v > 0:
                    styles[i] = "background-color:#D6EEF9"
                    nz += 1
        if nz >= 2 and variants_col in df.columns:
            j = list(df.columns).index(variants_col)
            styles[j] = "background-color:#FCE4E4; font-weight:700; color:#9C1006"
        return styles

    try:
        return df.style.format(fmt, na_rep="").apply(_highlight, axis=1)
    except Exception:  # noqa: BLE001
        return df


def page_spend_report() -> None:
    st.title("Spend Report")
    st.markdown(
        "Upload a spend file. The mapper classifies every line into a Business Category, "
        "then builds a spend analysis — spend by category, Pareto 80/20, top vendors, and "
        "vendor consolidation — that you can download as an Excel report. "
        "**Runs on rules only — no AI, and your file stays in this session.**"
    )

    uploaded = st.file_uploader(
        "Upload a spend file (CSV or Excel).",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=False,
        key="sr_upload",
    )
    result = st.session_state.get("sr_result")

    # No file in the uploader (e.g. you navigated back to this page): keep showing
    # the last report you generated, so nothing is lost when you switch pages.
    if uploaded is None:
        if result is not None:
            st.success(
                f"Showing your most recent report — **{result.get('file_name', 'your file')}**. "
                "It stays here while you move between pages. Upload a new file to replace it.")
            if st.button("Clear this report"):
                del st.session_state["sr_result"]
                st.rerun()
            _render_spend_report(result)
        else:
            st.info("Awaiting file upload…")
        return

    try:
        df_in = read_uploaded_table(uploaded)
    except Exception as e:  # noqa: BLE001
        st.error(f"Could not read the file: {e}")
        return
    df_in.columns = [str(c).strip() for c in df_in.columns]
    file_sig = f"{uploaded.name}:{getattr(uploaded, 'size', len(df_in))}"

    st.success(f"Loaded {len(df_in):,} rows × {len(df_in.columns)} columns.")
    st.dataframe(df_in.head(), use_container_width=True, hide_index=True)

    profile = sr.profile_columns(df_in)
    roles = dict(profile["roles"])
    cols = list(df_in.columns)
    NONE = "— none —"

    def _opt_idx(val):
        return cols.index(val) + 1 if val in cols else 0

    st.markdown("#### Detected columns")
    st.caption(
        "The tool inspected **every column** and guessed each one's role — it adapts to any "
        "file, whatever the format. In most cases just click **Generate** below; only change a "
        "box if a guess is wrong. Only **Description** is required.")
    c1, c2, c3 = st.columns(3)
    with c1:
        desc_col = st.selectbox("Description (required)", cols,
            index=cols.index(roles["description"]) if roles.get("description") in cols else 0,
            key="sr_desc")
        amount_col = st.selectbox("Amount / spend", [NONE] + cols,
            index=_opt_idx(roles.get("amount")), key="sr_amount")
    with c2:
        vendor_col = st.selectbox("Vendor", [NONE] + cols,
            index=_opt_idx(roles.get("vendor")), key="sr_vendor")
        dept_col = st.selectbox("Department", [NONE] + cols,
            index=_opt_idx(roles.get("department")), key="sr_dept")
    with c3:
        date_col = st.selectbox("Date", [NONE] + cols,
            index=_opt_idx(roles.get("date")), key="sr_date")
        dims_sel = st.multiselect(
            "Breakdown dimensions (contract flags, status, etc.)",
            options=cols, default=roles.get("dimensions", [])[:6],
            help="Any low-cardinality column becomes a 'Spend by …' breakdown.",
            key="sr_dims")
    amount_col = None if amount_col == NONE else amount_col
    vendor_col = None if vendor_col == NONE else vendor_col
    dept_col = None if dept_col == NONE else dept_col
    date_col = None if date_col == NONE else date_col

    roles.update({"description": desc_col, "amount": amount_col, "vendor": vendor_col,
                  "department": dept_col, "date": date_col, "dimensions": dims_sel})

    with st.expander("What the tool will analyze for this file (and what it will skip)"):
        for p in sr.plan_analyses({"roles": roles}):
            if p["feasible"]:
                st.markdown(f"✅ **{p['label']}**")
            else:
                st.markdown(f"⤫ {p['label']} — _skipped: {p['reason']}_")

    max_rows = len(df_in)
    if len(df_in) > 25000:
        st.warning(
            f"Large file ({len(df_in):,} rows). Classifying every row can take a couple of "
            "minutes. You can cap the row count for a faster first look.")
        max_rows = st.number_input(
            "Rows to analyze", min_value=1000, max_value=len(df_in),
            value=min(25000, len(df_in)), step=5000, key="sr_maxrows")

    if st.button("Generate spend report", type="primary"):
        work = df_in.head(int(max_rows)).copy()
        with st.spinner(f"Classifying {len(work):,} rows (rules only)…"):
            cls = sr.classify_series(work[desc_col], get_rules())
            work = pd.concat([work.reset_index(drop=True), cls.reset_index(drop=True)], axis=1)

        b = sr.compute_all(work, roles)
        buf = io.BytesIO()
        sr.build_excel_report(buf, tiles=b["tiles"], cat_tbl=b["cat"], pareto_tbl=b["par"],
                              vendors_tbl=b["vend"], consolidation_tbl=b["cons"], cov=b["cov"],
                              exec_summary=b["es"], dept_tbl=b["dept_tbl"], trend_tbl=b["trend"],
                              tail=b["tail"], concentration=b["concentration"],
                              single_multi=b["single_multi"], matrix_tbl=b["matrix"],
                              dimensions=b["dimensions"], item_consol_tbl=b.get("item_consol"),
                              nigp_consol_tbl=b.get("nigp_consol"))
        # Store the whole computed bundle so it survives page navigation.
        st.session_state["sr_result"] = {
            "b": b, "excel": buf.getvalue(), "file_sig": file_sig, "file_name": uploaded.name}
        result = st.session_state["sr_result"]

    # Render if we have a report for THIS uploaded file (freshly generated or saved).
    if result is not None and result.get("file_sig") == file_sig:
        _render_spend_report(result)
    elif result is not None:
        st.caption(
            "You've uploaded a different file. Click **Generate spend report** to analyze it. "
            "Your previous report is still saved — it reappears if you navigate away, or if you "
            "re-upload that file.")


# =========================================================================
# PAGE: Spend Report Methodology & How-To
# =========================================================================
def page_spend_report_methodology() -> None:
    st.title("Spend Report — Methodology & How-To")
    st.markdown(f"**Created by:** {AUTHOR_NAME}")
    st.markdown("---")
    st.markdown(
        "The **Spend Report** page turns a raw spend file into a categorized spend "
        "analysis. This page explains — for leadership and staff — exactly how it works, "
        "how to run it, how to read each result, and where its limits are. "
        "**Everything here is deterministic: plain arithmetic on your data. No AI runs, "
        "and your file never leaves the session.**"
    )

    st.subheader("What it produces — the pipeline")
    st.markdown(
        "Upload a spend file **in any format** and the tool runs four stages, in order:\n"
        "1. **Profile every column** — it inspects each column (name *and* the data inside it) "
        "and assigns a role: description, amount, vendor, department, date, or a breakdown "
        "dimension (contract flag, status, etc.). You can override any guess.\n"
        "2. **Plan** — from the roles it found, it decides which analyses this particular file "
        "can support, and **skips the ones it can't** (e.g., no date column → no trend), telling "
        "you why.\n"
        "3. **Classify + analyze** — the mapper's keyword rules add a Business Category to every "
        "row (rules only, no AI), then it computes every feasible analysis with `groupby` + "
        "share-of-total math.\n"
        "4. **Deliver** — results render on screen and download as one formatted Excel workbook "
        "(a tab per analysis, with charts, in the Chicago palette)."
    )
    st.markdown(
        f"<div style='border-left: 4px solid {CHI_BLUE}; padding: 10px 14px; "
        f"background: {CHI_LT_BLUE}; margin: 8px 0; font-size:14px;'>"
        "<strong>It adapts to the data, not the other way around.</strong> Drop in any "
        "procurement spend file — the tool figures out what's there and runs whatever the data "
        "supports. It never invents data it doesn't have; if a column isn't present, the "
        "dependent analysis is simply skipped with a note."
        "</div>",
        unsafe_allow_html=True,
    )

    st.subheader("How to run it (step by step)")
    st.markdown(
        "1. Go to the **Spend Report** page.\n"
        "2. **Upload** a CSV or Excel spend file.\n"
        "3. **Map your columns** — confirm the auto-detected Description (required), Amount, "
        "Vendor, and Department columns, or pick the right ones from the dropdowns.\n"
        "4. Click **Generate spend report**. (Large files show a row cap you can adjust.)\n"
        "5. Read the on-screen reports, then click **Download full spend report (Excel)**."
    )
    st.markdown(
        f"<div style='border-left: 4px solid {CHI_BLUE}; padding: 10px 14px; "
        f"background: {CHI_LT_BLUE}; margin: 8px 0 16px 0; font-size:14px;'>"
        "<strong>Minimum you need:</strong> a <em>description</em> column (to classify). "
        "An <em>amount</em> column unlocks dollar analysis; a <em>vendor</em> column unlocks "
        "top-vendors and consolidation; a <em>department</em> column enriches the "
        "consolidation view. Without them, the report gracefully falls back to transaction "
        "counts."
        "</div>",
        unsafe_allow_html=True,
    )

    st.subheader("The reports it can produce — what each is, and how to read it")
    st.caption("Which of these run depends on what your file contains — the tool includes every "
               "one the data supports and skips the rest.")
    _spend_report_methods_table()

    st.subheader("How the numbers are calculated (no black box)")
    st.markdown(
        "- **Spend by Category** = `sum(amount)` and row count grouped by Business Category, "
        "with each category's **% of total** and an **Examples** column of representative "
        "descriptions. Every row lands in a real category — anything without a specific-commodity "
        "rule match falls into **General & Other Procurement**, so the total always reconciles to "
        "your file.\n"
        "- **Pareto 80/20** = the specific-commodity category totals sorted high-to-low, with a "
        "running **cumulative %**; the report names how many categories it takes to reach 80%. "
        "*The General & Other catch-all is excluded here* — it isn't a single commodity to focus "
        "sourcing on.\n"
        "- **Top Vendors** = `sum(amount)` grouped by vendor, sorted, with % of total.\n"
        "- **Consolidation / Fragmentation** = for each category, the **distinct vendor "
        "count** (and department count); categories bought from more than one vendor are "
        "flagged and ranked by spend — the biggest consolidation opportunities first.\n"
        "- **Same Item — Multiple Vendors / Departments** = group rows by the *exact item "
        "description*, then for each item count the **distinct vendors** and **distinct "
        "departments** and sum spend; keep only items bought from >1 vendor or across >1 "
        "department, ranked by spend. Vendor and department names are shown so you can act on "
        "it. (Grouping is exact text — it doesn't fuzzy-match — so a flag is real, not inferred.)\n"
        "- **Same Commodity (NIGP code)** = the same calculation, but grouped by the **NIGP "
        "commodity code** (5-digit item where the rule assigned one, otherwise 3-digit class) "
        "instead of the wording — so two differently-worded descriptions of the same commodity "
        "roll together. Presented as a matrix: **each department is its own column** (spend), the "
        "blue cells show which departments bought the commodity, and rows that span two or more "
        "departments are flagged **red** as consolidation targets. The top consolidatable items "
        "are also listed in the Executive Summary.\n"
        "- **Spend Trend** = `sum(amount)` grouped by the year of the date column, with "
        "year-over-year % change.\n"
        "- **Vendor Concentration** = the spend share of the top 1/5/10 vendors, plus an "
        "**HHI** (sum of squared vendor shares × 10,000; higher = more concentrated).\n"
        "- **Tail Spend** = the vendors *outside* the top group that makes up 80% — counted and "
        "summed to show how little of total spend the long tail represents.\n"
        "- **Spend by Dimension** = `sum(amount)` grouped by any categorical column found "
        "(contract flag, diversity, status). Contract/diversity flags are recognized and "
        "labeled automatically.\n"
        "- **Coverage** = how many rows mapped to a specific commodity category vs. fell into "
        "**General & Other Procurement** (the catch-all). This is the report's honesty check — "
        "and the Examples column shows exactly what's in the catch-all."
    )
    st.markdown(
        "Currency text is cleaned automatically (`$`, commas, and parenthesized negatives "
        "are handled). Amounts that can't be parsed are treated as zero, not guessed."
    )

    st.subheader("Why you can trust it")
    st.markdown(
        "- **Deterministic.** Every number is a `groupby`/sum you could reproduce in Excel — "
        "there is no model opinion in the analysis.\n"
        "- **Classification matches the rest of the tool.** The Spend Report uses the exact "
        "same rule engine as the Classify page (verified identical), so a category here means "
        "the same thing everywhere.\n"
        "- **Honest coverage.** The report never hides what it couldn't classify — it counts "
        "it and shows it.\n"
        "- **Private.** The file is processed in-session and is not stored."
    )

    st.subheader("Limits — stated plainly")
    st.markdown(
        "- **Coverage depends on description quality.** Classification is keyword-rules-only "
        "and was tuned on one agency's vocabulary. A new dataset — or one with terse "
        "descriptions — will classify a lower share; the coverage note shows exactly how "
        "much. Raise it by adding rules.\n"
        "- **Dollars are analyzed, never used to classify.** Spend is an output of the "
        "analysis, not an input to the category decision.\n"
        "- **Consolidation needs a vendor column** (and ideally department) to be meaningful.\n"
        "- **Large files** (100K+ rows) take a couple of minutes and can be row-capped for a "
        "faster first look.\n"
        "- **This is the deterministic first version.** A written executive-summary narrative "
        "and Word/PDF export are planned later iterations."
    )

    st.subheader("Need help?")
    st.markdown(f"Contact **{AUTHOR_NAME.split(',')[0]}**.")


def _spend_report_methods_table() -> None:
    st.markdown(
        f"""
        <div style="margin: 8px 0 4px 0;">
          <table style="width:100%; border-collapse: collapse; font-size: 14px;">
            <thead>
              <tr style="background:{CHI_NAVY}; color:white;">
                <th style="padding:8px 10px; text-align:left;">Report</th>
                <th style="padding:8px 10px; text-align:left;">What it shows</th>
                <th style="padding:8px 10px; text-align:left;">Decision it supports</th>
              </tr>
            </thead>
            <tbody>
              <tr style="background:{CHI_LT_BLUE};">
                <td style="padding:8px 10px;"><strong>Summary tiles</strong></td>
                <td style="padding:8px 10px;">Total spend, transactions, categories, vendors.</td>
                <td style="padding:8px 10px;">Instant scale check.</td>
              </tr>
              <tr>
                <td style="padding:8px 10px;"><strong>Spend by Category</strong></td>
                <td style="padding:8px 10px;">Dollars &amp; share by the 17 Business Categories.</td>
                <td style="padding:8px 10px;">Where the money actually goes, by <em>what</em> is bought.</td>
              </tr>
              <tr style="background:{CHI_LT_BLUE};">
                <td style="padding:8px 10px;"><strong>Pareto 80/20</strong></td>
                <td style="padding:8px 10px;">The few categories that drive ~80% of spend.</td>
                <td style="padding:8px 10px;">Where to focus sourcing effort first.</td>
              </tr>
              <tr>
                <td style="padding:8px 10px;"><strong>Top Vendors</strong></td>
                <td style="padding:8px 10px;">Largest vendors by spend and share.</td>
                <td style="padding:8px 10px;">Vendor rationalization candidates.</td>
              </tr>
              <tr style="background:{CHI_LT_BLUE};">
                <td style="padding:8px 10px;"><strong>Consolidation / Fragmentation</strong></td>
                <td style="padding:8px 10px;">Categories bought from many vendors (and departments).</td>
                <td style="padding:8px 10px;">The savings story — where to consolidate demand.</td>
              </tr>
              <tr>
                <td style="padding:8px 10px;"><strong>Same Item — Multiple Vendors / Departments</strong></td>
                <td style="padding:8px 10px;">The <em>same item</em> (by description) bought from more than one vendor and/or across more than one department — with the vendor and department names.</td>
                <td style="padding:8px 10px;">Line-level maverick buying — the sharpest, most defensible consolidation targets.</td>
              </tr>
              <tr style="background:{CHI_LT_BLUE};">
                <td style="padding:8px 10px;"><strong>Same Commodity (NIGP code)</strong></td>
                <td style="padding:8px 10px;">Groups by the <em>NIGP commodity code</em> (5-digit item where present, else 3-digit class) so differently-worded descriptions of the same commodity roll together — shows the merged wordings.</td>
                <td style="padding:8px 10px;">Catches split buying hidden behind wording differences.</td>
              </tr>
              <tr>
                <td style="padding:8px 10px;"><strong>Spend Trend Over Time</strong></td>
                <td style="padding:8px 10px;">Spend by year with year-over-year change. <em>Needs a date column.</em></td>
                <td style="padding:8px 10px;">Is spend rising or falling, and when did it peak?</td>
              </tr>
              <tr style="background:{CHI_LT_BLUE};">
                <td style="padding:8px 10px;"><strong>Vendor Concentration / Risk</strong></td>
                <td style="padding:8px 10px;">Top-1/5/10 vendor share and an HHI concentration score.</td>
                <td style="padding:8px 10px;">How dependent are we on a few suppliers?</td>
              </tr>
              <tr>
                <td style="padding:8px 10px;"><strong>Tail-Spend Analysis</strong></td>
                <td style="padding:8px 10px;">The many small vendors making up the last ~20% of spend.</td>
                <td style="padding:8px 10px;">The classic consolidation target.</td>
              </tr>
              <tr style="background:{CHI_LT_BLUE};">
                <td style="padding:8px 10px;"><strong>Single- vs Multi-Source</strong></td>
                <td style="padding:8px 10px;">Categories bought from one vendor vs. many.</td>
                <td style="padding:8px 10px;">Vendor rationalization &amp; supply risk.</td>
              </tr>
              <tr>
                <td style="padding:8px 10px;"><strong>Spend by Department</strong></td>
                <td style="padding:8px 10px;">Dollars by department (actual names). <em>Needs a department column.</em></td>
                <td style="padding:8px 10px;">Who is spending, and on what.</td>
              </tr>
              <tr style="background:{CHI_LT_BLUE};">
                <td style="padding:8px 10px;"><strong>Category × Department Matrix</strong></td>
                <td style="padding:8px 10px;">A spend cross-tab of category by department.</td>
                <td style="padding:8px 10px;">Cross-department demand for the same commodity.</td>
              </tr>
              <tr>
                <td style="padding:8px 10px;"><strong>Spend by any Dimension</strong></td>
                <td style="padding:8px 10px;">A breakdown for any categorical column it finds — contract vs non-contract, supplier diversity, status, etc.</td>
                <td style="padding:8px 10px;">Off-contract leakage, diversity participation, and more — whatever the file carries.</td>
              </tr>
              <tr style="background:{CHI_LT_BLUE};">
                <td style="padding:8px 10px;"><strong>Coverage</strong></td>
                <td style="padding:8px 10px;">Share of rows classified vs. no-rule / no-description.</td>
                <td style="padding:8px 10px;">How much of the file the analysis actually covers.</td>
              </tr>
            </tbody>
          </table>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================================
# Sidebar + dispatch
# =========================================================================
def render_sidebar() -> str:
    with st.sidebar:
        st.markdown(f"### NIGP-Sourced")
        st.markdown(f"**Procurement Category Mapper**")
        st.markdown(
            f"<div style='color:{CHI_GRAY}; font-size:11px; margin-top:-6px;'>"
            f"Created by<br><strong>{AUTHOR_NAME}</strong></div>",
            unsafe_allow_html=True,
        )
        st.markdown("---")
        st.markdown(f"**Pages**")
        page = st.radio(
            "Navigate to:",
            [
                "Classify",
                "Bulk Classify",
                "Spend Report",
                "Spend Report Methodology",
                "Procurement Taxonomy Logic",
                "Methodology",
                "Business Categories",
                "Rule Lookup",
            ],
            label_visibility="collapsed",
        )
        st.markdown("---")
        st.markdown(
            f"<div style='color:{CHI_GRAY}; font-size:11px;'>"
            f"v1.1 — production run finalized 30 April 2026.<br><br>"
            f"For questions about a specific classification, consult the "
            f"<strong>Methodology</strong> page or contact {AUTHOR_NAME.split(',')[0]}."
            f"</div>",
            unsafe_allow_html=True,
        )
    return page


def render_footer() -> None:
    st.markdown(
        f"<div class='chi-footer'>"
        f"{AUTHOR_TITLE_LINE}  •  NIGP Procurement Taxonomy & Classification Engine v1.1  •  "
        f"Created by {AUTHOR_NAME}"
        f"</div>",
        unsafe_allow_html=True,
    )


def main() -> None:
    apply_css()
    if not check_password():
        st.stop()
    init_session_state()

    page = render_sidebar()
    if page == "Classify":
        page_classify()
    elif page == "Bulk Classify":
        page_bulk()
    elif page == "Spend Report":
        page_spend_report()
    elif page == "Spend Report Methodology":
        page_spend_report_methodology()
    elif page == "Procurement Taxonomy Logic":
        page_taxonomy_logic()
    elif page == "Methodology":
        page_methodology()
    elif page == "Business Categories":
        page_categories()
    elif page == "Rule Lookup":
        page_rule_lookup()

    render_footer()


if __name__ == "__main__":
    main()
