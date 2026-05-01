"""
Streamlit web app — Chicago NIGP Procurement Classifier.

Multi-page browser app wrapping classifier_JHK3.py. Authorized DPS staff
sign in once, then navigate via the sidebar between:
  - Classify (single description + recent-history list)
  - Bulk Classify (paste-many or CSV upload, downloadable result)
  - Methodology (plain-language overview + "human review queue" explainer)
  - Business Categories (the 17-category reference)
  - Rule Lookup (search the curated + AI-mined rule catalog)

Author: James H. Kirby III, CSCP, MS-SCM
Department of Procurement Services, City of Chicago

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

# -- Brand palette ---------------------------------------------------------
CHI_NAVY = "#002F6C"
CHI_BLUE = "#41B6E6"
CHI_RED = "#DA291C"
CHI_LT_BLUE = "#D6EEF9"
CHI_GREEN = "#0E7C3A"
CHI_GRAY = "#595959"

AUTHOR_NAME = "James H. Kirby III, CSCP, MS-SCM"
AUTHOR_TITLE_LINE = "City of Chicago — Department of Procurement Services"

# -- Page config -----------------------------------------------------------
st.set_page_config(
    page_title="Chicago DPS — NIGP Classifier",
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

    st.title("Chicago DPS — NIGP Classifier")
    st.markdown(
        f"**Authorized Department of Procurement Services use only.**  "
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
        badge_text = "NO MATCH — HUMAN REVIEW"
        category_display = "Sent to human-review queue"
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
            f"NIGP Code:  <em>none assigned</em> — no rule fired, sent to "
            f"human-review queue for triage."
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
        df_in = pd.read_csv(uploaded)
    except Exception as e:
        st.error(f"Could not read the CSV: {e}")
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
    st.title("Methodology")
    st.markdown(
        "**Project:** NIGP-Aligned Procurement Taxonomy & Classification Engine  \n"
        "**Document version:** 1.1 — finalized 30 April 2026  \n"
        "**Prepared for:** City of Chicago Department of Procurement Services Leadership"
    )

    st.markdown(
        f"""
        <div class='chi-author-card'>
          <strong style='color:{CHI_NAVY};'>Project Author:</strong> {AUTHOR_NAME}<br>
          <span style='color:{CHI_GRAY}; font-size:13px;'>
            Department of Procurement Services, City of Chicago.
            28 years in public-sector procurement. Designed and built the taxonomy,
            the classification engine, the production rule base, and this web tool.
          </span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ---- Executive Summary ----
    st.header("Executive Summary")
    st.markdown(
        "The City of Chicago has not historically maintained its own commodity "
        "classification system for procurement spend. The EY raw data file analyzed "
        "here is a consulting deliverable — useful as raw transactional data, but "
        "**not** an authoritative City of Chicago classification of what the City buys."
    )
    st.markdown(
        "This project delivers Chicago's **first internally-owned commodity taxonomy** "
        "plus a **reusable, defensible classification engine** that the City can run "
        "on this historical file and on every future procurement extract. The "
        "deliverable is fully self-contained: no recurring vendor dependencies, no "
        "per-classification API calls, no third-party software requirements beyond "
        "standard open-source Python."
    )
    st.markdown(
        "**Headline outcome (production run, 30 April 2026):** 86.4% of all 784,556 "
        "rows were auto-classified by deterministic rule. End-to-end runtime on a "
        "standard workstation was 14 minutes 37 seconds. The remaining 17.8% are "
        "routed to a procurement-staff review queue rather than being guessed."
    )

    # ---- 1. Project Objective and Scope ----
    st.header("1.  Project Objective and Scope")
    st.markdown(
        "Build a transparent, defensible, accuracy-first classification engine that:"
    )
    st.markdown(
        "- **Uses the NIGP commodity code framework** as the public-standard backbone "
        "for inter-agency comparability and audit defensibility.\n"
        "- **Layers a custom Business Category rollup** on top, designed for "
        "Chicago's organizational reporting needs and business-friendly leadership review.\n"
        "- **Operates in two modes:** batch processing of large historical files, and "
        "single-record classification for live PO and requisition work.\n"
        "- **Prioritizes accuracy over auto-coding.** Records that cannot be classified "
        "with sufficient confidence are flagged for human review rather than guessed.\n"
        "- **Is fully repeatable** on future raw extracts with no code modification — "
        "the rule files are externalized as CSVs that procurement staff can edit directly."
    )

    # ---- 2. Source Data ----
    st.header("2.  Source Data")
    st.markdown(
        "**File analyzed:** EY raw data extract (`ey raw data.xlsx`, 326 MB).  \n"
        "**Records:** 784,556 purchase order and invoice line items.  \n"
        "**Time span:** 2002 – 2025 (23 years of Chicago spend).  \n"
        "**EY's prior NIGP labels (~30% of rows):** preserved in the raw file but "
        "**NOT used as classifier input.** Chicago is building its own classification "
        "independently from description text."
    )
    st.markdown(
        "**Data quality highlights (from profiling):**"
    )
    st.markdown(
        "- Line-level descriptions are populated on virtually all rows. Only 967 of "
        "784,556 rows (0.1%) have no usable description across any of the four description fields.\n"
        "- 17.8% of high-level `PO Description` values are uninformative "
        "(`Misc`, `Per contract`). The classifier falls back to line-level descriptions, "
        "which carry substantive content.\n"
        "- Two date fields contained typos or open-ended placeholders. Documented; no "
        "impact on classification."
    )

    # ---- 3. Three-Level Taxonomy ----
    st.header("3.  Three-Level Taxonomy Design")
    st.markdown(
        "Every classified record carries three taxonomy levels. This serves both "
        "executive dashboards (Business Category) and sourcing/audit needs (NIGP codes)."
    )
    st.markdown("**Level 1 — Business Category (Custom Chicago Rollup, 17 buckets).** "
                "Mutually exclusive, collectively exhaustive, and immediately recognizable "
                "to City leadership and audit reviewers. See the **Business Categories** "
                "page for the full list with definitions.")
    st.markdown("**Level 2 — NIGP 3-digit Class.** 138 distinct classes appear in the EY "
                "data, mapped 1:1 to the 17 Business Categories. NIGP is a public, "
                "inter-agency-compatible commodity framework — aligning to it enables "
                "peer benchmarking, audit defensibility, and future portability.")
    st.markdown("**Level 3 — NIGP 5-digit Class-Item.** When a description supports "
                "specificity, the classifier assigns a 5-digit code (470 distinct items "
                "in the working catalog). When confidence at 5-digit is insufficient, "
                "the row carries the broader 3-digit Class only.")

    # ---- 4. Classifier Inputs ----
    st.header("4.  Inputs to the Classifier")
    st.markdown(
        "**Inputs the classifier USES:**"
    )
    st.markdown(
        "- **Transaction description text** — `PO Description`, `PO Item Description`, "
        "`AP Invoice Line Description`, `Invoice Distribution Description`. The classifier "
        "selects the first non-blank, substantive description across these four fields.\n"
        "- **Chicago FMPS account/object/fund codes** — used as supplemental signal "
        "where description text alone is ambiguous."
    )
    st.markdown(
        "**Inputs explicitly EXCLUDED by design:**"
    )
    st.markdown(
        "- **Vendor name.** The same vendor often sells across many commodity categories. "
        "Vendor-based inference introduces misclassification risk that descriptions and "
        "account codes do not.\n"
        "- **EY-supplied NIGP codes.** These represent EY's prior classification work, "
        "not Chicago's authoritative judgment. Chicago is building its own classification "
        "independently."
    )

    # ---- 5. Decision Pipeline ----
    st.header("5.  Rule Hierarchy and Order of Operations")
    st.markdown(
        "The classifier evaluates each record through a deterministic three-step pipeline. "
        "**The first rule that fires assigns the classification — subsequent rules don't "
        "run for that row.** This produces a clean audit trail: every classified record "
        "carries the exact rule that drove the decision."
    )
    st.markdown("**Pass 1 — Keyword Rules on Description Text.** "
                "148 hand-curated rules drafted by procurement leadership take priority "
                "within any given match-type tier, plus 6,766 AI-mined rules harvested "
                "during the build. Three match types are supported: `exact`, `starts_with`, "
                "and `contains`. Match priority is `exact > starts_with > contains`.")
    st.markdown("**Pass 2 — Account-Code Patterns (Chicago FMPS supplemental signal).** "
                "Maps Chicago FMPS account codes to (Business Category, NIGP Class). "
                "Most consequential is the 220xxx subgrant series (220005, 220044, "
                "220100, 220801, 220300, 220999) — empirically 93–100% used for "
                "subgrant disbursements regardless of description text quality.")
    st.markdown("**Pass 3 — Human Review.** "
                "Records that don't match a keyword rule or account pattern are flagged "
                "with `Classification_Method = human_review` and `Review_Flag = Yes`. "
                "**The classifier never guesses.**")

    # ---- 6. Human Review Queue ----
    st.header("6.  The Human Review Queue — Plain-Language Definition")
    st.markdown(
        f"""
        <div class='chi-callout'>
          <div class='chi-callout-label'>What it is and how it works</div>
          <p style='margin-top:8px; margin-bottom:0;'>
            When a description doesn't confidently match any rule, the classifier
            sets it aside instead of guessing. About <strong>14% of all historical
            records</strong> — roughly 106,000 rows out of 784,556 — are in this
            queue today. These are typically thin descriptions like
            <em>"Misc supplies"</em> or <em>"Per contract"</em>, or department-specific
            codes that haven't been encoded as rules yet.
            <br><br>
            <strong>Triage cadence:</strong> Procurement staff triage the queue on a
            monthly or quarterly cadence. Each triage decision becomes a new rule
            in <code>keyword_rules_DRAFT_JHK3.csv</code>, so the queue <strong>shrinks
            over time</strong> and the classifier's confidence grows.
            <br><br>
            <strong>Why "REVIEW RECOMMENDED" is a feature, not a failure:</strong>
            The alternative — guessing on weak signal — would produce classification
            noise that misleads spend reports, sourcing analysis, and audit reviewers.
            Better to flag and triage than guess and mislead.
          </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ---- 7. AI Use ----
    st.header("7.  Use of Artificial Intelligence — Limited, Transparent, One-Time")
    st.markdown(
        f"""
        <div class='chi-callout' style='border-left-color:{CHI_NAVY};'>
          <div class='chi-callout-label' style='color:{CHI_NAVY};'>Defensibility upfront</div>
          <p style='margin-top:8px; margin-bottom:0;'>
            <strong>The production classification system Chicago operates is rules-only.</strong>
            No AI is in the runtime path. Any City staff member can run the classifier
            from a workstation with no API key, no internet dependency, no recurring
            vendor cost, and no per-classification charge.
          </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        "AI was used **once during the initial taxonomy build** for one bounded purpose: "
        "to mine recurring patterns from the long tail of descriptions where no "
        "hand-curated rule applied. The AI's output was harvested into the keyword "
        "rules file and then frozen."
    )
    st.markdown("**Model:** Anthropic Claude Haiku 4.5.")
    st.markdown("**Invocation:** A single batch run over 30,342 unique long-tail descriptions.")
    st.markdown(
        "**Constraints:** The AI's output was constrained by JSON schema with a closed "
        "enumeration of the 17 Business Categories and 138 NIGP 3-digit Classes. The "
        "model could not invent codes outside the catalog."
    )
    st.markdown(
        "**Input visibility:** Each AI call sent only the description text plus a system "
        "prompt containing the controlled vocabulary. Vendor name, EY codes, and Chicago "
        "account codes were NOT passed to the AI."
    )
    st.markdown("**Promotion thresholds:**")
    st.markdown(
        "- High-confidence AI proposals (6,419 / 21.2%) → auto-promoted into the rule file.\n"
        "- Medium-confidence proposals (11,645 / 38.4%) → promoted only if the description "
        "recurred in 5+ source rows. 1,998 met the bar; 9,647 long-tail singletons were dropped.\n"
        "- Low-confidence proposals (12,278 / 40.5%) → never promoted. Those descriptions "
        "remain in the human-review queue.\n"
        "- **Net:** 6,766 AI-mined rules promoted out of 30,342 candidates (22.3%)."
    )
    st.markdown("**Why this AI use is defensible:**")
    st.markdown(
        "1. **One-time, not ongoing.** Production runs do not call AI. Once the rules "
        "are accepted, the AI dependency is severed.\n"
        "2. **Bounded by Chicago's own taxonomy.** The AI cannot output codes outside "
        "the 17 categories and 138 classes that Chicago's procurement leadership approved.\n"
        "3. **Transparent.** Every AI-mined rule is auditable. The source CSV "
        "(`keyword_rules_from_ai_JHK3.csv`) carries provenance metadata in every row.\n"
        "4. **Reviewable.** Procurement staff can edit, demote, or remove any AI-mined rule.\n"
        "5. **Replaceable.** If Chicago later decides to remove all AI-mined content, "
        "the file can be deleted and the classifier still operates on hand-curated rules alone."
    )

    # ---- 8. Confidence ----
    st.header("8.  Confidence, Match Levels, and Audit Trail")
    st.markdown(
        "Every classified record carries a confidence triple plus a reason — a complete "
        "decision record:"
    )
    st.markdown(
        "- **`NIGP_Match_Level`** — `exact` (5-digit Class-Item assignable), `broad` "
        "(3-digit Class only), `review` (best broader category but flagged), or empty "
        "(no match — Grants category).\n"
        "- **`Classification_Confidence`** — `High` / `Medium` / `Low`.\n"
        "- **`Confidence_Score`** — numeric value from 0.0 to 1.0. Records scoring below "
        "0.75 are auto-flagged with `Review_Flag = Yes` regardless of method.\n"
        "- **`Classification_Reason`** — records the exact rule pattern that fired and "
        "any notes; provides a complete audit trail end-to-end."
    )

    # ---- 9. Production Results ----
    st.header("9.  Production Run Results — 30 April 2026")
    st.markdown(
        "End-to-end run against all 784,556 rows. Wall-clock runtime: 14 minutes 37 seconds."
    )
    st.markdown(
        "| Outcome | Pre-AI baseline | Final | Change |\n"
        "|---|---|---|---|\n"
        "| Auto-classified | 609,541 (77.7%) | **678,085 (86.4%)** | +68,544 / +8.7 pp |\n"
        "| Sent to human review | 175,015 (22.3%) | 106,471 (13.6%) | −68,544 |\n"
        "| Review-flag QA queue | 408,037 (52.0%) | **139,868 (17.8%)** | −268,169 / −34.2 pp |"
    )
    st.markdown(
        "- 148 curated rules accounted for 609,889 row-hits (90.2% of classified rows).\n"
        "- 6,766 AI-mined rules accounted for 66,017 row-hits (9.8%), averaging ~10 rows "
        "per rule — the long-tail role they were designed for.\n"
        "- The largest review-flag improvement came from AI exact-match rules superseding "
        "broader curated `contains` rules (~200,000 rows reclassified out of the QA queue)."
    )

    # ---- 10. Limitations ----
    st.header("10.  Honest Limitations")
    st.markdown(
        "1. **Single-source dataset.** The taxonomy was derived from one consulting "
        "deliverable. New Chicago extracts may reveal commodity types not represented "
        "in the EY file.\n"
        "2. **Time-agnostic by design.** The classifier ignores transaction date. 23 "
        "years of inflation, contract restructuring, and reorganizations don't affect "
        "classification consistency.\n"
        "3. **No vendor signal.** Classification reflects what was bought, not who sold "
        "it. Vendor-based views can be produced separately from the classified output.\n"
        "4. **Description-quality dependent.** Thin descriptions (`Misc supplies`) "
        "cannot be classified from text alone and are correctly routed to human review "
        "rather than guessed.\n"
        "5. **NIGP working set is partial.** The 138 classes derived from the EY file "
        "are a subset of the full ~9,000-code NIGP standard. Future work should consider "
        "licensing the full Periscope catalog."
    )

    # ---- 11. Governance ----
    st.header("11.  Governance and Future Maintenance")
    st.markdown(
        "1. **Annual taxonomy review.** Procurement leadership reviews the 17 Business "
        "Categories annually for continued fit with reporting needs. Edits go directly "
        "into `business_categories_JHK3.csv`.\n"
        "2. **Rule-file ownership.** Designate a procurement analyst as owner of "
        "`keyword_rules_DRAFT_JHK3.csv`. Add new rules as new commodity types appear; "
        "retire outdated ones.\n"
        "3. **Review-queue triage cadence.** Triage the review-flagged subset on a "
        "monthly or quarterly cadence. Triaged decisions become new rules — the queue "
        "shrinks over time.\n"
        "4. **Single-record feedback loop.** When DPS workflow integration is built, "
        "analyst overrides of single-record proposals become candidate rule data.\n"
        "5. **Periscope NIGP catalog.** Consider a commercial license to expand the "
        "working catalog from 138 codes to the full ~9,000-code NIGP standard.\n"
        "6. **Annual independent audit benchmark.** Sample 100 classified rows at random; "
        "have a procurement analyst independently classify them blind. Track agreement rate."
    )

    # ---- Where to go for more ----
    st.markdown("---")
    st.markdown(
        f"**For the canonical, full-detail methodology document** — including the "
        f"locked-decisions appendix, NIGP catalog references, and technical architecture "
        f"specs — see `spend-analysis/METHODOLOGY_JHK3.md` in the project repository, or "
        f"the leadership-formatted Word version at "
        f"`outputs/NIGP_Methodology_for_Leadership_JHK3.docx`."
    )


# =========================================================================
# PAGE: Business Categories
# =========================================================================
def page_categories() -> None:
    st.title("17 Business Categories — Reference")
    st.markdown(
        "Every record is assigned to exactly one of these 17 categories. The categories "
        "are designed to be **mutually exclusive** (no record fits two) and **collectively "
        "exhaustive** (every Chicago purchase fits one)."
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
# Sidebar + dispatch
# =========================================================================
def render_sidebar() -> str:
    with st.sidebar:
        st.markdown(f"### Chicago DPS")
        st.markdown(f"**NIGP Classifier**")
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
    elif page == "Methodology":
        page_methodology()
    elif page == "Business Categories":
        page_categories()
    elif page == "Rule Lookup":
        page_rule_lookup()

    render_footer()


if __name__ == "__main__":
    main()
