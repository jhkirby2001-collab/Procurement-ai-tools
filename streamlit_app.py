"""
Streamlit web app — NIGP-Sourced Procurement Category Mapper.

Multi-page browser app wrapping classifier_JHK3.py. Authorized users
sign in once, then navigate via the sidebar between:
  - Classify (single description + recent-history list)
  - Bulk Classify (paste-many or CSV upload, downloadable result)
  - Methodology (plain-language overview + "human review queue" explainer)
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
    st.title("How to Use This Tool")
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
        "- **Procurement Taxonomy Logic** — three-tier explainer of how spend rolls up.\n"
        "- **Methodology** — this page.\n"
        "- **Business Categories** — reference list of all 17 categories with definitions.\n"
        "- **Rule Lookup** — search the rule catalog by keyword to see why a description routed where it did."
    )

    st.subheader("Reading a result")
    st.markdown(
        "Each classification shows a confidence badge:\n"
        "- **HIGH** — strong rule match, trust it.\n"
        "- **MEDIUM** — rule fired with caveats, spot-check.\n"
        "- **REVIEW RECOMMENDED** — no confident match; the record is in the human-review queue. "
        "The classifier never guesses."
    )

    st.subheader("How the confidence badge is calculated")
    st.markdown(
        "Every NIGP code assigned by the classifier carries a confidence badge. The badge "
        "reflects how specific the evidence was — not how the classifier \"felt.\" It is set "
        "by the rule that fired, not invented after the fact."
    )
    st.markdown(
        "- **HIGH** — The description matched a specific rule that maps directly to a precise "
        "NIGP 5-digit Item code (e.g., \"BULK ROAD SALT\" → NIGP 285-83). The NIGP assignment "
        "is exact, defensible, and traceable to a single rule.\n"
        "- **MEDIUM** — The description matched a broader rule that resolves to an NIGP 3-digit "
        "Class but not a specific 5-digit Item (e.g., generic \"PIPE FITTINGS\" lands in the "
        "plumbing class without pinning the exact item). The Business Category and 3-digit "
        "Class are reliable; the 5-digit Item is best-effort. Spot-check recommended for "
        "high-dollar transactions.\n"
        "- **REVIEW RECOMMENDED** — No rule fired with sufficient specificity, OR the matching "
        "rule is itself flagged for review. The classifier does **not** assign a guess. The row "
        "is routed to the human-review queue and any NIGP code shown is provisional. This is "
        "by design — the system never invents a code."
    )
    st.markdown(
        "Every classification is auditable. The exact rule that fired is recorded on every row, "
        "so any NIGP code can be traced back to the rule that produced it."
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

    st.subheader("How the three tiers fit together")
    st.markdown(
        f"""
        <div style="max-width: 520px; margin: 24px auto;">
          <div style="background:{CHI_NAVY}; color:white; padding:18px 20px; border-radius:8px; text-align:center;">
            <div style="font-size:12px; letter-spacing:1px; opacity:0.85;">TIER 1</div>
            <div style="font-size:20px; font-weight:700; margin-top:4px;">Executive View</div>
            <div style="font-size:14px; margin-top:6px;">17 Business Categories</div>
            <div style="font-size:12px; opacity:0.85; margin-top:4px;">Audience: Leadership &amp; Budget Decisions</div>
          </div>
          <div style="text-align:center; color:{CHI_GRAY}; font-size:24px; line-height:1; margin:6px 0;">▼</div>
          <div style="background:{CHI_BLUE}; color:{CHI_NAVY}; padding:18px 20px; border-radius:8px; text-align:center;">
            <div style="font-size:12px; letter-spacing:1px; opacity:0.9;">TIER 2</div>
            <div style="font-size:20px; font-weight:700; margin-top:4px;">Sourcing View</div>
            <div style="font-size:14px; margin-top:6px;">138 NIGP 3-digit Classes</div>
            <div style="font-size:12px; opacity:0.9; margin-top:4px;">Audience: Category Managers &amp; Sourcing Teams</div>
          </div>
          <div style="text-align:center; color:{CHI_GRAY}; font-size:24px; line-height:1; margin:6px 0;">▼</div>
          <div style="background:{CHI_LT_BLUE}; color:{CHI_NAVY}; padding:18px 20px; border-radius:8px; text-align:center; border:1px solid {CHI_BLUE};">
            <div style="font-size:12px; letter-spacing:1px;">TIER 3</div>
            <div style="font-size:20px; font-weight:700; margin-top:4px;">Audit View</div>
            <div style="font-size:14px; margin-top:6px;">784,556 Individual Transactions</div>
            <div style="font-size:11px; font-style:italic; color:#E57373; margin-top:4px; padding:0 6px;">PO/invoice line records in the historical dataset the classifier was validated against (October 2002 – May 2025).</div>
            <div style="font-size:12px; margin-top:6px;">Audience: Auditors, Analysts &amp; Anyone Asking "Why?"</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.subheader("Tier 1 — Executive View")
    st.markdown(
        "**17 Business Categories** that show where procurement dollars go. This is the "
        "strategic level for leadership attention and budget decisions. Categories are "
        "**mutually exclusive** (no record fits two) and **collectively exhaustive** "
        "(every record fits one)."
    )

    st.subheader("Tier 2 — Sourcing View")
    st.markdown(
        "**Subcategories within each Business Category**, aligned to NIGP 3-digit "
        "commodity classes. This is where consolidation opportunities become visible — "
        "when the same commodity is purchased across multiple departments from multiple "
        "vendors without coordination."
    )

    st.subheader("Tier 3 — Audit View")
    st.markdown(
        "**Individual transactions** with full classification detail including the NIGP "
        "code, confidence level, and the exact rule that fired. Every classification is "
        "transparent and defensible — no black-box decisions."
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
