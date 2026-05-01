"""
Streamlit web app — Chicago NIGP Procurement Classifier.

A password-gated, browser-based wrapper around classifier_JHK3.py. Lets
authorized DPS staff paste a description and immediately see the proposed
Business Category, NIGP code, confidence, and reason — without a terminal,
Python, or any technical setup.

DEPLOYMENT (one-time):
  1. Push this repo to GitHub (already done).
  2. Visit https://share.streamlit.io and sign in with GitHub.
  3. Click "New app", select this repo, branch `main`, main file `streamlit_app.py`.
  4. Click "Advanced settings" → "Secrets" and paste:
         password = "YOUR-CHOSEN-PASSWORD"
  5. Click "Deploy". Live URL is generated within ~2 minutes.
  6. Share the URL and the password with authorized DPS staff.

LOCAL TESTING:
  Create .streamlit/secrets.toml with:
      password = "test-password"
  Then run:
      streamlit run streamlit_app.py
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import streamlit as st

# Make spend-analysis/scripts importable so we can call into classifier_JHK3
ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "spend-analysis" / "scripts"))
from classifier_JHK3 import classify_one, load_keyword_rules  # noqa: E402

# -- Brand palette (City of Chicago) ---------------------------------------
CHI_NAVY = "#002F6C"
CHI_BLUE = "#41B6E6"
CHI_RED = "#DA291C"
CHI_LT_BLUE = "#D6EEF9"
CHI_GREEN = "#0E7C3A"

st.set_page_config(
    page_title="Chicago DPS — NIGP Classifier",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# -- Page styling ----------------------------------------------------------
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
      .chi-badge {{
        display: inline-block;
        padding: 2px 8px;
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
      .chi-footer {{
        color: #777;
        font-size: 11px;
        text-align: center;
        margin-top: 24px;
        padding-top: 12px;
        border-top: 1px solid #ddd;
      }}
    </style>
    """,
    unsafe_allow_html=True,
)


# -- Password gate ---------------------------------------------------------
def check_password() -> bool:
    """Block the app behind a shared password stored in st.secrets."""
    if st.session_state.get("password_correct"):
        return True

    st.title("Chicago DPS — NIGP Classifier")
    st.markdown(
        "**Authorized Department of Procurement Services use only.**  "
        "Enter the access password to continue."
    )

    pw = st.text_input("Password", type="password", key="pw_input")
    if st.button("Sign in"):
        try:
            expected = st.secrets["password"]
        except (KeyError, FileNotFoundError):
            st.error(
                "This server is missing the access-password configuration. "
                "Contact the administrator (J. Kirby) to set the `password` secret."
            )
            return False
        if pw == expected:
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.markdown(
        f"<div class='chi-footer'>"
        f"City of Chicago  •  Department of Procurement Services  •  "
        f"NIGP Classifier v1.1"
        f"</div>",
        unsafe_allow_html=True,
    )
    return False


if not check_password():
    st.stop()


# -- Cached rule loader ----------------------------------------------------
@st.cache_resource(show_spinner="Loading classification rules...")
def get_rules() -> pd.DataFrame:
    """Load both hand-curated and AI-mined rule files once per server lifetime."""
    return load_keyword_rules()


# -- Authenticated app -----------------------------------------------------
st.title("Chicago DPS — NIGP Classifier")

st.markdown(
    "Paste a purchase order, requisition, or invoice description below. "
    "The classifier will assign a **Business Category**, **NIGP code**, and "
    "**confidence score** — drawing on the same rule base used for the full "
    "historical extract."
)

with st.form("classify_form"):
    description = st.text_area(
        "Description to classify",
        placeholder="e.g.,  RENTAL OF HEAVY EQUIPMENT FOR AIRPORT TAXIWAY REPAIR",
        height=120,
    )
    classify = st.form_submit_button("Classify")

if classify:
    if not description.strip():
        st.warning("Please enter a description.")
        st.stop()

    rules = get_rules()
    result = classify_one(description, rules)
    d = result.as_dict()

    # Status badge logic
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

    st.markdown("### Result")
    st.markdown(
        f"""
        <div class='chi-result-box'>
          <span class='chi-badge' style='background:{badge_color};'>{badge_text}</span>
          <div class='chi-result-category'>{category_display}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("#### Classification detail")
    detail_rows = [
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
    st.table(pd.DataFrame(detail_rows, columns=["Field", "Value"]))

    st.caption(
        "If the description was sent to human review, that means the classifier "
        "did not find a confident rule match. The corresponding row in the bulk "
        "review queue (NIGP_Mapping_Review_Queue_JHK3.csv) is reviewed during "
        "the regular procurement-staff triage cadence."
    )

# -- About / help expander -------------------------------------------------
with st.expander("About this tool"):
    st.markdown(
        """
        - **What it does:** Classifies a single PO, requisition, or invoice description
          into one of 17 Chicago Business Categories plus an NIGP commodity code.
        - **What it uses:** Description text only. Vendor name and prior consultant-supplied
          NIGP codes are not inputs (by design).
        - **How it decides:** Hand-curated keyword rules first, then AI-mined long-tail
          rules. No AI is consulted at runtime — the classifier is rules-only and
          fully auditable.
        - **Where the rules live:** `spend-analysis/data/reference/keyword_rules_DRAFT_JHK3.csv`
          and `keyword_rules_from_ai_JHK3.csv`, both editable by procurement staff.
        - **Bulk classification:** This tool handles one description at a time. To
          classify a whole file, run `python spend-analysis/scripts/classifier_JHK3.py --batch`.
        """
    )

# -- Footer ---------------------------------------------------------------
st.markdown(
    f"<div class='chi-footer'>"
    f"City of Chicago  •  Department of Procurement Services  •  "
    f"NIGP Procurement Taxonomy & Classification Engine v1.1  •  "
    f"Prepared by James H. Kirby III, CSCP, MS-SCM"
    f"</div>",
    unsafe_allow_html=True,
)
