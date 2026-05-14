"""
Targeted AI top-up: classify the residual unique descriptions that the
review-queue resolver couldn't fill from the existing AI output.

Input: data/processed/review_queue_uncovered_uniques_JHK3.csv (written by
resolve_review_queue_JHK3.py) — a single-column CSV of normalized descriptions.

Action: call Haiku 4.5 once per description, with the same taxonomy-cached
prompt as the original ai_classify_JHK3.py. Results are APPENDED to
ai_classified_unique_descriptions_JHK3.csv so the resolver can re-merge.

Idempotent on the AI output side: descriptions already present in the AI
output file are skipped automatically. Safe to re-run.

Cost target: ~$0.50 for a few hundred descriptions with prompt caching.

Run from repo root:
  python spend-analysis/scripts/ai_topup_uncovered_JHK3.py
"""
from __future__ import annotations

import json
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from threading import Lock
from typing import Optional

import pandas as pd

try:
    import anthropic
except ImportError:
    print("ERROR: anthropic SDK not installed. Run: pip install anthropic", file=sys.stderr)
    sys.exit(1)

ROOT = Path(__file__).resolve().parents[1]
UNCOVERED_PATH = ROOT / "data" / "processed" / "review_queue_uncovered_uniques_JHK3.csv"
AI_OUT_PATH = ROOT / "data" / "processed" / "ai_classified_unique_descriptions_JHK3.csv"
REF_3D = ROOT / "data" / "reference" / "nigp_codes_3digit_JHK3.csv"
MAPPING_PATH = ROOT.parent / "outputs" / "NIGP_Mapping_JHK3.csv"

MODEL = "claude-haiku-4-5"
MAX_DESC_CHARS = 600
MAX_WORKERS = 20
WHITESPACE_RE = re.compile(r"\s+")

BUSINESS_CATEGORIES = [
    "Vehicles & Fleet",
    "Heavy Equipment & Machinery",
    "Equipment Rental & Leasing",
    "Construction Materials",
    "Construction & Trades Services",
    "Facilities Operations & Maintenance",
    "Landscaping, Grounds & Irrigation",
    "Janitorial, Sanitation & Waste",
    "Chemicals & Water Treatment",
    "Public Safety, Uniforms & PPE",
    "Medical & Health Services",
    "Animal Care & Veterinary",
    "IT, Telecom & Audio/Visual",
    "Office, Print & Marketing",
    "Furniture & Furnishings",
    "Professional & Administrative Services",
    "Grants & Pass-Through Funding",
]


def normalize(s: str) -> str:
    if s is None:
        return ""
    return WHITESPACE_RE.sub(" ", str(s).strip()).upper()


def build_system_prompt() -> str:
    ref3 = pd.read_csv(REF_3D, dtype={"nigp_class_3digit": str})
    nigp_lines = []
    for _, r in ref3.iterrows():
        desc = (r["canonical_description"] or "")[:80]
        nigp_lines.append(f"  {r['nigp_class_3digit']}: {desc}")
    nigp_block = "\n".join(nigp_lines)
    cats_block = "\n".join(f"  - {c}" for c in BUSINESS_CATEGORIES)
    return f"""You are a public-sector procurement classifier.

Your job: given a single PO/invoice description, assign it to ONE Business Category and ONE NIGP 3-digit Class. Use the controlled vocabularies below. Do NOT invent codes outside these lists.

The 17 BUSINESS CATEGORIES (you must pick exactly one):
{cats_block}

Special handling: "Grants & Pass-Through Funding" is for subgrant disbursements only (no NIGP class — leave nigp_class_3digit empty for this category).

The 138 NIGP 3-DIGIT CLASSES (with canonical descriptions) — you must pick exactly one or leave empty:
{nigp_block}

Rules:
1. Pick the BEST FIT business category. If too generic, pick the most likely one and set confidence="low".
2. Pick the NIGP 3-digit class only if you can defensibly map. Otherwise leave nigp_class_3digit empty.
3. Confidence: "high" = clearly matches a category and class; "medium" = category clear, class uncertain; "low" = ambiguous.
4. Reason: ONE short sentence. No filler.
5. NEVER invent codes outside the 138 listed. NEVER invent business categories outside the 17 listed."""


def get_response_schema() -> dict:
    return {
        "type": "object",
        "properties": {
            "business_category": {"type": "string", "enum": BUSINESS_CATEGORIES},
            "nigp_class_3digit": {"type": "string"},
            "confidence": {"type": "string", "enum": ["high", "medium", "low"]},
            "reason": {"type": "string"},
        },
        "required": ["business_category", "nigp_class_3digit", "confidence", "reason"],
        "additionalProperties": False,
    }


def classify_one(client, system_prompt: str, description: str) -> dict:
    desc_clean = (description or "").strip()[:MAX_DESC_CHARS]
    response = client.messages.create(
        model=MODEL,
        max_tokens=512,
        system=[{"type": "text", "text": system_prompt, "cache_control": {"type": "ephemeral"}}],
        output_config={"format": {"type": "json_schema", "schema": get_response_schema()}},
        messages=[{"role": "user", "content": f"Description: {desc_clean}"}],
    )
    text = next((b.text for b in response.content if b.type == "text"), "")
    parsed = json.loads(text)
    return {
        "business_category": parsed.get("business_category", ""),
        "nigp_class_3digit": parsed.get("nigp_class_3digit", ""),
        "confidence": parsed.get("confidence", ""),
        "reason": parsed.get("reason", ""),
        "input_tokens": response.usage.input_tokens,
        "output_tokens": response.usage.output_tokens,
        "cache_read_tokens": getattr(response.usage, "cache_read_input_tokens", 0),
        "cache_create_tokens": getattr(response.usage, "cache_creation_input_tokens", 0),
    }


def load_existing_ai_keys() -> set[str]:
    if not AI_OUT_PATH.exists():
        return set()
    df = pd.read_csv(AI_OUT_PATH)
    return set(df["description"].fillna("").apply(normalize))


def load_uncovered_with_counts() -> pd.DataFrame:
    """Read uncovered list and join row counts from the mapping file."""
    uncovered = pd.read_csv(UNCOVERED_PATH)
    mapping = pd.read_csv(MAPPING_PATH, low_memory=False)
    mapping["_k"] = mapping["Description_Best"].fillna("").apply(normalize)
    counts = mapping.groupby("_k").size().reset_index(name="row_count")
    out = uncovered.merge(counts, left_on="description_norm", right_on="_k", how="left")
    out["row_count"] = out["row_count"].fillna(1).astype(int)
    return out[["description_norm", "row_count"]].rename(columns={"description_norm": "description"})


def main() -> None:
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("ERROR: ANTHROPIC_API_KEY env var not set.", file=sys.stderr)
        print("       export ANTHROPIC_API_KEY='sk-ant-...'", file=sys.stderr)
        sys.exit(1)

    if not UNCOVERED_PATH.exists():
        print(f"ERROR: {UNCOVERED_PATH} not found. Run resolve_review_queue_JHK3.py first.", file=sys.stderr)
        sys.exit(1)

    print("Loading taxonomy reference...")
    system_prompt = build_system_prompt()

    print("Loading uncovered descriptions...")
    todo_df = load_uncovered_with_counts()
    existing = load_existing_ai_keys()
    todo_df = todo_df[~todo_df["description"].apply(normalize).isin(existing)].reset_index(drop=True)
    print(f"To classify: {len(todo_df):,}")
    if len(todo_df) == 0:
        print("Nothing to do.")
        return

    client = anthropic.Anthropic()
    results: list[dict] = []
    counters = {"input": 0, "output": 0, "cache_read": 0, "cache_create": 0, "done": 0, "errors": 0}
    lock = Lock()
    start = time.time()

    def task(desc: str, row_count: int):
        try:
            r = classify_one(client, system_prompt, desc)
            return {"description": desc, "row_count": int(row_count), **r}
        except Exception as e:
            print(f"\nAPI error on '{desc[:60]}...': {e}", file=sys.stderr)
            return None

    print(f"Submitting {len(todo_df):,} tasks to {MAX_WORKERS}-worker pool...")
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futures = {ex.submit(task, str(r["description"]), int(r["row_count"])): i for i, r in todo_df.iterrows()}
        for fut in as_completed(futures):
            row = fut.result()
            if row is None:
                with lock:
                    counters["errors"] += 1
                continue
            with lock:
                results.append(row)
                counters["input"] += row["input_tokens"]
                counters["output"] += row["output_tokens"]
                counters["cache_read"] += row["cache_read_tokens"]
                counters["cache_create"] += row["cache_create_tokens"]
                counters["done"] += 1
                if counters["done"] % 50 == 0:
                    elapsed = time.time() - start
                    rate = counters["done"] / elapsed if elapsed else 0
                    cost = (
                        counters["input"] * 1.0 / 1_000_000
                        + counters["output"] * 5.0 / 1_000_000
                        + counters["cache_read"] * 0.1 / 1_000_000
                        + counters["cache_create"] * 1.25 / 1_000_000
                    )
                    print(f"  [{counters['done']:>5,}/{len(todo_df):,}]  rate={rate:5.1f}/s  cost=${cost:.2f}  errors={counters['errors']}")

    # Append results
    if results:
        df_new = pd.DataFrame(results)
        if AI_OUT_PATH.exists():
            df_new.to_csv(AI_OUT_PATH, mode="a", header=False, index=False)
        else:
            df_new.to_csv(AI_OUT_PATH, index=False)
        print(f"\nAppended {len(results):,} rows → {AI_OUT_PATH}")

    elapsed = time.time() - start
    cost = (
        counters["input"] * 1.0 / 1_000_000
        + counters["output"] * 5.0 / 1_000_000
        + counters["cache_read"] * 0.1 / 1_000_000
        + counters["cache_create"] * 1.25 / 1_000_000
    )
    print(f"Done in {elapsed/60:.1f} min. Cost: ${cost:.2f}. Errors: {counters['errors']}.")


if __name__ == "__main__":
    main()
