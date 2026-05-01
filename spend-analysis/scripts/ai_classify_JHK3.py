"""
ONE-TIME-USE AI classifier for the no-rule-fire long tail.

Per locked decision (option B, 2026-04-29): AI is invoked ONCE during taxonomy
build to mine patterns from descriptions that no keyword rule matched. Output
is harvested into keyword_rules.csv. The PRODUCTION classifier is rules-only;
this script never runs during normal operation.

Strategy:
- Filter classified_full_JHK3.csv to human_review rows (no rule fired)
- Dedupe to unique descriptions (~30K, down from ~172K rows)
- Call Haiku 4.5 once per unique description with prompt caching of taxonomy
- Output is constrained by JSON schema with enum on business_category and
  nigp_class_3digit — model CANNOT invent codes outside our 17 categories or
  138 classes
- Resumable: partial results checkpointed every 50 calls

Setup required to run:
1. Create an Anthropic API account at https://console.anthropic.com
2. Generate an API key, set spending cap (recommend $50)
3. export ANTHROPIC_API_KEY="sk-ant-..."
4. pip install anthropic
5. python scripts/ai_classify_JHK3.py

Estimated cost: ~$25-30 with Haiku 4.5 + prompt caching at 30K unique descriptions.
"""

from __future__ import annotations

import json
import os
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
SRC = ROOT / "data" / "processed" / "classified_full_JHK3.csv"
REF_3D = ROOT / "data" / "reference" / "nigp_codes_3digit_JHK3.csv"
OUT = ROOT / "data" / "processed" / "ai_classified_unique_descriptions_JHK3.csv"

MODEL = "claude-haiku-4-5"
CHECKPOINT_EVERY = 50
MAX_DESC_CHARS = 600  # truncate ultra-long descriptions
MAX_WORKERS = 20  # parallel API calls — well under Tier 1's 1,000 RPM

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


def build_system_prompt() -> str:
    """Stable taxonomy reference — cached across all API calls."""
    ref3 = pd.read_csv(REF_3D, dtype={"nigp_class_3digit": str})
    nigp_lines = []
    for _, r in ref3.iterrows():
        desc = (r["canonical_description"] or "")[:80]
        nigp_lines.append(f"  {r['nigp_class_3digit']}: {desc}")
    nigp_block = "\n".join(nigp_lines)

    cats_block = "\n".join(f"  - {c}" for c in BUSINESS_CATEGORIES)

    return f"""You are a public-sector procurement classifier for the City of Chicago.

Your job: given a single PO/invoice description, assign it to ONE Business Category and ONE NIGP 3-digit Class. Use the controlled vocabularies below. Do NOT invent codes outside these lists.

The 17 BUSINESS CATEGORIES (you must pick exactly one):
{cats_block}

Special handling: "Grants & Pass-Through Funding" is for subgrant disbursements only (no NIGP class — leave nigp_class_3digit empty for this category).

The 138 NIGP 3-DIGIT CLASSES (with EY-data canonical descriptions) — you must pick exactly one or leave empty:
{nigp_block}

Rules:
1. Pick the BEST FIT business category. If the description is too generic to confidently determine, pick the most likely one and set confidence="low".
2. Pick the NIGP 3-digit class only if you can defensibly map the commodity to one. Otherwise leave nigp_class_3digit empty.
3. Confidence: "high" = clearly matches a category and class; "medium" = category clear, class uncertain; "low" = ambiguous.
4. Reason: ONE short sentence explaining your reasoning. No filler.
5. NEVER invent codes outside the 138 listed. NEVER invent business categories outside the 17 listed."""


def get_response_schema() -> dict:
    """JSON schema with enums — constrains the model to valid values."""
    return {
        "type": "object",
        "properties": {
            "business_category": {"type": "string", "enum": BUSINESS_CATEGORIES},
            "nigp_class_3digit": {
                "type": "string",
                "description": "3-digit NIGP class code (e.g. '485') or empty string if no clean mapping",
            },
            "confidence": {"type": "string", "enum": ["high", "medium", "low"]},
            "reason": {"type": "string"},
        },
        "required": ["business_category", "nigp_class_3digit", "confidence", "reason"],
        "additionalProperties": False,
    }


def load_unique_descriptions() -> pd.DataFrame:
    df = pd.read_csv(SRC, low_memory=False)
    no_rule = df[df["Classification_Method"] == "human_review"].copy()
    unique = (
        no_rule.groupby("Description_Best")
        .size()
        .reset_index(name="row_count")
        .rename(columns={"Description_Best": "description"})
    )
    unique = unique[unique["description"].fillna("").str.strip() != ""]
    unique = unique.sort_values("row_count", ascending=False).reset_index(drop=True)
    return unique


def load_existing_progress() -> set[str]:
    if not OUT.exists():
        return set()
    done = pd.read_csv(OUT)
    return set(done["description"].astype(str))


def append_to_output(rows: list[dict]) -> None:
    df_new = pd.DataFrame(rows)
    if OUT.exists():
        df_new.to_csv(OUT, mode="a", header=False, index=False)
    else:
        df_new.to_csv(OUT, index=False)


def classify_one(client: anthropic.Anthropic, system_prompt: str, description: str) -> dict:
    desc_clean = (description or "").strip()[:MAX_DESC_CHARS]
    response = client.messages.create(
        model=MODEL,
        max_tokens=512,
        system=[
            {
                "type": "text",
                "text": system_prompt,
                "cache_control": {"type": "ephemeral"},
            }
        ],
        output_config={
            "format": {"type": "json_schema", "schema": get_response_schema()}
        },
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


def main() -> None:
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("ERROR: ANTHROPIC_API_KEY env var not set.", file=sys.stderr)
        print("       export ANTHROPIC_API_KEY='sk-ant-...'", file=sys.stderr)
        sys.exit(1)

    print("Loading taxonomy reference and unique descriptions...")
    system_prompt = build_system_prompt()
    descriptions = load_unique_descriptions()
    done = load_existing_progress()
    todo = descriptions[~descriptions["description"].isin(done)].reset_index(drop=True)

    print(f"Unique descriptions total:    {len(descriptions):,}")
    print(f"Already classified (resume):  {len(done):,}")
    print(f"To classify in this run:      {len(todo):,}")
    if len(todo) == 0:
        print("Nothing to do.")
        return

    client = anthropic.Anthropic()
    rows_buffer: list[dict] = []
    buffer_lock = Lock()
    counters = {"input": 0, "output": 0, "cache_read": 0, "cache_create": 0, "done": 0, "errors": 0}
    counters_lock = Lock()
    start = time.time()

    def task(desc: str, row_count: int) -> Optional[dict]:
        try:
            result = classify_one(client, system_prompt, desc)
            return {
                "description": desc,
                "row_count": int(row_count),
                **result,
            }
        except anthropic.APIError as e:
            print(f"\nAPI error on '{desc[:60]}...': {e}", file=sys.stderr)
            return None

    print(f"Submitting {len(todo):,} tasks to {MAX_WORKERS}-worker pool...")
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {
            executor.submit(task, str(r["description"]), int(r["row_count"])): i
            for i, r in todo.iterrows()
        }

        for future in as_completed(futures):
            row = future.result()
            if row is None:
                with counters_lock:
                    counters["errors"] += 1
                continue

            with buffer_lock:
                rows_buffer.append(row)
                buf_len = len(rows_buffer)
            with counters_lock:
                counters["input"] += row["input_tokens"]
                counters["output"] += row["output_tokens"]
                counters["cache_read"] += row["cache_read_tokens"]
                counters["cache_create"] += row["cache_create_tokens"]
                counters["done"] += 1
                done_so_far = counters["done"]

            if buf_len >= CHECKPOINT_EVERY:
                with buffer_lock:
                    flush, rows_buffer = rows_buffer, []
                if flush:
                    append_to_output(flush)
                elapsed = time.time() - start
                rate = done_so_far / elapsed
                eta_min = (len(todo) - done_so_far) / rate / 60 if rate > 0 else 0
                with counters_lock:
                    cost_so_far = (
                        counters["input"] * 1.0 / 1_000_000
                        + counters["output"] * 5.0 / 1_000_000
                        + counters["cache_read"] * 0.1 / 1_000_000
                        + counters["cache_create"] * 1.25 / 1_000_000
                    )
                    cache_read = counters["cache_read"]
                    n_errors = counters["errors"]
                print(
                    f"  [{done_so_far:>6,}/{len(todo):,}]  "
                    f"rate={rate:5.1f}/s  ETA={eta_min:5.1f} min  "
                    f"cost=${cost_so_far:6.2f}  cache_read={cache_read:>11,}  "
                    f"errors={n_errors}"
                )

    # Final flush
    with buffer_lock:
        if rows_buffer:
            append_to_output(rows_buffer)

    elapsed = time.time() - start
    cost = (
        total_input * 1.0 / 1_000_000
        + total_output * 5.0 / 1_000_000
        + total_cache_read * 0.1 / 1_000_000
        + total_cache_create * 1.25 / 1_000_000
    )
    print()
    print(f"Done in {elapsed/60:.1f} minutes")
    print(f"Estimated cost: ${cost:.2f}")
    print(f"Output: {OUT}")


if __name__ == "__main__":
    main()
