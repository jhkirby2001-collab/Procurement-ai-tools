# How the NIGP-Sourced Procurement Category Mapper Works — and How to Operate It

**Prepared by:** James H. Kirby III, CSCP, MS-SCM
**Document version:** 2026-08-07 (expanded operate-and-trust edition)
**Project:** NIGP-Sourced Procurement Category Mapper
**Who this is for:** Leadership who need to understand and stand behind the tool's output, and procurement staff who operate it day to day.

---

## 0. Read this first — the one-paragraph summary

This tool takes a plain-English procurement description — "elevator maintenance," "bulk rock salt," "chain link fence repair" — and assigns it a standardized commodity classification: a business-friendly **Business Category**, an industry-standard **NIGP 3-digit Class**, and, where the description is specific enough, a **5-digit NIGP Item**. It was validated on **784,556 historical purchase-order and invoice-line records** and maps **100% of them**. Critically, it is **rules-driven and fully auditable**: every answer comes with the exact rule that produced it, a confidence level, and a plain-language reason. **No AI model is called when you use it** — the AI was used exactly once, at build time, to help write rules, and then frozen. That is what makes the output defensible in an audit.

---

## 1. What problem this solves

Most public agencies never build their own classification of *what they buy*. Without one, the basic strategic questions can't be answered with confidence:

- Where does spend actually go, by commodity?
- What are we buying across multiple departments without coordinating it?
- Can any single purchase be traced, categorized, and **defended** to an auditor?

This mapper is that missing classification layer. It is **independent** — it does not rely on any prior consultant's labels — and it is **reusable**: it re-points at any future spend extract or contract portfolio without re-running AI or hiring anyone.

---

## 2. The taxonomy — three levels on every row

Every classified record carries **three** levels of classification. This serves two audiences at once: leadership reads the top level; auditors and sourcing analysts use the NIGP codes underneath.

```
Level 1  →  Business Category      (17 buckets — the plain-English answer)
Level 2  →  NIGP 3-digit Class     (industry-standard commodity class)
Level 3  →  NIGP 5-digit Item      (specific item, when the text supports it)
```

### The 17 Business Categories

| # | Business Category | Typical contents |
|---|---|---|
| 1 | Facilities Operations & Maintenance | HVAC, plumbing, lighting, doors, building hardware, electrical supplies |
| 2 | Public Safety, Uniforms & PPE | Uniforms, badges, fire equipment, gloves, masks, body armor, ammunition |
| 3 | Construction Materials | Concrete, asphalt, steel, pipe, rebar, pavement marking |
| 4 | Vehicles & Fleet | Vehicles, fleet parts, repairs, tires, fuel |
| 5 | Professional & Administrative Services | Management services, security guards, translation, legal/records services |
| 6 | Office, Print & Marketing | Office supplies, printing, advertising, signage |
| 7 | Equipment Rental & Leasing | Heavy-equipment rentals, copier leases, traffic-control rentals |
| 8 | Janitorial, Sanitation & Waste | Cleaning supplies, paper products, carpet/floor covering, dumpsters, disposal |
| 9 | Heavy Equipment & Machinery | Loaders, sweepers, lifts, forklifts, sprayers (purchased, not rented) |
| 10 | IT, Telecom & Audio/Visual | Network services, telephony, A/V, printers, toner |
| 11 | Landscaping, Grounds & Irrigation | Landscape services, irrigation, pest management, fencing, window washing, snow removal |
| 12 | Chemicals & Water Treatment | Industrial chemicals, water treatment, deicers / road salt |
| 13 | Medical & Health Services | Medical gases, defibrillators/AEDs, exam services, behavioral health |
| 14 | Animal Care & Veterinary | Veterinary supplies, animal feed, beekeeping |
| 15 | Furniture & Furnishings | Office furniture, mattresses, fitness equipment |
| 16 | Construction & Trades Services | Bridge construction, expansion-joint repair, contracted tradesmen |
| 17 | Grants & Pass-Through Funding | Subgrant disbursements to community-based organizations (not commodity purchases) |

**Why category 17 is separate:** Subgrants are financial transfers to subrecipients, not commodity buys. They have different governance and audit treatment, so mixing them into "professional services" would obscure real structure. This category has no NIGP class because NIGP is a commodity framework and has no "subgrants" code.

The full mapping of which NIGP 3-digit Classes roll up to each Business Category lives in `business_categories_JHK3.csv` (138 rows, each with a judgment note).

---

## 3. How the mapper was built — the part that earns your confidence

This section explains, in order, exactly how the tool was constructed. Understanding this is what lets you defend the numbers.

### 3.1 The source data

The taxonomy and rules were derived from a **784,556-row public-sector procurement extract** (87 columns, AP activity years **2017, 2020, 2021, 2023**). Every purchase-order and invoice line was treated as raw transactional data to be classified independently — **not** re-labeled from anyone's prior categorization.

**A deliberate choice you should know about:** the source file already carried a partial set of prior commodity codes on about 30% of rows. **The mapper does not use those as inputs.** Building an independent, owned classification means classifying from the description text itself, not inheriting a prior work product. The prior codes are preserved only for optional after-the-fact cross-checking ("did our independent answer agree?").

### 3.2 What the classifier is allowed to look at — and what it is not

The classifier accepts only two kinds of signal:

1. **Description text** — it reads up to four description fields per row and picks the first substantive one (`Description_Best`).
2. **Account / object / fund codes** — the agency's own financial codes, used as a backup signal when the description is too thin to classify.

**Excluded on purpose:**

- **Vendor name.** The same vendor often sells across many categories on one contract, so "who sold it" is an unreliable guide to "what it is." Including vendor would *introduce* misclassification, not reduce it.
- **Spend dollars.** The classifier answers *what was bought*, not *how much was spent*. Dollars are a downstream analysis that consumes this taxonomy — never an input to it.
- **Prior consultant codes.** See 3.1.

### 3.3 The rule base — how classifications are actually decided

Classification runs on a library of **rules stored as plain CSV files** that procurement staff can read and edit without touching any code. As of this document there are:

| Rule source | Count | What it is |
|---|---:|---|
| Hand-curated keyword rules | **280** | Written and owned by the project author. The high-volume, highest-trust rules. |
| AI-mined keyword rules | **6,766** | Harvested from a one-time AI pattern-mining pass (see 3.4). Each carries full provenance. Frozen. |
| Account-code patterns | **6** | Map specific financial account codes to a category (mainly subgrant disbursement accounts). |
| **Total rules** | **7,046** | Plus one resolver that consumes saved AI output (3.4). |

Each keyword rule says: *"if the description matches this pattern, assign this Business Category and NIGP class."* Matching supports three modes — `exact` (whole string), `starts_with` (prefix), and `contains` (substring, the most common) — evaluated in that priority order.

### 3.4 The one and only use of AI — bounded, one-time, and frozen

This is the single most important point for defensibility, so it is stated plainly:

> **The production tool is rules-only. No AI model is called when you classify anything.**

AI was used **once**, at build time, for one narrow job: to read the **long tail** of ~30,000 distinct descriptions that no hand-written rule covered, and propose patterns. Specifics:

- **Model:** a single batch run of Anthropic Claude (Haiku tier). Total one-time cost ≈ **$27**.
- **Constrained output:** the model was boxed into the **17 approved categories and 138 NIGP classes** by a strict schema. It **could not invent a code** outside that list.
- **Blind to context:** each AI call saw only the description text — never vendor, dollars, account codes, or prior labels.
- **Promotion to rules:** high-confidence proposals were promoted into the rule file automatically; medium-confidence only if the phrase recurred at least 5 times; low-confidence never. That yielded the **6,766 frozen AI-mined rules**.
- **The saved output is a file, forever.** The full mining output is committed as a CSV. A later "resolver" step (see 4, Tier 3) reads that saved file to fill any remaining rows — **still no new API call; it's a local file read.**

**Why this is defensible, in six points:**

1. **One-time, not ongoing.** Runs never call AI. No API key, no internet, no per-classification charge.
2. **Bounded.** The AI could only choose from the approved taxonomy — no invented codes.
3. **Transparent.** Every AI-mined rule records the model, confidence, source frequency, and the model's reasoning.
4. **Reviewable & editable.** Staff can demote or delete any AI-mined rule; it has no special status.
5. **Replaceable.** Delete all AI content and the tool still runs on hand-curated rules + account patterns. Coverage drops; nothing breaks.
6. **Confidence-labeled, never laundered.** AI-derived answers are tagged `AI-high / AI-medium / AI-low` in the output so you can filter or exclude them at will. They are not hidden inside a generic "auto-classified" bucket.

---

## 4. How a row gets classified — the four tiers

Every record runs through four tiers **in order**. The first tier that matches wins, and the row records which tier produced its answer (the `Classification_Method` field). This is what produces a clean, per-row audit trail.

| Tier | Method label | What it means |
|---|---|---|
| 1 | `keyword_rule` | The description matched a curated or AI-mined keyword rule. **Strongest evidence.** |
| 2 | `account_pattern` | The description was too thin; a financial account code resolved it (e.g. subgrant accounts). |
| 3 | `ai_assist` | No rule matched; the saved one-time AI output supplied the answer. Tagged `AI-high/medium/low`. **No new API call.** |
| 4 | `no_description` | No usable text in any description field. Honestly tagged `Unclassified — No Description` rather than guessed. |

### Coverage from the last full production run

The table below is from the **last complete batch run (2026-05-14)**, which classified all 784,556 rows using the rule set as it stood on that date:

| Tier | Rows | Share |
|---|---:|---:|
| Tier 1 — Keyword rule | 688,044 | 87.7% |
| Tier 2 — Account-code pattern | 2,028 | 0.3% |
| Tier 3 — AI-assist resolver (saved output) | 93,518 | 11.9% |
| Tier 4 — Unclassified — No Description | 966 | 0.1% |
| **Total mapped to a Business Category** | **784,556** | **100.0%** |
| Terminal human-review queue | **0** | **0.0%** |

> **Note on freshness:** 34 curated rules were added on 2026-08-07 (see §7). Those improve the interactive tool immediately, but the batch percentages above will only change on the **next full batch refresh**. The headline "100% mapped" does not change — the new rules mostly shift a small number of rows from Tier 3 (AI-assist) to Tier 1 (keyword rule), which strengthens the evidence behind those rows without moving the top-line coverage.

**On "0 review queue" vs. "0 error":** these are not the same claim. Tier 3 rows carry an `AI-medium` or `AI-low` tag and rest on an AI proposal, not a deterministic rule. The confidence field exposes that on every row, so you can always filter to the strongest-evidence tiers when accuracy matters most. Choosing to *fill and label* rather than *queue and ignore* 140,000 rows was a deliberate decision — a labeled 100% beats a review pile no one will ever triage.

---

## 5. How to operate it — the web app (primary) and the CLI (advanced)

There are two ways to use the tool. **Most people should use the web app.**

### 5A. The web app — eight pages

The tool is deployed as a password-gated web app (the address and password are maintained by the project author — request access from James). It has eight pages:

| Page | What it does | Who uses it |
|---|---|---|
| **Classify** | Paste one description → get Business Category, NIGP class, confidence, and the rule that fired. | Anyone classifying a single item |
| **Bulk Classify** | Upload a spreadsheet of descriptions → download it back with a category on every row. | Staff processing a list |
| **Spend Report** | Upload a spend file → it classifies, then produces a spend analysis (spend-by-category, Pareto 80/20, top vendors, consolidation) + downloadable Excel. | Anyone analyzing a spend file |
| **Spend Report Methodology** | How the spend analysis works, how to run it, how to read each report, and its limits. | Leadership, staff learning the tool |
| **Procurement Taxonomy Logic** | Visual walkthrough of the three levels and the four tiers. | Anyone learning the model |
| **Methodology** | The full build narrative and defensibility argument. | Leadership, auditors |
| **Business Categories** | The 17 categories and what rolls into each. | Anyone checking scope |
| **Rule Lookup** | Search all 7,046 rules to see exactly how any phrase routes and why. | Anyone validating a decision |

**Two everyday workflows on the web app:**

- *"What category is this?"* → **Classify** page → paste the description → read the answer and the reason.
- *"Why did it decide that?"* → **Rule Lookup** page → search the keyword → see the exact rule, its NIGP class, and its note.

### 5B. Look up how something was already classified (no tools needed)

If you just want to see how similar items were categorized historically, open `NIGP_Mapping_JHK3.csv` in Excel (it's large — give it a minute), turn on **Data → Filter**, and filter the `Description_Best` column for your keyword. The `Business_Category` column shows how comparable purchases were classified across the full history.

### 5C. The command line (advanced / technical users)

For a single description from a terminal in the project environment:

```bash
python spend-analysis/scripts/classifier_JHK3.py --describe "PASTE THE DESCRIPTION HERE"
```

It prints the Business Category, NIGP code, confidence, and the rule that fired. If you don't have access to the environment, send the description to James and he'll run it.

### 5D. The Spend Report — spend analysis from any file

The **Spend Report** page takes the tool one step past classification: upload a spend file **in any format** and it classifies every line, then builds a full spend analysis you can download as one Excel workbook.

**It's adaptive.** The tool inspects *every* column (name and content), figures out which is the description, amount, vendor, department, date, and any breakdown dimensions, then **runs every analysis the data supports and skips the ones it can't** — telling you why. You don't have to shape your file to fit the tool.

1. Open the **Spend Report** page → **upload** a CSV or Excel spend file.
2. **Confirm the detected columns** — the tool pre-fills them; only change a box if a guess is wrong (only Description is required).
3. Click **Generate spend report** → read the results and **download the Excel workbook** (a tab per analysis, with charts, in the Chicago palette).

Depending on what your file contains, it can produce: an **executive summary** (findings + recommended steps), **spend by category**, **Pareto 80/20**, **top vendors**, **vendor concentration/risk (HHI)**, **tail-spend**, **single- vs multi-source**, **spend trend over time** (needs a date), **spend by department**, a **category × department matrix**, **vendor consolidation/fragmentation**, and **spend by any dimension** it finds — contract vs non-contract, supplier diversity, status, and more. All deterministic arithmetic — **no AI, and the file stays in your session.** Full detail is on the in-app **Spend Report Methodology** page.

**Honest limit:** classification is keyword-rules-only, so coverage depends on description quality — the coverage note always shows exactly how much of the file was classified. The tool never invents data it doesn't have; a missing column just means that analysis is skipped.

---

## 6. How to fix or improve a classification — permanently

This is the maintenance workflow, and it's deliberately simple: **teach it once, and every future similar description is right.** This is exactly how the "fence" gap was closed on 2026-08-07 (see §7) — a worked example follows.

1. Open `spend-analysis/data/reference/keyword_rules_DRAFT_JHK3.csv` in Excel.
2. Add a row with these columns:
   - `pattern` — the text to match, e.g. `ROAD SALT`
   - `match_type` — `contains` (most common), `starts_with`, or `exact`
   - `business_category` — one of the 17, spelled exactly
   - `nigp_class_3digit` — the correct NIGP class (verify it against `nigp_codes_5digit_JHK3.csv` — do **not** guess)
   - `nigp_item_5digit` — leave blank unless the text is specific enough
   - `nigp_match_level` — `broad` for normal cases, `exact` for very specific full-string rules
   - `notes` — a short justification (date, initials, why)
3. Save. The new rule takes effect immediately in the web app and CLI, and folds into the numbers at the next batch run.

> **Worked example — the "fence" fix (2026-08-07).** A user typed "fence" and got no result. Investigation showed there was no rule for the bare word, and the one related rule pointed at the wrong category. Checking the actual data settled it: every fence row was a *landscape salt-fence service*, and the NIGP item for chain-link fence/guardrail repair sits under class **988 → Landscaping, Grounds & Irrigation**. Three rules were added/aligned to that verified home. "fence," "fencing," and "chain link fence repair" now all classify correctly. **That is the whole maintenance loop: a gap is found, the correct NIGP home is verified against the reference table, and a one-line rule closes it for good.**

---

## 7. Recent improvement — the interactive robustness pass (2026-08-07)

The "fence" report surfaced a broader pattern: the interactive tool only answers when your typed words match a rule, and many **short, common** words had no rule (even though the full 784K batch was already 100% mapped, because real descriptions are long and detailed). A 45-word test of everyday procurement terms found only **17 of 45 (38%)** returned a result.

**What was done:** 34 curated rules were added, each mapped to its **correct** NIGP class verified against the reference table — deliberately *not* auto-matched, which would have produced confident-but-wrong answers. Coverage on those common terms rose to **35 of 45 (78%)**, and the plain-English regression test moved to a clean **72/72, zero mis-categorizations.** Runtime stayed rules-only, no API key. Newly covered terms include gloves, uniforms, ammunition, PPE/masks, paper towels, carpet, road salt/deicer, translation, security guards, defibrillator, batteries, electrical wire, forklift, toner, window cleaning, snow removal, fuel, and pavement marking.

**Honest limit:** a handful of terms (e.g. hand sanitizer, traffic signals, surveying, architecture services, scaffolding) still have no home, because the working NIGP reference only contains the ~470 codes that actually appeared in the source data. Covering those would require importing the full official NIGP dictionary — a separate, deliberate decision.

---

## 8. Reading a result — what every field means

| Field | What it tells you |
|---|---|
| `Business_Category` | **The main answer** — one of the 17 buckets. |
| `NIGP_Class_3digit` | The industry-standard 3-digit commodity class. |
| `NIGP_Item_5digit` | The specific 5-digit item, when the text supports it (else blank). |
| `NIGP_Code_Assigned` | The code landed on (5-digit if available, else 3-digit). |
| `NIGP_Match_Level` | `exact` = specific item assignable • `broad` = 3-digit class only • `review` = flagged. |
| `Classification_Confidence` | High / Medium / Low (or `AI-high/medium/low` for Tier 3). |
| `Classification_Method` | Which of the four tiers produced the answer. |
| `Classification_Reason` | **The exact rule that fired — the audit trail.** |

**How to read confidence in practice:** a `keyword_rule` / High row is deterministic and strongly evidenced. An `ai_assist` / `AI-low` row is a reasonable inference that deserves a glance if the dollars or audit exposure are significant. The point of the design is that you can always *see* which is which.

---

## 9. Limitations — stated plainly

1. **Single-source origin.** The taxonomy was derived from one dataset; new extracts may reveal commodity types not represented yet.
2. **Description-quality dependent.** Thin descriptions ("Misc," "Per contract") are resolved at Tier 3 with an explicit low-confidence tag; truly empty ones are honestly tagged `Unclassified — No Description` rather than guessed.
3. **Partial NIGP working set.** The ~138 classes / ~470 items in use are the subset that appeared in the source data — a fraction of the full ~9,000-code NIGP catalog. Some novel terms have no home until the catalog is expanded (see §7).
4. **No vendor or dollar signal, by design.** The answer reflects *what was bought*, not who sold it or how much it cost.

---

## 10. Keeping it trustworthy over time — governance

1. **Rule-file ownership.** One named owner for `keyword_rules_DRAFT_JHK3.csv`. Add rules as new commodities appear; retire stale ones. The curated base has grown 148 → 246 → **280** on this principle.
2. **AI-tier audit cadence.** Quarterly, sample `ai_assist` rows tagged `AI-medium/AI-low`. Where the AI got a recurring pattern right, promote it into a curated rule so the next run no longer depends on AI-assist for it.
3. **Annual taxonomy review.** Confirm the 17 categories still fit reporting needs; edit `business_categories_JHK3.csv` directly.
4. **Blind audit benchmark.** Once a year, have an analyst independently classify 100 random rows blind and compare. Track the agreement rate by tier over time.

---

## 11. Files that travel with this deliverable

| File | What it is |
|---|---|
| `NIGP_Mapping_JHK3.csv` | The full classified dataset. 100% mapped. (~225 MB; not stored in version control by size.) |
| `HOW_TO_USE_NIGP_Mapping_JHK3.md` | This guide. |
| `METHODOLOGY_JHK3.md` (in `spend-analysis/`) | The full technical methodology and locked-decision record. |
| `business_categories_JHK3.csv` | The 138-row NIGP-class → Business-Category map with judgment notes. |
| `keyword_rules_DRAFT_JHK3.csv` | The 280 hand-curated rules — editable by procurement staff. |
| `keyword_rules_from_ai_JHK3.csv` | The 6,766 AI-mined rules, each with full provenance. Frozen. |

---

## Questions

Direct to **James H. Kirby III, CSCP, MS-SCM**, project author. The full methodology and audit trail are in `METHODOLOGY_JHK3.md`, and every classification decision can be traced to the exact rule that produced it.
