# spend-analysis — sub-project guide

Ad-hoc procurement spend analytics, separate from the NIGP classifier documented in the
repo-root `CLAUDE.md`. Keep NIGP content out of this file and spend-analysis content out
of that one.

## Active work

| Project | Doc | Status |
|---|---|---|
| **Maverick Spend Conversion Tracker** — Exhibit B → awarded contract; metric is **Contract Conversion Rate (CCR)** | `EXHIBIT_B_CONVERSION_JHK3.md` | Complete, delivered 2026-08-21, safe to resume |

## Naming

Projects here get a vendor-neutral, industry-standard name because the repo is public;
internal jargon ("Exhibit B") belongs in subtitles and in the body where it labels source
data, not in project names. Name the *metric* as well as the project — a named metric
(CCR) gets tracked quarter over quarter, an unnamed analysis gets forgotten.

## Conventions

- `_JHK3` suffix on every deliverable, script and rule file.
- Deliverables land in repo-root `outputs/`.
- `data/raw/` and `data/processed/` are **gitignored** — organizational source extracts are
  never committed, and this repo is public. Scripts must therefore locate their source at
  runtime (env var, conventional path glob, or argv) and fail with a clear message rather
  than assuming a hardcoded path. Several older scripts here still carry absolute
  `/workspaces/...` paths from a previous codespace and will not run as-is.
- Excel deliverables: summary figures should be live formulas over the detail tabs, not
  values computed in Python and pasted in. Verify computed cells against the analysis code
  before shipping — a clean recalculation proves formulas evaluate, not that they are right.
- `libreoffice-calc` is absent from fresh containers; install it before recalculating
  formula workbooks or the check silently times out.
