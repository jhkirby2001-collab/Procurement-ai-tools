# NIGP-Sourced Procurement Category Mapper — Web App Factsheet

**Prepared by:** James H. Kirby III, CSCP, MS-SCM
**Document type:** Layer 1 factual brief (for Layer 3 strategic positioning)
**Date:** 2026-05-14

---

## What It Is

A password-gated web application that lets any authorized user classify procurement descriptions against the 17 Business Category / NIGP 3-digit / NIGP 5-digit taxonomy in real time. The same engine that classified 784,556 historical rows runs behind the web interface — single record, on demand, no IT ticket.

## Public URL

**`https://chicago-nigp-classifier.streamlit.app/`**

Alias URL also resolves: `https://nigp-sourced-category-mapper.streamlit.app/` (redirects to canonical).

## Access

- Password-gated. Access credentials are operational and shared separately from this document.
- No City SSO integration (independent tool posture).
- Reads as an independent public-sector procurement utility — not branded as a City-institutional product.

## Six Pages

| # | Page | What It Does | Typical User |
|---|---|---|---|
| 1 | **Classify** | Paste a single description, get full classification (Business Category + NIGP class + NIGP item + confidence + reason). Recent-history sidebar. | Procurement staff doing live PO classification |
| 2 | **Bulk Classify** | Paste many descriptions (one per line) OR upload a CSV. Downloads results as CSV. | Buyers processing a batch of requisitions |
| 3 | **Procurement Taxonomy Logic** | Visual explainer of the three-tier hierarchy (Executive View → Sourcing View → Audit View). | Leadership, auditors |
| 4 | **Methodology** | How the engine decides; confidence-badge logic; AI use defensibility. | Auditors, technical reviewers |
| 5 | **Business Categories** | Browseable list of all 17 Business Categories with definitions. | Any user learning the taxonomy |
| 6 | **Rule Lookup** | Search across all 7,012 rules (246 hand-curated + 6,766 AI-mined) to verify how any specific phrase routes. | Anyone validating a classification decision |

## How It Works (One Sentence)

The web app loads the same rule files used by the batch classifier; pasting a description runs the same Python classification function that processed all 784,556 historical rows — no API calls, no AI at runtime.

## Hosting & Cost

| Item | Detail |
|---|---|
| Hosting platform | Streamlit Community Cloud |
| Hosting cost | **$0 / month** (free tier) |
| Deploy mechanism | Auto-redeploys on git push to `main` |
| Backing repository | Public GitHub repo |
| Runtime AI cost | **$0** (no API calls at runtime — rules-only) |
| One-time build AI cost | $27 (Anthropic Haiku 4.5, 2026-04-30, frozen) |

## Security Posture

- Password gate at the entry point; no anonymous access.
- Application secrets (password, etc.) stored in gitignored `.streamlit/secrets.toml`; never committed to the repository.
- No personally identifiable information processed; classification operates on description text only.
- All source code public on GitHub — auditable end-to-end.

## Availability

- Streamlit Community Cloud apps hibernate after extended inactivity and wake on request (cold start ~30–60 seconds).
- Verified live and serving the production rules as of 2026-05-14.
- No formal SLA — the tool is an internal productivity aid, not a transactional system.

## Strategic Value Points (Layer-3 Input)

These are factual capability statements; Layer 3 turns them into strategic narrative.

- **No IT ticket required.** Any authorized user accesses on browser, phone, tablet, laptop.
- **Live demonstrable.** Leadership can see the classification engine work in real time — not just read about it.
- **Auditable in front of stakeholders.** Rule Lookup page lets anyone confirm "how does this phrase get classified" without anyone touching code.
- **Methodology transparency.** The Methodology page documents the confidence-badge logic and explicitly notes the system never invents codes outside the controlled vocabulary.
- **Zero recurring vendor lock-in.** No SaaS subscription. No vendor invoice. No license renewal.
- **Independent posture.** Reads as an independent tool, not a City-institutional product — protects optionality for future deployment.
- **Same engine as the batch.** Single-record results match the historical 784,556-row file row-for-row; consistency is built in, not asserted.

## Defensible Numbers (For Anticipated Questions)

| Question | Answer |
|---|---|
| "How many rules?" | 7,012 total: 246 hand-curated + 6,766 AI-mined |
| "What did the AI cost?" | $27 one-time (2026-04-30 Haiku 4.5 mining); $0 ongoing |
| "How long to classify a single record?" | Sub-second |
| "How long for the full 784,556-row file?" | ~15 minutes |
| "How many users can hit it concurrently?" | Streamlit Cloud free tier handles small-team concurrent use; not stress-tested at City-wide scale |
| "Can I trust the result?" | Every result carries the exact rule that fired, confidence level, and the rule's reasoning |

## What This Document Is For

This is **Layer 1 factual input** for the Leadership Intelligence Project (Layer 3). Layer 3 builds the executive narrative, persuasion architecture, and stakeholder-specific framing on top of these facts. The facts are defensible line-by-line.
