# CLAUDE.md — OG_Recon operating memory

Auto-loaded every Claude Code session. Read before touching the engine.

## What this is

A deterministic, "on-rails" **forward-only** reconciliation engine for
University of Tennessee (UT) bank accounts in Oracle Cash Management (DASH):
match open bank statement lines (BSL) to open system transactions (ST) →
**Match / Candidate / Review**. The pool draws on every available source —
ST, Receivables receipts, and the MET/ORT chain.

The **backward** engine (re-audit reconciled groups → recommend unwinds) was
split into the separate **Unreconcile2** repo for speed. An `*All_Data*`
workbook is recognized here but never loaded — do not reintroduce it.

The binding behavioral contract is **`UT_Recon_Engine_BUILD_SPEC.md`**. When code
and spec disagree, the spec wins — fix the code.

## Files

| File | Role |
|---|---|
| `recon_engine.py` | Self-contained forward engine + CLI. Primitives (§7), router (§4), binder (§5), pool (§8), forward P0–P10 (§9), writer (§13), `run()` + JSON run log (§15). Backward (§10) lives in Unreconcile2. |
| `recon_audit.py` | Independent audit (§14). **Imports nothing** from the engine — keep it that way. Re-parses raw sources with its own binder; enforces C1–C10; gates delivery. |
| `run_recon.py` | Per-run (per-upload) wrapper. Stages one upload into an immutable `runs/<run_id>/` folder (strips 8-hex upload prefixes; keeps plausible `YYYYMMDD` date prefixes), pre-flights routing (fails loud on no-BSL / any mixed-account token / case-insensitive staged-name collisions, and removes the failed folder so the run-id stays free), writes a SHA-256 provenance `manifest.json`, then calls `recon_engine.run`. Exit 0 = audit PASS; 2 = audit FAIL (outputs quarantined — written for forensics, not approved for delivery); 1 = unusable upload. |
| `test_recon.py` | Unit + synthetic end-to-end tests (engine and per-run wrapper). |
| `UT_Recon_Engine_BUILD_SPEC.md` | The spec. Binding. |

## Run / test

```bash
python3 run_recon.py <upload_dir_or_files>          # one upload = one run folder
python3 recon_engine.py <input_dir> -o ./outputs    # direct engine invocation
python3 -m unittest test_recon -v                   # 56 tests
```

Web sessions install deps via `.claude/hooks/session-start.sh`; locally,
`pip install -r requirements.txt`.

## Non-negotiable doctrine (do not "improve" these away)

1. **Integer cents only.** Money is `cents()` → signed int. **Never** float
   math, never rounding/tolerance/fuzzy matching. Exact signed-cent equality.
2. **No pandas.** It float-coerces zero-padded refs and merchant IDs. Stdlib +
   `openpyxl` + `decimal` only.
3. **Fail loud.** Unresolved required file/column/relationship → raise
   `InvalidSourceData` / `MissingRequiredFile` / `AmbiguousColumn` naming the
   file, role, and candidates. Never guess a column; never silently drop a row.
4. **Amount alone never makes a Match** — with ONE owner exception
   (2026-07-11): a **distinctive amount** (non-zero cents, >= $1,000) that is
   unique on BOTH sides (one open counterpart in the pool/deposit chain, one
   bank line at that amount) is valid match evidence (`amount_distinctive`,
   Medium confidence, `DISTINCTIVE_AMOUNT`). Everything else requires exact
   amount **and** corroboration (reference tie, or payer tie for named-payer
   rules). Transaction type alone confers nothing. Date supports, never
   suffices.
5. **Conservation.** Each BSL appears exactly once across the three tabs; each
   ST is consumed at most once across Matches + Candidates (shared `Ledger`).
   Both asserted at end of `forward_reconcile`.
6. **Fixed pipeline.** A later pass never overrides an earlier one. Availability
   is re-derived inside every loop (`ledger.is_available`), never precomputed.
7. **Guardrails:** Journal-source STs never match bank lines. A MID receipt
   reconciles only through the merchant lane. Sibling references (equal-length
   numerics differing in the last 1–2 digits) are **conflicts**, never ties.
8. **Dedup before summing.** Keep-largest applies only to the
   total-plus-splits signature (largest == sum(rest), signed cents); borrow
   counterparty from a dropped split. Same-key rows that don't sum that way
   are distinct receipts sharing a label — ALL kept, ids disambiguated
   `<id> [<amount>]` with `base_id` preserving the MET bridge join. Engine
   and audit dedup independently.
8b. **Payer contradiction (owner, 2026-07-11).** Zero-corroboration
   (amount-only) pairings are barred — even as Candidates — when both sides
   carry payer tokens and share none ("City of Chattanooga has nothing to do
   with Israel"). Reference ties outrank payer text; silence never
   contradicts. **Feed-session silence (owner, 2026-07-16):** Oracle ORT
   stamps unattributed External lines with a generic `FEED SESSION <n>` batch
   label in the counterparty column — that names the load batch, not a
   payer/beneficiary, so it is SILENCE on payer identity and never
   contradicts a real bank-side payer (`_FEED_SESSION_RE` strips it before
   `_contra_tokens`). This is what lets a distinctive-amount Heartland
   settlement match its ORT deposit chain instead of being falsely
   contradicted. **Payer-family aliases (owner, 2026-07-14):** distinct trade
   names for one payer are agreement, not contradiction — VSHP / Volunteer
   State Health Plan / TN CARE SELECT / TennCare **==** BlueCare Tennessee
   (BCBST) via `payer_family`. Pass `P8c_payer_family`: a positive ACH covered
   by ≥1 same-day Receivables receipts of the same payer family whose own refs
   don't tie the bank line, summing EXACTLY (whole same-day same-family group,
   never subset-sum), is a `PAYER_FAMILY_GROUP` **Candidate** (never a Match —
   the addenda's SPNs may cite other receipts).
8c. **Deposit-type / merchant / correction (owner, 2026-07-11).**
   "Deposit-type consistency" confers NOTHING — an exact-sum deposit group
   without a reference or payer tie is a plain amount-only Candidate, never
   a Match. Merchant-lane (MID) lines corroborate deposit groups ONLY via a
   reference/MID tie. Deposit-correction lines are manual fixes — Candidates
   flagged `MANUAL_ECT` at best, never amount-sum Matches; edge cases err
   toward Candidate over rejection. Audit C7 enforces all three. Convera
   lines are international wires and ALWAYS Payables — they never pair with
   a non-Payables ST (central `_type_gate_ok`). Chargeback / merchant-fee
   DEBITS (owner, 2026-07-12) pair ONLY on MID equality: when the bank line
   carries a MID, an ST without that same MID is barred even as a
   Candidate ("the MID is the critical matching string"). Audit C7
   enforces it.
8d. **12-day stale-candidate ceiling (owner, 2026-07-12).** An
   External-source ST entered 12+ days BEFORE the BSL statement date is
   almost certainly not the counterpart — barred even as a Candidate
   (`_ext_stale_barred`, all four stale-candidate sites). 8-11 days stale
   may still surface as a flagged Candidate; Matches stay barred at 8+;
   BSL-before-ST stays unbounded-valid; non-EXT sources untouched. Audit
   C8 enforces it on the Candidate tab.
8e. **Type incongruence (owner, 2026-07-13).** Transaction type is
   important but NOT dispositive. An electronic (ACH/EFT) ST paired to a
   Miscellaneous bank line on amount ALONE — no reference, payer, or MID
   tie — is almost certainly a coincidence and is barred even as a
   Candidate (`_type_incongruent_uncorroborated`, the amount-only lanes
   P9b + P4 deposit groups). ANY tie overrides (type is not dispositive).
8f. **ORT / Receivables 1:M reference search (owner, 2026-07-14).** The ORT
   chain (External STs) and the Receivables ST reference columns are ALWAYS
   searched for 1:M ties (pass `P9c_ref_1m_review`). When a bank line's
   reference is `znorm`-EXACTLY equal (never containment; ≥6 chars) to the
   `Reference`/`Transaction Number` of ≥2 open STs that share the BSL's sign
   and sum SHORT of it (`|group| < |BSL|` — the rest already auto-reconciled/
   stranded, absent from the UNR export; opposite-sign or oversized groups are
   shared-originator-ID collisions, excluded), the exact-sum guardrail bars a
   Match/Candidate, so the line
   surfaces as an ENRICHED REVIEW naming the tied members, the partial sum,
   and the shortfall (`PARTIAL_REFERENCE_GROUP` / `POSSIBLE_AUTO_REC_SPLIT` →
   run Unreconcile2). Exact-summing groups stay P4's; single coincidental
   ties and short/date-like refs are excluded. Review placements never
   consume their STs.
9. **Determinism.** No randomness, no clock. `Date.now`/serials excepted where
   parsing Excel. Sort candidate sets by (amount, date, id) before choosing.

## Conventions

- `datetime` is checked **before** `date` (`datetime` subclasses `date`).
- Falsy-zero guard: test `is not None`, never truthiness (0 cents is valid).
- **Position never binds.** Columns bind by content-first scoring over all
  columns; content samples the first 50 NON-BLANK values per column across
  the whole sheet (sparse columns score on what they carry); the header row
  is located, never assumed row 0; optional roles stay unbound on zero
  evidence or a blind tie — except verbatim-duplicate columns (bind leftmost)
  and signed/unsigned amount twins (bind the signed one); newest-file ties,
  ambiguous sheet substring matches, and MID-master GL conflicts fail loud. The audit re-binds
  with the same alias vocabulary (duplicated literals — still imports nothing
  from the engine). Column rearrangement is pinned identical-output by
  `TestColumnRobustness`.
- Output cells starting with `=` get a leading space (no formula injection);
  workbooks are **static values only**, zero formula cells.
- Output format is locked by §13 (HARD GUARDRAIL, owner 2026-07-11):
  Carlito 11pt, navy `FF1F4E78` header, freeze `A4`, 19 fixed columns
  carrying ALL BSL identifier fields and ALL ST detail fields, ST lists
  never truncated, no ST reused, no provenance rows. The audit's C9/C10
  will fail you if you drift. (The dark-red forensic/unwind book is Unreconcile2's.)

## Working rules

- Any change to matching logic must keep `python3 -m unittest test_recon` green
  and the synthetic end-to-end audit at `status: PASS`.
- **Real-data validated (2026-07-10, FHB Master UNR exports):** router,
  binder, MET scope join (all-accounts export filtered by long→short bank
  name), MET↔ST 1:1 bridge (13,715/13,863 bridged; never duplicated), native
  `DEPOSIT_ID`/`RECEIPT_ID` columns preferred over `d:/r:` description parse,
  ORT deposit-group pass (P4 phase 2), NA-placeholder nulling, integer
  reference cells (an int is a reference, not an Excel date serial), `.xlsb`
  via pyxlsb. Real exports are NOT committed; `TestRealDataShapes` pins their
  shapes synthetically. The relationship docs
  (`UT_Recon_ORT_Data_Relationships.md`, SPN companion) are the domain
  authority for joins/gates alongside the spec.
- UNR-only exports are residuals: Oracle already took the easy matches, so
  low Match counts with precise Candidate/Review causes are CORRECT there,
  not a defect. Receipts/Edison/GMS exports enrich what can match.
