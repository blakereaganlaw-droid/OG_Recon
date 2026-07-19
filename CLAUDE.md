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
python3 -m unittest test_recon -v                   # 112 tests
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
   ST is consumed at most once across Matches + Candidates + Misdirected
   (shared `Ledger`). Both asserted at end of `forward_reconcile`.
6. **Fixed pipeline.** A later pass never overrides an earlier one.
   Availability is derived through the ledger (`ledger.is_available`) at
   EVERY access; indexes may pre-bucket entries by immutable fields
   (amount, reference znorm, SPN, MID, source — set before
   `forward_reconcile` starts), but never by availability.  The
   cross-reference tie relation is precomputed once (`_build_tie_index`,
   6-gram candidate generation + exact verification — provably equal to
   the per-BSL full-pool scans it replaced, order included).
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
8c2. **Ref-tied split outranks amount-only coincidence (owner, 2026-07-17,
   FHB UTIA $40 merchant line).** In P4 phase 2, a deposit whose members
   carry the BSL's reference but sum to it only WITH closed members
   (auto-rec split) outranks coincidental equal-sum OPEN deposits that
   carry no tie — the ref-tied split surfaces as the
   `POSSIBLE_AUTO_REC_SPLIT` Candidate instead of the coincidence's
   amount-only path barring the line into Review. Corroborated exact-open
   deposits still win over splits.
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
8g. **Misdirected transactions (owner HARD GUARDRAIL, 2026-07-18).** The
   bank deposit can land in ONE account while the system transaction /
   receipt is booked to a DIFFERENT bank account — any source (AR/AP/EXT),
   any bank (FHB/Regions), any campus/institute/department. Real case:
   City of Memphis $70,992.66 hit FHB UTHSC while receipt 300045836 remits
   to FHB Master ("300045836 ST does not exist" on UTHSC). ALWAYS search
   for this. Foreign-account entries (receipts whose
   `Remittance Bank Account` maps to another configured account and have
   no ST in this account's export; MET rows scoped out to another
   configured account) join a SHADOW pool — `foreign_account` set,
   `available=False`, never eligible for Matches/Candidates. Pass
   `PM_misdirected` (before P9c) pairs an unplaced BSL with a shadow entry
   on exact signed cents AND a reference tie (never amount alone; NO
   payer-contradiction screen — reference ties outrank payer text, e.g.
   "STATE OF NE" paying for "University of Nebraska"). Placements go to
   the dedicated **Misdirected** workbook tab (4th tab, same 19 columns),
   explanation naming the foreign account — they can never auto-reconcile
   without a reroute/ECT. Conservation spans all four tabs; audit enforces
   citation+sum, foreign-account naming ("booked to"), non-reuse, and the
   4-tab structure. **Reverse direction (ALL_BSL, owner 2026-07-19):** the
   all-accounts open-BSL export (`Oracle_OTBI_All_BSL_UNR`, role `ALL_BSL`)
   feeds a read-only mirror search (`_reverse_misdirected`) — a THIS-account
   open, unconsumed ST/receipt whose bank line landed in ANOTHER account
   (exact signed cents + znorm-EXACT unique reference; MID/shared-originator
   refs excluded). ST-anchored, so it is a runlog finding
   (`reverse_misdirected`), never a workbook placement (BSL conservation
   untouched).
8g2. **PAYMENTS feed (owner, 2026-07-19).** `Oracle_Payables_Payments`
   (role `PAYMENTS`, tokens `payables`+`payments`) is the AP analogue of
   RECEIPTS: `Payment Number` is the AP identity (= ST Transaction Number on
   AP rows), `Payment Amount` is a POSITIVE disbursement magnitude the pool
   NEGATES to the bank's signed cents, and only OUTSTANDING payments
   (`Negotiable`; Cleared=Reconciled/Voided excluded, `OPEN_PAYMENT_STATUSES`)
   enter the pool as open AP entries — merged onto the ST export's AP rows by
   payment number at equal signed cents, else appended. On real FHB data this
   adds ~8,900 open AP payments the UNR ST export lacks.
8h. **Orphan doctrine (owner, 2026-07-19 — `UT_Recon_Forward_Orphan_Doctrine.md`,
   binding).** The open pool is shaped by prior (often automated, often wrong)
   reconciliations: it CONTAINS orphans (open STs whose money is already
   consumed) and is MISSING consumed counterparts. Incorporated forward rules:
   **R3 — MET status outranks the ST export**: the MET↔ST bridge now
   propagates a closed MET status onto the open-looking ST
   (`available=False`, runlog `met_status_overrides`) — never consume a
   candidate whose MET row is REC/cleared. **Dual-fire twin guard (2.3)**:
   available entries sharing one MET `RECEIPT_ID` at the same signed cents
   keep exactly ONE available (deterministic; runlog `dual_fire_twins` — 33
   live on real UTIA data). **R8 — check rails**: the check number IS the
   identity; same amount + different check number is a CONFLICT
   (`_check_number_conflict`, both the P9b amount-only lane AND the
   P9_payables candidate lane) — the FHB AP $1,100 cascade; blank
   references are silence. **Signature #6 — cherry-pick
   split**: a stranded line with no exact counterpart but an ALL-closed MET
   deposit summing exactly gets an enriched Review naming d:/members
   (`POSSIBLE_AUTO_REC_SPLIT` → run Unreconcile2), never a forced match.
   Already covered elsewhere: R1 (amount never identity, rule 4), R6 (type
   gates, 8e), R7 (Regions dual-encoded deposit refs — never build a
   deposit-slip ref-tie check), R9 (refer, don't route around — Review with
   named evidence). R2/R5 (Recon History `Created By`, CFG_TCR coverage)
   activate only when those exports are present in a run.
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
  via pyxlsb. **`.xlsb` integral-float normalization (owner, 2026-07-16):**
  pyxlsb returns EVERY numeric cell as a float, so an integral Oracle id
  (`DEPOSIT_ID`/`RECEIPT_ID` = 65105) arrives as `65105.0` and stringifies to
  `"65105.0"` — corrupting the MET↔ST bridge join and the `d:/r:` citations
  (audit C4). `_xlsb_norm` collapses integral floats back to int at read time
  in BOTH the engine (`_read_xlsb_rows`) and the audit's independent
  `_read_xlsb`, so an `.xlsb` MET yields byte-identical output to its `.csv`
  twin (money and true fractional serials untouched). The audit also now reads
  MET/BSL `.xlsb` (it previously only accepted `.xlsx/.xlsm/.csv`, so an
  `.xlsb` MET made its real-deposit set empty and every citation failed C4).
  Real exports are NOT committed; `TestRealDataShapes` pins their
  shapes synthetically. The relationship docs
  (`UT_Recon_ORT_Data_Relationships.md`, SPN companion,
  `UT_Recon_Forward_Orphan_Doctrine.md`) are the domain
  authority for joins/gates alongside the spec.
- **COA GL scope key (owner, 2026-07-18).** The Chart of Accounts assigns
  each depository bank account a natural-account GL code, stamped in the
  MET's `ASSET_CONCATENATED_SEGMENTS` (`ENTITY-FUND-DEPT-ACCOUNT-...`).
  `_GL_CASH_ACCOUNTS` / `account_of_gl_segments` map it to the engine
  account (100210=FHB_MASTER, 100221=FHB_UTIA, 100310=FHB_UTC,
  100330=REGIONS_UTM, 100500=FHB_UTHSC, …; clearing/payroll GLs
  deliberately unmapped). Used as FALLBACK scope for the MET pool when the
  bank-name column is unbound (else an all-accounts export leaks every
  account); when both keys bind, bank name stays authoritative and GL
  disagreements surface as `met_gl_conflicts` in the runlog (real exports
  carry them: 38 UTHSC-GL rows inside the Master MET, 34 Master-GL rows
  inside the UTIA MET). Entity segment codes (10=UTK, 40=UTC, 50=UTM,
  60=UTSO, 70=UTHSC, 17/18/19=UTIA family) are reference only — campus
  "consistency" confers NOTHING (rule 8c).
- **COA decode bundle — Tier 1 (owner, 2026-07-19).** The Chart of Accounts
  reference set (role `CHART_OF_ACCOUNTS`: the seven `AcctCombos` shards +
  `Segments.csv` + `ComboSets`/`CombosTech`) is a multi-file, OPTIONAL bundle
  loaded by `load_chart_of_accounts` into `{combo_decode, entity_desc,
  postable_efdp}` (5,515 distinct combos / 29 entities / 1,257 postable
  E-F-D-P keys on real data). Its ONLY job is to DECODE a GL combo into human
  labels (`coa_decode`, `segments_of`/`dept_segment_of`/`entity_segment_of`)
  for two NON-GATING text surfaces: `_p10_review_cause` annotates each named
  ST that carries `asset_segments` with "(dept …, entity …)", and
  `recommend_gl_string` falls back — after a MID_MASTER miss — to the
  exact-amount counterpart's decoded combo. It is ADVISORY: placements stay
  byte-identical with the bundle present or absent (campus/entity consistency
  confers nothing — rule 8c); when absent, every consumer no-ops. The loader
  skips the huge `ORT_Activity_GL_Departments` routing table and
  `RelatedValueSets` (out of scope). Diagnostic only: build_pool logs
  `coa_combo_validity` (rows_seen / unrecognized_combo / non_postable_efdp)
  beside `met_gl_conflicts` — a counter that never drops or downgrades a row.
  The entity-divergence downgrade (Tier 2) is designed but not yet wired.
  **Recommended GL comes from the OFFSET combo only (adversarial review,
  2026-07-19):** `recommend_gl_string`'s CoA fallback decodes the exact-amount
  counterpart's `OFFSET_CONCATENATED_SEGMENTS` (the ECT posting side, §6) —
  NEVER the cash-side asset combo (its account segment is the bank's own cash
  GL), and NEVER a foreign-account shadow entry (rule 8g: another
  depository's posting must not steer this account's ECT).  P10 builds one
  amount index for the whole residual pass (no per-line pool rescans).
- **CM configuration audit (owner, 2026-07-19 — orphan-doctrine R5
  activation).** The five `CM_Configurations_*` exports load as `CFG_TCR` /
  `CFG_PARSE` / `CFG_MATCHING` / `CFG_TOLERANCE` / `CFG_RULESETS`
  (`load_cm_config`; paginated-export reader `_cfg_read_table` keeps ORPHAN
  TCR rows with a blank bank account — 23 live on the real export, 3 still
  enabled).  When present, `config_audit` (end of `run()`, AFTER writer +
  audit) REPLAYS the TCRs against this run's lines via `like_match` (Oracle
  LIKE, case-sensitive; a case-insensitive-only hit is the case-bug
  signature) and classifies every Review line: creation FAILURE (claimed by
  an enabled rule, no pool entry at the amount), fired-but-stranded, claimed
  only by a DISABLED rule (the Student Refund smoking gun), or uncovered
  (recurring signatures → proposed CREATE-TCR search strings).  Static
  checks: null trx codes, orphan rules, duplicate enabled pairs, CASH=None,
  CASH-GL posting to another depository.  Parse-rule gaps per bank family
  (codes with blank references and no rule).  Output: runlog `config_audit`
  + `<account>_config_recommendations.md/.json` in outputs — NEVER the
  locked 19-col workbook; placements provably byte-identical with configs
  present (tested).  All fire counts labeled SIMULATED (Oracle evaluates
  untruncated addenda; the export truncates ~1000 chars).
- **Native BAI2 `.txt` (owner, 2026-07-19).** `route_folder` accepts `.txt`
  ONLY when it classifies as BAI2; `_read_bai2_txt` parses the raw
  transmission (01/02/03/16/88 records) into rows binding `BAI2_ROLES` —
  full untruncated 88-continuation addenda in DETAIL columns (`Customer
  ID:`, `Trace Number:`, `TRN1` reassociation keys), BAI sign convention
  (100-399 credit / 400-699 debit), group as-of dates, integer cents.  Real
  UTHSC file: 18,793 details.  Feeds the existing BAI2 enrichment — richer
  payer/MID visibility legitimately moves weak amount-only placements into
  the correct guardrail lanes (validated: $0.01 penny-test and stale-MID
  cases).  Raw-file vocabulary differs from the Oracle feed's relabeling
  (`Company Name:` ↔ `SENDING CO NAME:` etc.) — parse-rule work must cover
  BOTH vocabularies (dual twins are safe no-ops).
- Student Refund depositories now registered for ALL five campuses (UTK,
  UTC, UTHSC, UTM, UTSO — the TCR export names each); the long-form
  "FHB - Accounts Payable" maps to FHB_AP.
- **Edison (State of TN) annotation (owner, 2026-07-19).** `EDISON_PAY` /
  `EDISON_INV` now LOAD (Reference = the State's zero-padded 10-digit
  payment id; Amount positive dollars; invoices carry Approval Status +
  Voucher).  ANNOTATION ONLY — the C6 State pass stays retired; State lines
  reconcile through the normal lanes.  `_edison_note` (P10) cites a payment
  on exact signed cents AND a reference-digit tie to the line text
  (reference outranks amount); an amount-only SINGLETON is cited as
  uncorroborated; ambiguous amount-only sets are never guessed.  Real
  Master data: 14 of 39 stranded State lines named, all reference-tied.
  Edison records are the payer's, never pool entries — they can never
  place anything.
- UNR-only exports are residuals: Oracle already took the easy matches, so
  low Match counts with precise Candidate/Review causes are CORRECT there,
  not a defect. Receipts/Edison/GMS exports enrich what can match.
