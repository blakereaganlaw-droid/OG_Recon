# OG_Recon — UT Cash Management Reconciliation Engine

A production-grade, deterministic **forward matching engine** for University
of Tennessee (UT) bank accounts in Oracle Cash Management (DASH): it matches
open bank statement lines (BSL) to open system transactions (ST), classifying
each line **Match / Candidate / Review**. Every available source feeds the
candidate pool — ST exports, Receivables receipts, and the MET/ORT chain
(the `d:`/`r:` deposit-receipt bridge).

The **backward un-reconciliation engine** (re-audit reconciled groups →
recommend unwinds) lives in its own project,
[Unreconcile2](https://github.com/blakereaganlaw-droid/Unreconcile2), so this
engine stays fast: an `*All_Data*` workbook is recognized but never loaded
here (~4x faster on history-heavy accounts).

The full behavioral contract is the build spec
(`UT_Recon_Engine_BUILD_SPEC.md`); this codebase implements its forward half
(§10, the backward engine, is implemented by Unreconcile2).

## Design temperament (on rails)

- **Deterministic** — same inputs → identical outputs. No float money, no
  randomness, no time-of-day behavior. Money is integer cents via `Decimal`.
- **Fail loud, never silent** — an unresolved required file/column/relationship
  raises a named exception (`InvalidSourceData`, `MissingRequiredFile`,
  `AmbiguousColumn`) naming the file, the role, and the candidates.
- **Every row accounted for** — each BSL appears exactly once across the three
  output tabs; each ST is consumed at most once across Matches and Candidates.
  Both are asserted at the end.
- **Fixed pipeline** — a later stage never overrides an earlier one;
  availability is re-derived inside every loop, never precomputed once.

## Files

| File | Purpose |
|---|---|
| `recon_engine.py` | Self-contained forward engine: primitives → router → binder → pool → forward P0–P10 → workbook writer → `run()` orchestrator + JSON run log. |
| `recon_audit.py` | Independent audit (imports nothing from the engine; re-parses raw sources with its own binder and enforces C1–C10). |
| `run_recon.py` | Per-run (per-upload) wrapper: stages one upload into an immutable run folder, pre-flights the routing, records a provenance manifest, then runs the engine + audit. |
| `test_recon.py` | Unit tests for primitives, router, binder, pool dedup, the per-run wrapper, plus a synthetic end-to-end run gated by the audit. |

## Requirements

- Python 3.10+
- `openpyxl` (`pip install openpyxl`); `pyxlsb` for `.xlsb` binary exports
- Standard library + `decimal` only. **No pandas** (it float-coerces
  zero-padded references and merchant IDs and silently corrupts keys).

## Usage

### Per-run (per-upload) — recommended

Each upload of export files is one run. `run_recon.py` stages the upload into
an isolated, immutable run folder, pre-flights it, and runs the engine:

```bash
# Point it at the uploaded files (a folder, individual files, or a mix).
python3 run_recon.py /path/to/upload_folder
python3 run_recon.py /root/.claude/uploads/<session>/          # a web upload
python3 run_recon.py 20240101_FHB_UTC_BSL.xlsx 20240101_FHB_UTC_Account_ST.xlsx \
        --runs-root ./runs --run-id 2026Q3_FHB_UTC
```

Each run produces `runs/<run_id>/` containing:

- `input/` — the staged copies actually reconciled (Claude-web hex upload
  prefixes like `933782d6-` are stripped; real `YYYYMMDD` date prefixes,
  which drive the router's newest-wins ordering, are kept).
- `manifest.json` — provenance: every file's origin, size, SHA-256, and router
  role; unrouted files and warnings; the SHA-256 of the engine/audit code and
  git commit that processed the run.
- `outputs/` — the reconciliation workbook and the JSON run log (below).

Pre-flight fails loud **before** the engine runs on an unusable upload: no
spreadsheets, no BSL file, a mixed-account upload (any two files naming
different account tokens), or a staged-filename collision (case-insensitive) —
and the failed run folder is removed so the run-id stays free. Exit codes:
`0` ran + audit PASS; `2` ran but audit FAILed — `outputs/` is quarantined
(the workbooks are on disk for forensics but not approved for delivery);
`1` upload unusable.

### Direct engine invocation

```bash
# Reconcile one account's folder of export files.
python3 recon_engine.py /path/to/account_input_dir -o ./outputs
```

Outputs written to `./outputs/`:

- `<ACCOUNT>_reconciliation.xlsx` — tabs **Matches**, **Candidate Matches**,
  **Review Notes** (navy headers, Carlito 11pt, freeze A4, static values only).
- `<ACCOUNT>_runlog.json` — files routed, roles/columns bound, pool sizes by
  source and status, per-pass placements, and the audit result. This is the
  on-rails proof that no step was skipped.

For the unwind/forensic workbook, run **Unreconcile2** on the same folder.

The `run()` orchestrator gates delivery on the audit: if the independent audit
does not pass, it raises and the workbooks are withheld (override for debugging
with `--no-present-gate`).

Run the audit standalone:

```bash
python3 recon_audit.py /path/to/account_input_dir ./outputs/<ACCOUNT>_reconciliation.xlsx <ACCOUNT>
```

## File router

The engine never assumes a fixed file set or column order. It scans the input
folder, classifies each file by a case-insensitive substring test on its
**filename** (`recon_engine.ROUTER_TABLE`), then binds each needed column by
**content first, header as tiebreak** (`bind_columns`). The account is inferred
from the BSL/ALL_DATA filename token (`FHB_UTC`, `Regions_UTM`, …).

## Layout robustness (columns may move)

Column position never matters: every role is bound by scoring **all** columns
(content predicate first, header alias as tiebreak), the header row is
**located** by scanning the first 12 rows (never assumed to be row 0), and the
independent audit re-binds with the same vocabulary so both sides pick the
same columns. Verified end-to-end: reordered/reversed columns, inserted decoy
columns (constant dates, row IDs, notes), preamble rows above the header,
alias renames, and ragged/trailing cells all produce **cell-for-cell identical
workbooks** (pinned by `TestColumnRobustness`).

When layout genuinely cannot be resolved, the run fails loud, never guesses:

- Two columns tying on content **and** header for a required role →
  `AmbiguousColumn` naming the file, role, and candidate columns.
- No recognizable header row → `InvalidSourceData` (`HEADER`) instead of a
  positional row-0 guess (which would emit phantom rows from report preambles).
- Optional roles never bind by position alone: zero evidence or a blind tie
  leaves the role unbound rather than grabbing the leftmost column.
- Two files tying for one role's newest-date slot → `InvalidSourceData`
  (nothing is silently ignored); 8-digit account numbers in filenames are not
  mistaken for date stamps.
- A MID-master row naming two distinct GL strings, or remapping a MID →
  `InvalidSourceData` (no rightmost-column-wins).

Known audit-scope limitation: the audit re-parses the standalone BSL and MET
sources; other pool sources (ST, Receipts) are validated by the engine's own
fail-loud binder and the unit tests rather than an independent re-parse.

## Tests

```bash
python3 -m unittest test_recon -v
```

The end-to-end test builds a synthetic FHB_UTC account, runs the full pipeline,
and asserts the audit passes and conservation holds. **Real-data validated**
(2026-07-10) against the FHB Master UNR exports (283 bank lines, 13,872 STs,
226k-row all-accounts MET, ORT `.xlsb`): the run completes with audit PASS and
`TestRealDataShapes` pins every real export shape synthetically (the real
files are never committed). The MET export spans every UT account and is
filtered by the §4.6 scope join; MET↔ST transactions bridge 1:1 and are never
double-pooled; the ORT deposit-group chain (`d:` sums with reference/payer/
deposit-type corroboration) is P4 phase 2.

## Scope status

Every forward spec section is implemented: primitives (§7), router (§4),
binder (§5), pool with keep-largest dedup (§8), forward passes P0–P10 (§9),
date doctrine (§11), workbook standard (§13), independent audit C1–C10 (§14),
and the `run()` orchestrator with JSON run log (§15). The backward engine
(§10) is implemented by the separate Unreconcile2 project. The State/Edison lane activates when its Edison source files are present
and otherwise records a typed "not-run" reason in the run log, routing
affected BSLs to Review; the Merchant/MID, SPN, and named-payer lanes run
over the always-loaded pool. The pipeline never silently continues as if a
stage succeeded.
