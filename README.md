# OG_Recon — UT Cash Management Reconciliation Engine

A production-grade, deterministic reconciliation engine for University of
Tennessee (UT) bank accounts in Oracle Cash Management (DASH). It runs two
engines over one account's export files:

- **Forward reconciliation** — matches open bank statement lines (BSL) to open
  system transactions (ST), classifying each line **Match / Candidate /
  Review**.
- **Backward un-reconciliation** — re-audits already-reconciled groups against
  doctrine and recommends **unwinding** the unsound ones.

The full behavioral contract is the build spec
(`UT_Recon_Engine_BUILD_SPEC.md`); this codebase implements it.

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
| `recon_engine.py` | Self-contained engine: primitives → router → binder → pool → forward P0–P10 → backward → workbook writers → `run()` orchestrator + JSON run log. |
| `recon_audit.py` | Independent audit (imports nothing from the engine; re-parses raw sources with its own binder and enforces C1–C10). |
| `test_recon.py` | Unit tests for primitives, router, binder, pool dedup, plus a synthetic end-to-end run gated by the audit. |

## Requirements

- Python 3.10+
- `openpyxl` (`pip install openpyxl`)
- Standard library + `decimal` only. **No pandas** (it float-coerces
  zero-padded references and merchant IDs and silently corrupts keys).

## Usage

```bash
# Reconcile one account's folder of export files.
python3 recon_engine.py /path/to/account_input_dir -o ./outputs
```

Outputs written to `./outputs/`:

- `<ACCOUNT>_reconciliation.xlsx` — tabs **Matches**, **Candidate Matches**,
  **Review Notes** (navy headers, Carlito 11pt, freeze A4, static values only).
- `<ACCOUNT>_unwind.xlsx` — tab **Unwind Recommendations** (dark-red forensic
  header).
- `<ACCOUNT>_runlog.json` — files routed, roles/columns bound, pool sizes by
  source and status, per-pass placements, backward defects by code, and the
  audit result. This is the on-rails proof that no step was skipped.

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

## Tests

```bash
python3 -m unittest test_recon -v
```

The end-to-end test builds a synthetic FHB_UTC account, runs the full pipeline,
and asserts the audit passes and conservation holds. **Validating against real
UT data** (per spec Section 16 step 8 — reproducing the known Oracle groups on
a fully-reconciled account such as FHB UTC) requires the actual export files,
which are not committed here; drop them into an input folder and run the CLI.

## Scope status

Every spec section is implemented: primitives (§7), router (§4), binder (§5),
pool with keep-largest dedup (§8), forward passes P0–P10 (§9), backward engine
with defect codes (§10), date doctrine (§11), workbook standard (§13),
independent audit C1–C10 (§14), and the `run()` orchestrator with JSON run log
(§15). Optional lanes (State/Edison, Merchant/MID, SPN, named-payer) activate
when their source files are present and otherwise record a typed "not-run"
reason in the run log, routing affected BSLs to Review — the pipeline never
silently continues as if a stage succeeded.
