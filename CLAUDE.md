# CLAUDE.md — OG_Recon operating memory

Auto-loaded every Claude Code session. Read before touching the engine.

## What this is

A deterministic, "on-rails" reconciliation engine for University of Tennessee
(UT) bank accounts in Oracle Cash Management (DASH). Two engines in one program:

- **Forward** (`recon_engine.forward_reconcile`, passes P0–P10): match open bank
  statement lines (BSL) to open system transactions (ST) → **Match / Candidate /
  Review**.
- **Backward** (`recon_engine.backward_reconcile`): re-audit already-reconciled
  groups against doctrine → recommend **unwinds**.

The binding behavioral contract is **`UT_Recon_Engine_BUILD_SPEC.md`**. When code
and spec disagree, the spec wins — fix the code.

## Files

| File | Role |
|---|---|
| `recon_engine.py` | Self-contained engine + CLI. Primitives (§7), router (§4), binder (§5), pool (§8), forward P0–P10 (§9), backward (§10), writers (§13/§10.6), `run()` + JSON run log (§15). |
| `recon_audit.py` | Independent audit (§14). **Imports nothing** from the engine — keep it that way. Re-parses raw sources with its own binder; enforces C1–C10; gates delivery. |
| `test_recon.py` | Unit + synthetic end-to-end tests. |
| `UT_Recon_Engine_BUILD_SPEC.md` | The spec. Binding. |

## Run / test

```bash
python3 recon_engine.py <input_dir> -o ./outputs   # reconcile one account
python3 -m unittest test_recon -v                  # 23 tests
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
4. **Amount alone never makes a Match.** Requires exact amount **and**
   corroboration (reference tie, or payer tie for named-payer rules). Date
   supports, never suffices.
5. **Conservation.** Each BSL appears exactly once across the three tabs; each
   ST is consumed at most once across Matches + Candidates (shared `Ledger`).
   Both asserted at end of `forward_reconcile`.
6. **Fixed pipeline.** A later pass never overrides an earlier one. Availability
   is re-derived inside every loop (`ledger.is_available`), never precomputed.
7. **Guardrails:** Journal-source STs never match bank lines. A MID receipt
   reconciles only through the merchant lane. Sibling references (equal-length
   numerics differing in the last 1–2 digits) are **conflicts**, never ties.
8. **Dedup before summing.** Keep the largest-magnitude row (the total, not an
   invoice split); borrow counterparty from a dropped split. Engine and audit
   apply the same keep-largest dedup independently.
9. **Determinism.** No randomness, no clock. `Date.now`/serials excepted where
   parsing Excel. Sort candidate sets by (amount, date, id) before choosing.

## Conventions

- `datetime` is checked **before** `date` (`datetime` subclasses `date`).
- Falsy-zero guard: test `is not None`, never truthiness (0 cents is valid).
- Output cells starting with `=` get a leading space (no formula injection);
  workbooks are **static values only**, zero formula cells.
- Output format is locked by §13: Carlito 11pt, navy `FF1F4E78` header (dark-red
  `FF7A1F1F` for the forensic/unwind book), freeze `A4`, 9 fixed columns, no
  provenance rows. The audit's C9/C10 will fail you if you drift.

## Working rules

- Any change to matching logic must keep `python3 -m unittest test_recon` green
  and the synthetic end-to-end audit at `status: PASS`.
- Real-data validation (spec §16 step 8: reproduce known Oracle groups on a
  fully-reconciled account like FHB UTC) needs the actual export files, which
  are **not** in the repo. The synthetic fixture stands in until they are.
