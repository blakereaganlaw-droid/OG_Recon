# Claude Project — Custom Instructions (starter)

Paste the block below into your claude.ai **Project → Settings → Custom
Instructions**. (Connecting the GitHub repo gives Claude the *code*; these
instructions give it the *operating rules*. You need both.)

This file is a convenience copy; the authoritative, always-loaded version for
Claude Code sessions is `CLAUDE.md` in the repo root.

---

You are working on **OG_Recon**, a deterministic, **forward-only**
reconciliation engine for University of Tennessee bank accounts in Oracle Cash
Management (DASH): it matches open bank statement lines to open system
transactions → Match / Candidate / Review, drawing candidates from every
available source (ST exports, Receivables receipts, the MET/ORT chain). The
backward engine (re-audit reconciled groups → recommend unwinds) lives in the
separate **Unreconcile2** repo; an *All_Data* workbook is recognized here but
never loaded.

The binding contract is `UT_Recon_Engine_BUILD_SPEC.md`. When code and spec
disagree, the spec wins.

Hold these invariants on every change — they are doctrine, not preferences:

1. Money is **integer cents** via `Decimal`. Never float math, never
   rounding/tolerance/fuzzy matching. Exact signed-cent equality only.
2. **No pandas.** Standard library + `openpyxl` + `decimal` only.
3. **Fail loud:** an unresolved required file/column/relationship raises a named
   exception (`InvalidSourceData` / `MissingRequiredFile` / `AmbiguousColumn`)
   naming the file, the role, and the candidates. Never guess a column or drop a
   row silently.
4. **Amount alone never makes a Match** — it needs corroboration (a reference
   tie, or a payer tie for named-payer rules). Date supports, never suffices.
5. **Conservation:** each bank line lands in exactly one output tab; each system
   transaction is used at most once across Matches + Candidates.
6. **Fixed pipeline:** a later pass never overrides an earlier one; availability
   is re-derived inside each loop.
7. **Guardrails:** journal-source transactions never match bank lines; merchant
   (MID) receipts reconcile only through the merchant lane; sibling references
   are conflicts, never ties.
8. `recon_audit.py` must **import nothing** from `recon_engine.py`; it
   independently re-parses sources and enforces checks C1–C10.

How to run and verify:
- Reconcile one account: `python3 recon_engine.py <input_dir> -o ./outputs`
- Tests: `python3 -m unittest test_recon -v` — must stay green, and the
  synthetic end-to-end audit must report `status: PASS`, on every change.

Output format is locked by spec §13 (Carlito 11pt, navy header, freeze A4, nine
fixed columns, static values only, no provenance rows); the audit will fail on
drift. Keep changes minimal and deterministic; explain any doctrine trade-off
explicitly.
