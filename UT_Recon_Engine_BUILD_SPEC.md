# UT Cash Management Reconciliation Engine — Master Build Specification

> **Addendum (2026-07-10, project owner decision):** this specification now
> governs **two repositories**. `OG_Recon` implements the **forward** engine
> only (all sections except §10); `Unreconcile2` implements the **backward**
> engine (§10, plus the shared §4 router, §5 binder, §7 primitives, §8 pool).
> The split exists for speed: the forward engine recognizes but never loads
> an `*All_Data*` workbook. Statements below describing "one program with two
> engines" predate the split and should be read per-repository.

**Audience:** an autonomous coding agent (Claude Code first; portable to Grok, ChatGPT, and Perplexity).
**Mandate:** write a complete, production-grade Python program that reconciles University of Tennessee (UT) bank accounts in Oracle Cash Management (DASH), forward (match open items) and backward (catch and unwind bad reconciliations).
**Non-negotiable temperament:** fast, accurate, reliable, thorough, resilient, and **on rails** — the pipeline is fixed, every stage validates its inputs, and the program never skips a step, never guesses at a column, and never silently drops a row.

> Read this whole document before writing any code. Implement every section. Do not summarize, reorder, or omit. Where this document gives pseudocode, treat the logic as binding and the syntax as a guide.

---

## 0. How to read and execute this document

1. **Determinism first.** Same inputs → identical outputs, every run. No randomness, no time-of-day behavior, no floating point.
2. **Fail loud, never silent.** If a required column, file, or relationship cannot be resolved, raise a named exception (`InvalidSourceData`, `MissingRequiredFile`, `AmbiguousColumn`) with the file, the role, and the candidates considered. Do not proceed on a guess.
3. **Every input row is accounted for.** Each bank statement line (BSL) appears exactly once in the output. Each system transaction (ST) is consumed at most once across Matches and Candidates. Assert both at the end.
4. **The pipeline is a fixed sequence of stages (Section 9 forward, Section 10 backward).** A later stage never overrides an earlier one. Availability is re-derived inside each loop, never precomputed once.
5. **Portability.** Use only the Python standard library plus `openpyxl` and `decimal`. No `pandas` (it float-coerces zero-padded references and merchant IDs and silently corrupts keys). The same file can be handed to Grok, ChatGPT, or Perplexity as a reasoning spec; the code must run unmodified under CPython 3.10+.

---

## 1. Mission and scope

### 1.1 Two engines in one program
- **Forward reconciliation** — take the open (unreconciled) bank statement lines for one account and match each to one or more open system transactions, classifying every line as **Match**, **Candidate**, or **Review**.
- **Backward un-reconciliation** — take the already-reconciled groups for the same account and re-verify each against doctrine; flag any group whose reconciliation is unsound (wrong amount, wrong reference, implausible date, one-to-many mis-pairing, illegal source) and recommend an **unwind**. This is not optional; catching and fixing bad reconciliations is core to the work.

### 1.2 Accounts
FHB (First Horizon Bank): Master, UTHSC, UTIA, UTC, UTM, UTSO, AP.
Regions: UTIA, UTIPS, UTM.
The engine is account-agnostic; the account is inferred from the file router (Section 4) and every match is scoped to a single account.

### 1.3 Deliverables per run
1. One reconciliation workbook (Section 13) with tabs **Matches**, **Candidate Matches**, **Review Notes**.
2. One un-reconciliation workbook (Section 10.6) with tab **Unwind Recommendations**.
3. One audit log (Section 14). The audit must pass before any workbook is presented.

---

## 2. Governing doctrine (binding on every stage)

1. **Exact signed-cent equality.** Amounts match on exact signed integer cents. No tolerance, no rounding, no fuzzy or similarity matching. Oracle's own amount tolerance is disabled; do not reintroduce it.
2. **Conservative classification.** A false Match is worse than an unreconciled item. On any ambiguity, route to Candidate or Review. Never force a pairing.
3. **A Receivable is just an ST.** Receivables receipts are system transactions matched on the same Amount + Reference logic as every other ST. The Receivables files supply the receipt's data; they are never a separate "SPN lane" with different match physics.
4. **Amount alone never makes a Match.** A Match requires exact amount **and** corroboration: a reference tie, or (for named-payer rules) a payer/counterparty tie. Date supports but never suffices.
5. **Date window is ±15 days** (Tolerance Rules, confirmed). Match requires the signed BSL−ST lag within the plausible band (Section 11). For State lines, receipt-entry lag of months is normal when the BSL precedes the ST.
6. **Dedup before summing.** Receipts and invoice applications are duplicated across export rows, and a receipt-total row coexists with its invoice-application splits under one transaction number. Deduplicate by transaction/receipt number, keeping the **total** (largest magnitude), before any group sum.
7. **No ST reuse.** One consumption ledger spans Matches and Candidates. An ST named in a Candidate is barred from later Match promotion. No Match without an available open ST.
8. **Anchor direction is BSL → ST.** The bank statement line is the anchor of every forward match.
9. **Guardrails:** Journal-source STs never match bank lines. A merchant-ID (MID) receipt is a card receipt; it reconciles through the merchant/reference lane and never to a non-merchant line by coincidence. Sibling references (equal-length numerics differing only in the last one or two digits) are conflicts, not partial matches.

---

## 3. Tooling and environment

- **Language/runtime:** Python 3.10+. Standard library + `openpyxl` + `decimal` only. `pyxlsb` only if an `.xlsb` appears.
- **Money:** integer cents via `cents(x) = int((Decimal(str(x)) * 100).quantize(Decimal("1")))`. Never use float for money or references.
- **Workbook loads:**
  - Data workbooks: `load_workbook(path, read_only=False, data_only=True)`. `read_only=True` falsely reports `max_row == 1` on Oracle BI exports — do not use it for row iteration you depend on; use it only for a cheap header peek if needed, then reopen.
  - Configuration workbooks: `load_workbook(path, data_only=False)` to read formula strings in the advanced-criteria column.
- **CSV loads:** `csv.reader` with `encoding="utf-8", errors="replace"`; strip a leading UTF-8 BOM from the first header cell.
- **Guards:** falsy-zero guard everywhere (`if x is not None`, never truthiness). `isinstance(v, datetime)` checked **before** `isinstance(v, date)` (datetime subclasses date). Never name a module `inspect.py`. Prepend a space to any output cell whose text starts with `=`.
- **Paths (Claude Code default; override by argument):** inputs in the working input folder; scripts in `./`; outputs in `./outputs/`.
- **Portability to Grok / ChatGPT / Perplexity:** ship the program as a single self-contained module plus this spec. Those models either (a) execute the Python directly if they have a code tool, or (b) follow this spec as an operating procedure over the same files. Keep all logic in pure functions with explicit inputs so the spec doubles as a runbook. Do not depend on any Claude-specific tool, filesystem quirk, or network access.

---

## 4. File router — identify every file by NAME, then bind by header

The engine never assumes a fixed set of files or a fixed column order. It scans the input folder, classifies each file by a case-insensitive substring test on its **filename**, and records its role. Column meaning is resolved later, by header/content (Section 5). This is the "on rails" contract: the router decides *what a file is for*; the binder decides *which column is which*.

### 4.1 Router table

Match the **first** rule whose `contains` tokens all appear (case-insensitive) in the filename. Some roles accept several filename shapes.

| Role (internal key) | Filename contains (all of) | Sheet to use | Status coverage | Required? |
|---|---|---|---|---|
| `BSL` (anchor, open bank lines) | `BSL` and (`UNR` or not `All_Data`) | first/`Exported` | open | **yes** |
| `ST` (open system transactions) | `_ST` or `Account_ST` | first/`Exported` | open | yes if no All_Data |
| `ALL_DATA` (all-status lifecycle) | `All_Data` | (multi-sheet, see 4.2) | REC + UNR + VOID | yes for backward; strongly preferred for forward |
| `MET` (OTBI ORT→ST bridge) | `MET` | first | all-status if `All` in name | yes |
| `RECEIPTS` (Receivables, all-status) | `Receivables_Receipts` or `Receipts_All` | `Export to Excel`/first | all statuses | yes |
| `ORT_AR` | `ORT` and `AR` | `Report`/first | 12-month | optional |
| `ORT_MISC` | `ORT` and `Misc` | `Report`/first (stream) | 12-month | optional |
| `BAI2` | `BAI` (incl. `BAIEXP`) | first (`CSVEXP…`) | current cycle | optional (gap-fill) |
| `EDISON_PAY` | `Edison_Payments` | first | — | yes for State |
| `EDISON_INV` | `Edison_Invoices` | first | — | yes for State |
| `MID_MASTER` | `MID_Master` | all sheets | — | yes for merchant |
| `ENRICHED` (Grok cross-ref) | `Enriched` and `CrossRef` | first | — | optional (category hint only) |
| `APPLIED_UNAPPLIED` | `Applied` and `Unapplied` | first | — | yes for SPN corroboration |
| `CONTRACTS_INV` | `Contracts_to_Receivable_Invoices` | first | — | optional (SPN chain) |
| `GMS_AGING` | `GMS_001` or (`Sponsored` and `Aging`) | first | — | optional (SPN chain) |
| `AR_INVOICES` | `AR_Invoices` | first (CSV) | — | optional |
| `AR_MATCHED` | `AR_Matched` or `Deposit_Receipts` | first (CSV) | — | optional (deposit/GL) |
| `AR_UNAPPLIED_SUMMARY` | `AR_063` or (`Unapplied_Receipts_Summary`) | first | — | optional (Review cause) |
| `DEPT_INFO` | `ORT_Department_Info` | first | — | optional (account strings) |
| `GMS_001/002/035` sponsor maps | `RPT_GMS_00` | first (header deep) | — | optional |
| `CFG_MATCHING` | `Matching_Rules` | first | — | optional (rule fidelity) |
| `CFG_PARSE` | `Parse_Rules` | first | — | optional |
| `CFG_TOLERANCE` | `Tolerance_Rules` | first | — | optional |
| `CFG_RULESETS` | `Recon_Rulesets` | first | — | optional |
| `CFG_TCR` | `Transaction_Creation_Rules` | first | — | optional |
| `RELATIONSHIP_MAP` | `Relationship_Map` or `Rosetta` | (reference only) | — | optional |

**Account inference:** take the account token from the BSL/ALL_DATA filename (`FHB_UTC`, `FHB_Master`, `Regions_UTM`, …). Normalize the long Oracle account name to this short form before any scope comparison.

**Multiple files for one role:** if two files map to the same role (e.g., two Receipts files across dates), keep the newest by the leading `YYYYMMDD` in the name; if tied, union the rows and dedup (Section 8).

### 4.2 `ALL_DATA` sheet map (the lifecycle file)
When present, `*_All_Data.xlsx` has these sheets; bind each by title substring (case-insensitive), then bind columns by header (Section 5):
- **`Bank Statement Lines`** — every bank line with `Rec Status`, `Rec Grp`, `Matching Rule Name`, `Match Type`. (`Rec Status` blank/`UNR` = open; `REC` = reconciled.)
- **`MISC Receipts`** — ORT/ECT receipts: `Dep Num`, `Rec Num`, `Rec Amnt`, `Rec Status`, `Rec Ref`, `Rec Grp`, `Recommended Bank Line`, `Offset` (GL).
- **`AR Matched Receipts`** — applied AR receipts with `Strc Pay Ref`, `Accnt Name`, deposit dates.
- **`Recon History`** — one row per reconciliation: `Recon Grp`, `Amount`, `Recon Src`, `Auto Recon Flag`, `Rule Name`, `Match Type`, and the four criteria flags `Type Match / Amount Match / Date Match / Ref Match`. **This is the ground truth for the backward engine.**

---

## 5. Column binding — content first, header as tiebreak (position-independent)

Never hard-code a column index. For every file, define the **roles** the engine needs and resolve each role to a column by scanning the data.

### 5.1 Binding algorithm (implement once, reuse everywhere)

```
def bind_columns(rows, role_specs, header_scan=12):
    # rows: list of tuples (already read). role_specs: {role: RoleSpec}
    # 1. Locate the header row: the first row within header_scan rows whose
    #    non-empty cell count is >= the max required arity and that matches
    #    >=2 header aliases across all roles. Records header_index.
    # 2. For each role, score every column:
    #      header_score  = 3 if normalized header == a header alias
    #                      2 if a header alias is a substring of the header
    #                      0 otherwise
    #      content_score = fraction of sampled data cells (next ~50 rows)
    #                      that satisfy the role's content predicate
    # 3. Choose the column with the highest (content_score, header_score).
    #    - Content decides. Header only breaks ties between content-equal columns.
    #    - If two columns tie on both and the role is required -> AmbiguousColumn.
    #    - If the best content_score is 0 for a REQUIRED role -> InvalidSourceData.
    # 4. Return {role: column_index}, header_index.
```

`RoleSpec` = `(required: bool, header_aliases: [str], content_predicate: fn(cell)->bool)`.
Normalize headers by uppercasing, stripping non-alphanumerics, and collapsing spaces.

### 5.2 Content predicates (the signatures that make binding robust)

- **date** — parses as a date by `parse_date` (Section 7). 
- **signed_amount** — parses via `cents`; allow parentheses and `$`/commas; a column is "amount" if ≥80% of sampled cells parse and at least one is non-integer-dollar or negative.
- **reference/id** — non-empty alphanumeric token, not a pure date, appears with high distinctness.
- **status** — values drawn from a small set like `{REC, UNR, VOID, APP, Cleared, Confirmed, Remitted, Unapplied, Reversed, OP, CL}`.
- **customer/payer** — free-text with letters and spaces, low numeric fraction.
- **transaction_type** — small vocabulary (`ACH, CHK, MSC, EFT, BKA, BKF, ZBA, Credit card, Check, Miscellaneous, …`).
- **GL string** — matches `^\d{2}-\d{7}-\d{6}-\d{6}-…` segmented pattern.
- **MID** — matches `^(80\d{8}|2000\d{6})$` (Section 7).
- **deposit/receipt/batch id** — numeric or `d:`/`r:` token; for MET description see 5.4.

### 5.3 Required roles per file (bind these; raise if a required role is unresolved)

**BSL (anchor)** — required: `date`, `amount(signed)`, `line_key` (statement/line-number identity). Optional but used: `reference`, `account_servicer_reference`, `customer_reference`, `additional_info`, `transaction_type`, `transaction_code`, `addenda`.
> **Reference parsing (critical):** the engine's match key `RECON_REFERENCE` is the **whole Account Servicer Reference** for FHB transaction codes 142/174/175/165/244/495/451/699/475/357/631/661 (Parse Rules pattern `(X~)` = take the entire field). If `account_servicer_reference` is absent, fall back to `reference`. Also parse the ACH payer from addenda `SENDING CO NAME: <X> ENTRY DESC` and from `Company Name: <X>` in an enriched addenda blob. Never tokenize BAI2 field labels (`Company`, `Customer`, `Entry`, `Description`, `Name`, `ID`) as payer text.

**ST (open)** — required: `date`, `amount(signed)`, `transaction_number`, `source` (AR/EXT/AP/PAY/GL). Optional/used: `reference` (= `RECON_MATCH_REFERENCE`), `structured_payment_reference`, `counterparty`, `transaction_type`, `unique_remittance_id`, `bank_deposit_number`, `receipt_batch_number`.
> `Source` is **required** because Journal exclusion is a hard guardrail with no other defense. If `source` cannot be bound → `InvalidSourceData`.

**RECEIPTS (all-status Receivables)** — required: `receipt_number`, `status`, `amount(entered)`, `customer_name`. Used: `receipt_date`, `deposit_date`, `batch_number`, `remittance_batch`, `reference`, `unapplied_amount`. The **SPN** is not a column; extract it from `receipt_number` (Section 7).

**MET** — required: `trx_id` (= ST transaction number), `amount`, `transaction_date`, `status`, `description`. Used: `deposit_id`, `receipt_id`, `cleared_date`, `offset` (GL). Parse the description as `d:{DEPOSIT_ID} | r:{RECEIPT_ID} | {payer/purpose}` (Section 5.4).

**ALL_DATA sheets** — bind by the headers listed in 4.2. On **Recon History**, required: `recon_grp`, `amount`, `auto_flag`, `rule_name`, `match_type`, and the four criteria flags.

**EDISON_PAY** — `reference`, `invoice_number`, `payment_date`, `amount` (amount is the **payment total repeated on every row — never sum it**; take one per reference). **EDISON_INV** — `invoice_number`, `gross_amount`.

**APPLIED_UNAPPLIED** — `trx_number`, `contract_number`, `receipt_number`, `customer_name`, `accounted_applied_amount`, `accounted_unapplied_amount`. **GMS_AGING** — `invoice_number`, `award_number`, `account_name`, `owning_org`, `invoiced_amount`. **AR_MATCHED** — `receipt_number`, `deposit_date`, `batch_id`, `gl_segments`, `bank_account`. **CONTRACTS_INV** — hierarchical: read a contract header (`contract_number`, `contract_name`) then its invoice rows (`invoice_number`).

**MID_MASTER** — scan every sheet for 10-digit tokens matching the MID pattern; exclude Heartland company id `6500000097`. Build `MID -> {campus, dba/department, oracle_dept, gl_codes, deposit_bank}` where those columns bind by header.

### 5.4 MET description parser
Split `CET_DESCRIPTION` on `|`. Segment 0 → `d:` deposit id; segment 1 → `r:` receipt id; segments 2+ joined → payer/purpose. Extract the integer after `d:` and after `r:`. Never use the 18-char truncated "Combined ID" as a join key; join on deposit id and receipt id separately. `cleared_date` is null on UNR rows.

---

## 6. The data relationship map (how every file ties to every other)

The engine is a graph walk anchored on the BSL. These are the load-bearing joins (see the two Rosetta Stone workbooks for the full field dictionary).

**Reference spine (the workhorse):**
`BSL.RECON_REFERENCE (= whole Account Servicer Reference)  ==  ST.RECON_MATCH_REFERENCE / ORT receipt Rec Ref / Receipt reference`.
Group every open ST/receipt sharing that reference; the deduped group sums to the BSL amount. This one relationship makes the large majority of reconciliations (Recon History: Amount+Ref matched 19,732 of 21,942).

**ORT / ECT chain (Rosetta Part 1):**
`MET.CET_TRANSACTION_ID == ST.Transaction Number` (1:1 bridge). `MET.d:` groups many `MET.r:` receipts that sum to one BSL (1:M). ORT AR/Misc files give the receipt inventory and clearing dates; a parked receipt id (`r:`) absent from MET is a `DATA_FEED_ERROR`.

**SPN / Receivables chain (Rosetta Part 2):**
`SPN` is embedded in `Receipt Number` and `ST Transaction Number` (regex `SPN\d+`). `Applied/Unapplied.Receipt Number → TRX Number → Contract Number`. `GMS.Invoice Number → Award Number → SPN customer / Owning Org`. `AR Matched.Receipt Number → deposit date, batch, GL segments`. The receipt reconciles to the BSL on Amount + Reference like any ST; the SPN chain supplies **corroboration** (SPN token, award/contract, customer) and the **GL account string**.

**State / Edison chain:**
`BSL(STATE-TN).Reference (Edison ref) → Edison_Payments → invoice set → Edison_Invoices.gross`. The invoice grosses sum to the State BSL (42 of 45 lines on the worked account). Each invoice maps to a Receivables receipt (its reference contains the invoice number). Corroboration for a State Match is the Edison reference tie; an Edison sum proves the payment exists, never which receipt records it.

**Merchant / MID chain:**
`BSL.reference` (or addenda `Customer ID`) is a `MID`; the merchant card receipts carry the same MID as their reference and group to the settlement. `MID_MASTER` maps the MID to campus/department/GL for the account string. Card window: ST precedes BSL by 1–4 days.

**Account-string sourcing for Review/ECT:**
`MET.OFFSET_CONCATENATED_SEGMENTS`, `AR Matched.CONCATENATED_SEGMENTS`, `MISC Receipts.Offset`, `MID_MASTER.gl_codes`, and `GMS.Owning Org` each yield the GL string to recommend for a manual ECT.

---

## 7. Normalization primitives (implement first; unit-test each)

```
N(v)            -> "" if v is None else str(v).strip()
cents(v)        -> signed integer cents via Decimal; handle "$", commas, (paren)=negative; None if unparseable
parse_date(v)   -> datetime.date; accept date, datetime (check datetime FIRST), "YYYY-MM-DD", "MM/DD/YYYY",
                   "MM/DD/YY", Excel serials, and ISO "….000+00:00"; None if unparseable
znorm(s)        -> uppercase, strip non-alphanumerics  (for reference equality)
digit_runs(s,n) -> set of maximal digit substrings of length >= n (default 5)  (for reference-tie)
payer_tokens(s) -> set of alpha tokens length>=4, minus a GENERIC stopword set that INCLUDES the BAI2
                   labels COMPANY/CUSTOMER/ENTRY/DESCRIPTION/NAME/ID and channel words
                   (ACH/WIRE/DEPOSIT/PAYMENT/MERCHANT/SERVICE/STATE/TENNESSEE/UNIVERSITY/…)
spn_of(s)       -> the SPN token (regex (?i)SPN\s*\d+) found in a receipt/transaction number, else ""
is_mid(s)       -> bool: znorm digits match ^(80\d{8}|2000\d{6})$ and != Heartland 6500000097
sibling(a,b)    -> True if a,b are equal-length pure-numeric of length>=7 differing only in the last 1-2 digits
signed_lag(bd,sd) -> (bd - sd).days   (BSL minus ST)
```

**Reference equality** uses `znorm` equality OR full containment of one `znorm` token inside the other for length ≥ 6. **Reference tie** (weaker, for corroboration) uses a shared `digit_run` of length ≥ 5, truncation allowed. A `sibling` pair is a **conflict**, never a tie.

---

## 8. Build the candidate ST pool (all-status, deduped) — the "different data"

The pool is the union of everything a bank line may legitimately match, each entry tagged with availability.

1. **Open non-Receivables STs** from `ST`: keep `EXT` (ECT), `AP`, and (guardrail) exclude `Journal/GL` from match eligibility but keep for the backward engine. Dedup by transaction number keeping the **largest-magnitude** amount (the receipt/transaction total, not an invoice split); when the kept total row lacks a counterparty, borrow it from a dropped split. Mark `available = True`.
2. **All-status Receivables receipts** from `RECEIPTS` (a Receivable is just an ST): one entry per `receipt_number`; `available = status in {UNR, Remitted, Confirmed, Unapplied}` (open) and `False` for `{Cleared, APP, Reversed, VOID}` (reconciled/closed). **Do not also add Receivables rows from the ST export** — the receipts file is authoritative; adding both double-counts and creates false competition.
3. **ORT receipts** from MET/ORT for the ECT chain: index by `d:` deposit and by `RECON_MATCH_REFERENCE`; carry status from MET (`cleared_date` present ⇒ closed).
4. **Availability is dynamic.** A consumption ledger (a set of consumed ST ids) is checked and updated inside each pass. Never precompute a static "available" list.

Each pool entry: `{id, amount_cents, date, reference, znref, digits, payer_tokens, counterparty, source, status, available, spn, is_mid}`.

---

## 9. FORWARD reconciliation pipeline (fixed order; a later pass never overrides an earlier)

Process the account's open BSLs (BSL file, or ALL_DATA Bank Statement Lines where `Rec Status` is blank/UNR). For each BSL compute: `amount_cents`, `date`, `RECON_REFERENCE` (Section 5.3), `ref_digits`, `payer_tokens`, `lane` (P1), `mid` (if the reference is a MID).

**P0 — Load & validate.** Run the router (Section 4) and binder (Section 5). Assert every required role bound. Build the pool (Section 8). If any required file/role is missing → raise; do not proceed.

**P1 — Classify each BSL into a lane** (routing selector; deterministic, first match wins):
- `STATE` if Additional Information starts with `STATE-TN` or contains `STATE OF TENN`, or the reference is an Edison payment reference.
- `MERCHANT` if the reference (or addenda Customer ID) `is_mid`, or the text contains `MERCHANT SERVICE / BANKCARD / TOUCHNET / CYBERSOURCE / PAYMENTECH`.
- else `GENERAL`.
Classify STATE before MERCHANT before GENERAL. (An enriched-file `Transaction_Category`, if present and populated, may inform the lane but never overrides these deterministic tests.)

**P2 — DATA_FEED_ERROR sweep.** For the account, compare open ORT parked receipt ids (`r:`) against MET receipt ids; any open `r:` absent from MET is flagged in Review Notes with its `d:`/`r:` coordinates. Never silently use or discard it.

**P3 — Exact reference 1:1 (Amount + Reference).** For each unplaced BSL, find open pool entries where `amount == BSL.amount` and `reference_equal(BSL.RECON_REFERENCE, entry)` and `|signed_lag| <= 15`. If exactly one and its source is legal (not Journal) and it is not a MID-into-non-merchant conflict → **Match** (confidence High). If more than one such entry → **Candidate** (`MULTIPLE_EQUAL_CANDIDATES`). Consume on Match.

**P4 — 1:M ECT / ORT reference group (the workhorse).** For each unplaced non-STATE BSL:
- Assemble the group of **all-status, deduped** pool receipts sharing the BSL's reference (via `RECON_MATCH_REFERENCE` / MET `d:` whose members carry that reference / ORT Rec Ref).
- If the deduped group sums exactly to the BSL amount and at least one member is date-plausible (±15) and the reference tie holds → **Match** if the group is unique and fully open; **Candidate** if some members are already closed (this is the auto-rec-stranded case: name the closed members and hand to the backward engine for an unwind) or if a competing equal-sum group exists.
- Reference is the corroboration; **payer is not required here** (Recon History: the 1:M ECT rule matches on Amount + Reference).

**P5 — STATE / Edison bundle.** For each STATE BSL:
- Resolve the Edison reference → invoice set (`EDISON_PAY`) → invoice grosses (`EDISON_INV`). If the invoice grosses sum to the BSL amount, the bundle is confirmed.
- Map each invoice to its Receivables receipt (receipt reference contains the invoice number, or the receipt total equals the invoice/bundle). The deduped receipt group (or the single receipt-total) is the ST side.
- **Match** when the bundle sums and the Edison reference ties a receipt; **Candidate** when the bundle sums but no receipt reference ties (`INCOMPLETE_REFERENCE_SUPPORT`); **Review** when Edison shows no bundle. State date rule: no ceiling when the BSL precedes the ST; demote if the ST precedes the BSL by > 20 days. Never match a STATE line through the ORT chain, and never to a card-batch receipt.

**P6 — Merchant / MID.** For each MERCHANT BSL:
- Group same-MID open card receipts in the 1–4-day card window (ST precedes BSL). If they sum to the settlement and the MID matches on both sides → **Match**. Competing groups or a gap outside the window → **Candidate**. A stale (>30-day) same-MID group → **Review** with an auto-rec note. A merchant BSL with no 1:1 card receipt defers to the reference/chain lane (P4), never straight to Review. Use `MID_MASTER` for the account string.

**P7 — Receivables SPN group.** For each unplaced non-MERCHANT BSL:
- Group open receipts by SPN root (strip only a trailing `-N` sequence, never a space-separated reference) or by shared structured payment reference. A multi-member group sums only when its members share a common corroborator (same SPN, same award/contract via Applied-Unapplied, or the reference tie to the BSL). Sum → **Match** (corroborated) or **Candidate**. Never blind subset-sum; never sum a shared cashier-batch root spanning unrelated payers.

**P8 — Named-payer rules (Amount + Payer).** For the specific configured payers (TVA, US Dept of Energy/ASAP, NSF, UT-Battelle, Jefferson Science, Princeton Plasma, US Fish & Wildlife, American Heart, State of Tennessee AP): if the BSL addenda contains the rule keyword AND an open ST's counterparty matches the rule's counterparty AND the amount is exact and date within ±15 → **Match**. These are the only lanes where payer, not reference, is the corroborator.

**P9 — Payables debit.** A negative BSL vs a single open Payables ST of the same amount, reference-tied → Match; else Candidate.

**P10 — Residual → Review.** Every still-unplaced BSL routes to Review with:
- the dominant cause named (RECEIPT_ENTRY_BACKLOG, NO_MATCH_FOUND, MID_GUARDRAIL, GROUPING_CONFLICT, DATA_FEED_ERROR, DUAL_FIRE, UNSUPPORTED_TRANSACTION, DATE_CONFLICT),
- any rejected chain disclosed with its `d:`/`r:` coordinates (never silently dropped),
- and the **recommended GL account string** for a manual ECT (from Section 6 account-string sources).

**Consumption & conservation.** Each Match/Candidate consumes its STs in the ledger. At the end assert: every BSL placed exactly once; no ST id consumed twice; Candidates never promote a barred ST.

---

## 10. BACKWARD un-reconciliation engine (catch mistakes and fix them)

This module re-audits reconciliations that already happened and recommends unwinding the unsound ones. It runs on the account's **reconciled groups** using `ALL_DATA` (Bank Statement Lines REC rows + MISC/AR receipts + **Recon History**) and, where available, the reconciled ST/BSL exports. It is deterministic and evidence-based; it never unwinds on suspicion alone, only on a doctrine violation it can demonstrate.

### 10.1 Assemble each reconciled group
For every `Recon Grp`: gather the bank line(s) (Bank Statement Lines where `Rec Grp` = group), the member receipts/STs (MISC + AR + ST where `Rec Grp` = group), and the history row (rule, auto flag, match type, four criteria flags). **Dedup members by receipt/transaction number keeping the total** before any arithmetic.

### 10.2 Re-verification checks (each yields a defect code if it fails)
Run all checks on every group; a group may carry several defects.
1. **AMOUNT_INTEGRITY** — the deduped member amounts must sum to the bank line amount on exact signed cents. If not → defect (the classic one-to-many mis-pairing; Oracle auto-rec cannot distinguish equal amounts meeting config).
2. **REFERENCE_INTEGRITY** — for a group whose history claims `Ref Match = Y`, the members must actually share the bank line's `RECON_REFERENCE` (Section 6 spine). If the claimed reference does not tie → defect.
3. **DATE_PLAUSIBILITY** — no member may sit outside a defensible window: a receipt dated 31+ days **after** the bank line, or a non-State receipt 31+ days **before**, is implausible (money reaches the bank within days). Flag with the signed lag. State entry-lag is exempt when the BSL precedes the ST.
4. **ONE_TO_MANY_INTEGRITY** — in a 1:M group, every member must share the grouping key (reference / deposit / SPN root). A member that does not belong (different reference, foreign SPN, different payer with no shared key) is a mis-pulled receipt → defect, and it is exactly the receipt that stranded some other bank line.
5. **SOURCE_LEGALITY** — a Journal-source ST reconciled to a bank line → defect. A card-batch (MID) receipt reconciled to a non-merchant line → defect.
6. **DUAL_FIRE** — one receipt id feeding two open external STs that both cleared → defect.
7. **STATUS_COHERENCE** — a member marked reconciled to this group while also open/unapplied elsewhere, or a VOID/Reversed receipt inside a live group → defect.
8. **CROSS_ACCOUNT_FLAG** — a member whose native bank account differs from the group's account. **Flag only** (advisory); the cross-account misdirection search is a separate project. Never open a cross-account lane here.

### 10.3 Corroborated re-match (the "fix")
For each group failing AMOUNT/REFERENCE/ONE_TO_MANY, run the **forward** engine (Section 9) over the group's bank line against the all-status pool to compute what the *correct* group should be. If a clean, corroborated, exact-sum alternative exists, attach it as the **recommended re-reconciliation**; otherwise recommend unwind-to-open so the item returns to the forward queue.

### 10.4 Priority & confidence
Rank recommendations by severity: AMOUNT_INTEGRITY and SOURCE_LEGALITY (High) → ONE_TO_MANY / REFERENCE (Medium) → DATE_PLAUSIBILITY (Low). Auto-reconciled groups (`Auto Recon Flag = Y`) with a failed criterion are the highest-yield targets (they are where Oracle's amount-blind one-to-many matching erred).

### 10.5 Absolute guardrails for the backward engine
- Never recommend an unwind without a demonstrated doctrine violation and the evidence (the failing check, the numbers, the member ids).
- Never modify Oracle; the engine only **recommends**. The operator posts the unwind.
- Deterministic and idempotent: re-running on the same data yields the same recommendations.

### 10.6 Un-reconciliation output
Workbook tab **Unwind Recommendations**, columns:
`Recon Grp · BSL Date · BSL Line Info · BSL Amount · Reconciled ST(s) · Defect Code(s) · Signed Lag · Rule Name (as reconciled) · Recommended Action · Recommended Re-Reconciliation · Explanation`.
Same formatting standard as Section 13 (dark-red headers mark this as a forensic/backward deliverable).

---

## 11. Date doctrine

Signed lag = BSL date − ST date. Bands: **Strong 0–3**, **Moderate 4–7**, **Weak 8–15**, **Suspicious 16+** for non-State; the hard tolerance ceiling is **±15** (Oracle config). A Match requires the tie within ±15 (Strong/Moderate/Weak) with corroboration; 16–30 stays Candidate only with strong reference corroboration; 31+ is rejected forward (named in Review) and flagged backward. **State exception:** no ceiling when the BSL precedes the ST (receipt-entry lag of months is routine); demote to Candidate only if the ST precedes the BSL by more than 20 days. Date alone is never sufficient evidence and never the sole corroborator.

---

## 12. Classification, confidence, and exception codes

- **Match** — exact amount + reference (or, for named rules, payer) + date within band + unique + fully open. Confidence High.
- **Candidate** — exact amount with weaker/ambiguous support: date in the outer band, competing groups, only-closed members (auto-rec-stranded), or corroboration limited to amount+date. Confidence Medium/Low.
- **Review** — no exact-amount counterpart, or every candidate barred by a guardrail.
Every non-Match row carries ≥1 exception code from: `AMBIGUOUS_GROUP, MISSING_MET_LINK, MISSING_REFERENCE, DATE_CONFLICT, PARTIAL_CHAIN, NO_MATCH_FOUND, MULTIPLE_EQUAL_CANDIDATES, DUPLICATE_ST_ATTEMPT, POSSIBLE_AUTO_REC_SPLIT, UNSUPPORTED_TRANSACTION, WEAK_DATE_SUPPORT, INCOMPLETE_REFERENCE_SUPPORT, GROUPING_CONFLICT, INVALID_SOURCE_DATA, RECEIPT_ENTRY_BACKLOG, DUAL_FIRE, DATA_FEED_ERROR, MID_GUARDRAIL, STATE_LANE_ISOLATION`. Backward defects use the codes in Section 10.2.

---

## 13. Output workbook standard (binding)

One reconciliation workbook per account. Tabs: **Matches**, **Candidate Matches**, **Review Notes**. Nine columns:
`BSL Date · BSL Line Info · BSL Amount · ST Date(s) · ST Number(s) · Confidence · ORT d: · ORT r: · Explanation`.
Formatting: Carlito 11pt; header fill **navy FF1F4E78** with white bold text for the forward reconciliation workbook (dark-red **FF7A1F1F** for the backward/forensic workbook and any reference map); title row 1; row 2 blank; headers row 3; data row 4+; freeze panes at A4; dates as `yyyy-mm-dd` text; USD negatives in parentheses; **static values only — zero formula cells** (prepend a space to any value starting with `=`); multi-value cells use comma-space (semicolon-space when a value embeds a thousands-separator comma), dates in the same order as ST numbers. **Every source BSL appears exactly once across the three tabs. No provenance line, band, comment, or extra row anywhere.**

---

## 14. Independent audit (gates every delivery)

A separate module that **imports nothing** from the engine and re-parses every raw source with its own binder (Section 5). It re-reads the produced workbook and enforces:
`C1` conservation (source BSL count == output count; multiset on date+signed cents); `C2` Match signed-cent equality (deduped ST group sums to BSL); `C3` ST non-reuse across Matches+Candidates; `C4` MET bridge validity (every cited `d:`/`r:` is a real MET pair); `C5` dual-fire excluded from Matches; `C6` STATE lane isolation (no ORT citation on a State Match); `C7` MID guardrail; `C8` date ceiling with the State carve-out (fail a State Match whose cited ST precedes the BSL by >20 days; no ceiling when BSL precedes ST); `C9` formatting (Carlito, correct header fill, freeze A4, zero formulas, tab/row structure); `C10` no provenance content. All checks must pass. On failure, fix the engine, rerun, re-audit — never relax a check. Disclose in the methodology prose every defect found and fixed mid-run. The audit must **independently apply the same keep-largest dedup** so engine and audit agree on the receipt total.

---

## 15. Execution guarantees — keeping it on rails

1. **Single entry point** `run(account_input_dir)` executes: route → bind → validate → build pool → forward P0–P10 → backward 10.x → write workbooks → audit → (only if audit passes) present. Each stage returns a typed result; the next stage asserts its preconditions.
2. **No step is optional or skippable.** If an input for a stage is absent, the stage records a typed "not-run" reason in the run log and the affected BSLs route to Review/Unwind with that reason — the pipeline does not silently continue as if the stage succeeded.
3. **Every row is traced.** Maintain a per-BSL and per-group ledger; assert at the end that counts reconcile (input == Matches + Candidates + Review; every reconciled group either passed all backward checks or appears in Unwind).
4. **Idempotent & deterministic.** Sort every candidate set by a total order (amount, then date, then id) before choosing, so ties resolve identically every run.
5. **Resilient parsing.** Wrap each file load in try/except that raises `InvalidSourceData(file, role, detail)`; never let a bad cell abort the run without a precise diagnosis. Unknown statuses default to "closed/ineligible" (conservative), never to "open".
6. **Self-diagnosis hook.** On any unexpected count (e.g., zero Matches on an account known to have them), the engine runs a diagnostic pass that prints, per lane, how many BSLs had an exact-amount counterpart, how many had a reference tie, and where each was suppressed — so a human can see the funnel without re-instrumenting.
7. **Run log.** Emit a JSON run log: files routed, roles bound (with the winning column per role and every tiebreak), pool sizes by source and status, per-pass placements, backward defects by code, and audit result. This log is the on-rails proof that no step was skipped and no information overlooked.

---

## 16. Build order for the coding agent (do these in sequence)

1. Primitives (Section 7) + unit tests on hand values.
2. Router (Section 4) + binder (Section 5) + `RoleSpec` predicates; test on the provided filenames.
3. Pool builder (Section 8) with dedup; assert no receipt double-counted across ST/RECEIPTS.
4. Forward passes P0–P10 (Section 9), each a pure function over (BSL, pool, ledger).
5. Backward engine (Section 10) over ALL_DATA + Recon History.
6. Workbook writers (Sections 13, 10.6) + audit (Section 14).
7. `run()` orchestrator (Section 15) + JSON run log.
8. Validate against a fully-reconciled account (e.g., FHB UTC): the forward engine must reproduce the known Oracle groups (target ≥ the 4,057 exact-sum groups seen), and the backward engine must surface only genuine defects. Only then run open accounts.

*End of specification.*

---

## Addendum (2026-07-11, answer-key reverse-engineering — owner-approved behavior)

Validated against the owner-supplied FHB Master reference reconciliation
(`FHB_Master_Reconciliation_FY26_Full_v2`). These refinements are binding:

1. **§5 sampling.** Content scoring samples the first 50 **non-blank** values
   per column across the whole sheet, not the first 50 rows. (A Structured
   Payment Reference column blank on every recent Journal row must still be
   scored on the values it carries, or reference-shaped neighbors blind-tie
   it out of binding and silently gut the 4×4 cross-reference screen.)
2. **§5 tie resolution.** A required-role tie among columns whose sampled
   content is value-for-value identical binds the leftmost (verbatim
   duplicate columns are not a real ambiguity). A signed/unsigned amount twin
   (equal absolute cents everywhere; exactly one column carries negatives)
   binds the signed column. Everything else still fails loud.
3. **§8.1 dedup signature.** Keep-largest applies only when the same-key
   group fits the total-plus-splits signature: largest == sum(rest) in signed
   cents. Same-key rows that do not sum that way are **distinct receipts
   sharing a label** — all are kept, ids disambiguated as
   `<id> [<amount>]` (comma-free), `base_id` preserving the MET bridge join
   (which then merges by equal signed cents, never by guess).
4. **§9 ordering.** Out-of-band (stale-ST) exact 1:1 reference ties defer
   from P3 to a post-group pass (P8b): a stale coincidence never pre-empts a
   live ORT deposit-chain, merchant, or SPN group for the same BSL.
5. **§9 P9b amount-only singles.** An unplaced BSL with open exact-cents
   ST(s) surfaces as a LOW Candidate (`AMOUNT_ONLY`, one cited ST, alternates
   named); stale-only exact-cents STs with at least a digit-run tie surface
   as `DATE_GATE` LOW Candidates. Zero-evidence stale coincidences stay
   Review. The hard guardrail (no Match without corroboration; cited STs sum
   exactly) is unchanged.
6. **Payer contradiction (owner directive).** When both sides carry payer
   tokens (BSL side BAI2-enriched; ST side from MET `CET_DESCRIPTION` /
   counterparty) and they share none, a **zero-corroboration** pairing is
   barred even as a Candidate ("City of Chattanooga has nothing to do with
   Israel"). Reference ties always outrank payer text; silence on either
   side never contradicts.

