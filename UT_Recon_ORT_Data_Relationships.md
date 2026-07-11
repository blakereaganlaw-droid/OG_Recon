# UT Reconciliation — ORT / ECT Data Relationships

**Scope.** This document explains, end to end, how the data on the ORT (departmental receipting) side of the University of Tennessee's Oracle Cash Management (DASH) reconciliation connects: which files exist, what each field means, which fields join to which, in what direction, with what transform, and what breaks each join. It is the prose companion to the Data Relationship Map workbook (Part 1) and the ORT/ST relationship diagram.

**Companion document.** `UT_Recon_SPN_Receivables_Data_Relationships.md` covers the Receivables/SPN side. The two sides share one anchor (the bank statement line) and one doctrine; they differ in which files carry the system-transaction detail.

---

## 1. The cast of entities

Six entities participate in every ORT reconciliation. Each lives in its own file.

| # | Entity | Grain (one row =) | Function |
|---|---|---|---|
| 1 | **BSL** — Bank Statement Line | One line on the bank statement | The **anchor**. Every match starts here and runs outward. |
| 2 | **ST** — System Transaction | One Oracle transaction (ECT, Receivable, Payable, Payroll) | The Oracle-side counterpart the BSL must match. |
| 3 | **MET / ORT report** | One external receipt (`r:`) inside one deposit (`d:`) | The **bridge** between ORT activity and Oracle STs. |
| 4 | **Edison** (Payments + Invoices) | One invoice line inside one State of Tennessee payment | State-payment decomposition and corroboration. |
| 5 | **MID Master** | One merchant ID | Merchant authority: MID → campus, DBA, processor, GL. |
| 6 | **CM Configuration** (Parse, Matching, Tolerance, Rulesets, Transaction Creation) | One rule | The **governing layer**: it decides how the fields above are produced and compared. |

**Direction of every read: BSL → ORT → ST.** The bank line is the anchor; the MET report resolves which ORT deposit funds it; the deposit's receipts resolve to STs one for one. No relationship in this document is read backward except in the un-reconciliation engine, which replays the same joins over already-reconciled groups.

---

## 2. Entity by entity: fields and what each one touches

### 2.1 BSL — the anchor

| Field | Oracle name | What it is | What it connects to |
|---|---|---|---|
| BSL Date | `BOOKING_DATE` | The day the money hit the bank | ST Date (date-support bands); MET `CET_CLEARED_DATE` (REC rows agree ±1 day) |
| BSL Amount | statement amount, signed | Exact signed cents | ST Amount; sum(STs); sum(ORT `r:` in one `d:`) |
| BSL Reference | `RECON_REFERENCE` | **The match key.** Produced by Parse Rules from the raw bank fields | ST `RECON_MATCH_REFERENCE`; Edison `Reference`; MID Master `MID` |
| Account Servicer Reference | `ACCNT_SERVICER_REF` | The bank's own reference; populated on every line | Source of `RECON_REFERENCE` for FHB codes 142/174/175/165/244/495/451/475/699/357/631/661 (Parse pattern `(X~)` = take the whole field) |
| Customer Reference | `CUSTOMER_REFERENCE` | Parsed payer name on ACH lines | MET `CET_DESCRIPTION` payer; ST `CPARTY_NAME` |
| Addenda text | `ADDENDA_TXT` | Full ACH addenda blob | Named-payer matching rules (`s.ADDENDA_TXT LIKE '%…%'`); MET `CET_REFERENCE_TEXT`; embedded MID |
| BSL Type | `TRX_TYPE` | Channel code: ACH / CHK / MSC / EFT / BKA / BKF / ZBA | ST Transaction Type (deterministic gate, §5.2) |
| Transaction Code | `TRX CDE` | Bank transaction code (142, 174, 475, …) | Parse Rules (selects the parse pattern); Transaction Creation Rules |
| Additional Information | — | Free text; `STATE-TN` marker lives here | Routing: State lane before merchant before general |
| Bank Account | `CBE_BANK_ACCOUNT_NAME` | Scope key | MET bank account name (long form — normalize before any lookup) |

**Critical production detail.** `RECON_REFERENCE` is not raw bank data; the Parse Rules manufacture it. For First Horizon, the rule for each listed transaction code copies the **whole** Account Servicer Reference into `RECON_REFERENCE`. Any engine that keys on a different field, or that tokenizes the reference, silently loses the workhorse join (§4.1).

### 2.2 ST — the Oracle counterpart

| Field | Oracle name | What it is | What it connects to |
|---|---|---|---|
| ST Date | `TRX_DATE` | Transaction date | BSL Date (bands); card STs normally precede the BSL by 1–4 days |
| ST Amount | signed | Exact signed cents | BSL Amount (1:1 or as a group member) |
| Transaction Number | — | ST identity | MET `CET_TRANSACTION_ID` (**perfect 1:1 bridge**); output "ST Number(s)" |
| ST Reference | `RECON_MATCH_REFERENCE` | **The match key on the ST side** | BSL `RECON_REFERENCE`; MID Master (a MID-valued reference marks a card receipt) |
| Counterparty | `CPARTY_NAME` | Payer name | BSL addenda / customer reference (named-payer rules only) |
| Transaction Type | `TRX_TYPE` | Credit Card / Check / EFT | BSL Type via the deterministic gate (§5.2) |
| Source | `TRX_SRCE` | AR / EXT / AP / PAY / GL | Playbook routing; **Journal (GL) never matches a bank line** |
| Structured Payment Reference | `STRUCTURED_PAYMENT_REFERENCE` | Receivables-side reference | Covered in the companion document |
| Recon status | UNR / REC | Eligibility | Only open STs are available; a REC ST explains a stranded BSL |

**Duplication trap.** ST exports carry a receipt-**total** row and its invoice-application **split** rows under one Transaction Number. Deduplicate by Transaction Number keeping the largest-magnitude amount before any sum; otherwise the engine and the audit disagree by signed cents.

### 2.3 MET / ORT — the bridge

The MET (OTBI all-accounts) report is the only bridge between departmental ORT activity and Oracle STs.

| Field | What it is | What it connects to |
|---|---|---|
| `DEPOSIT_ID` (**ORT d:**) | The deposit group — **parent** | Groups its `r:` receipts; the deposit's deduped receipt sum equals one BSL amount |
| `RECEIPT_ID` (**ORT r:**) | One receipt — **child** | Meaningful only inside its `d:`; joins to ORT parked-receipt detail |
| `CET_TRANSACTION_ID` | The ST Transaction Number | **1:1 bijection to the ST.** The single authoritative ORT→ST link |
| `CET_DESCRIPTION` | Pipe-delimited: `d:{DEPOSIT_ID} \| r:{RECEIPT_ID} \| {payer/purpose}` | Payer corroboration; the `d:`/`r:` coordinates for the output |
| `CET_REFERENCE_TEXT` | Receipt reference text | The BSL addenda originator or merchant ID reappears here (corroboration) |
| `CET_CLEARED_DATE` | Clearing date | **Null on UNR rows.** On REC rows it equals the BSL booking date ±1 day |
| Bank account name | Long form ("FHB - Master Account") | Normalize to the short engine name **before** keying any deposit index |
| Combined ID | 18-character composite | **Never a join key** — it truncates. Join on `DEPOSIT_ID` and `RECEIPT_ID` separately |
| Amount | Receipt amount, signed cents | Member of the deposit sum |
| Status | UNR / REC | REC rows are context: an all-REC deposit explains a stranded BSL |

### 2.4 Edison — the State of Tennessee layer

Two files decompose every State payment.

- **Edison_Payments.** One row per invoice inside a payment. `Reference` (10 digits, leading zeros) repeats on every row of the payment; **`Amount` also repeats the payment total on every row — never sum the Amount column.** Take one total per Reference.
- **Edison_Invoices.** One row per invoice; `Gross` is the per-invoice amount. For a multi-invoice payment, the invoice grosses sum exactly to the bank line.

### 2.5 MID Master — the merchant authority

Scan **every sheet** for tokens matching `^(80\d{8}|2000\d{6})$`; exclude Heartland company ID `6500000097`, which is not a MID. Each MID maps to owning campus, DBA/department, processor, and GL codes. Known file hazards: the Ticketmaster tab has no header row; duplicate `Sheet1`/`Sheet1 (2)` tabs exist; column labels vary between tabs — bind by content, not position.

### 2.6 CM Configuration — the governing layer

| File | Governs | The fact that matters |
|---|---|---|
| Parse Rules | How raw bank fields become `RECON_REFERENCE` / `CUSTOMER_REFERENCE` | FHB copies the **whole** Account Servicer Reference (`(X~)`) for the listed codes |
| Matching Rules | Which fields Oracle compares | The workhorse compares **Amount + Reference**; named-payer rules add `s.ADDENDA_TXT LIKE` paired with `t.CPARTY_NAME` |
| Tolerance Rules | Windows | Date ±15 days; **amount tolerance disabled** |
| Recon Rulesets | Firing order | Sequence number = execution order, most deterministic first |
| Transaction Creation Rules | Auto-creation | 905 rules; an addenda `SRCH STRNG` hit auto-creates an ST and auto-reconciles — the primary BSL-stranding mechanism |

---

## 3. The primary chain, walked once

A departmental deposit reconciles through exactly one path:

```
ORT r: (receipt)  →  Oracle ST  →  ORT d: (deposit)  →  BSL
```

Step by step, with the joins named:

1. **Normalize.** Map the MET long bank-account name to the short engine name. Skip this and every deposit lookup silently returns nothing.
2. **Group.** Index MET rows by `DEPOSIT_ID`. Each `d:` group's rows are its `r:` receipts.
3. **Bridge.** For each `r:` row, `CET_TRANSACTION_ID` names the ST — a perfect 1:1.
4. **Dedup.** Collapse the group by unique `RECEIPT_ID` (a dual-fire receipt spawns two identical STs from one receipt) and by Transaction Number keeping the total row.
5. **Sum.** The deduped receipt amounts must equal the BSL amount on exact signed cents: `sum(ORT r:) = sum(STs) = BSL Amount`.
6. **Corroborate.** An exact-sum group is **necessary, never sufficient**. One of: (a) the group's reference ties the BSL `RECON_REFERENCE`; (b) a MET payer segment matches the BSL payer; (c) the BSL is a bundled DEPOSIT/REMOTE DEPOSIT line, where deposit-type consistency itself satisfies the gate.
7. **Date.** At least one member inside ±15 days (bands in §5.3).
8. **Classify.** Unique, corroborated, fully open, in-window → **Match**. Anything less → **Candidate** or **Review**, with the reason named.

---

## 4. The join crosswalk — every key, transform, and hazard

### 4.1 The reference spine (the workhorse)

**`BSL.RECON_REFERENCE = ST.RECON_MATCH_REFERENCE` (or = the ORT receipt reference).**
Cardinality 1:1 or 1:M. Transform: none beyond trim/case (the Parse Rules already produced the key). This single join, with exact amount, makes the large majority of reconciliations: on the fully reconciled validation account, 19,732 of 21,942 historical reconciliations matched on Amount + Reference with type and date unused, and the 1:M ECT rule alone produced 11,252 groups. **Payer is not part of this join.**

- Group every open receipt sharing the BSL's reference; the deduped group sums to the BSL.
- A partial reference (≥4 shared digits, truncation allowed) is **corroboration only**, never a sole basis.
- **Sibling references** — equal-length numerics of 7+ digits differing only in the final one or two digits — are adjacent deposit slips, i.e., **conflicts**, never partial matches.

### 4.2 The MET bridge

**`MET.CET_TRANSACTION_ID = ST.Transaction Number`** — 1:1 bijection, no transform, 100% coverage on every account tested (13,522 / 1,603 / 436 external STs across three accounts). Never substitute the truncated Combined ID.

### 4.3 The deposit group

**`ORT d:` contains `ORT r:`** — 1:M parent→child. An `r:` receipt has no standalone meaning; it is priced and attributed only inside its deposit. The deduped `r:` amounts in one `d:` sum to one BSL.

### 4.4 The Edison joins (State lane)

1. **`BSL.RECON_REFERENCE = Edison_Payments.Reference`** — M:1; transform: **strip leading zeros** (BSL `7182042` = Edison `0007182042`). Never let a tool float-coerce the reference.
2. **`Edison_Payments.Reference → Edison_Invoices`** — 1:M; the invoice grosses (from the **Invoices** file, never the Payments Amount column) sum to the bank line. On the worked account, 42 of 45 State lines reconstructed exactly (24 single-invoice, 18 multi-invoice).

The Edison bundle confirms the **payment's identity**. It never by itself proves which DASH receipt records the payment — a State Match still requires a reference tie between line and receipt (the receipt side is in the companion document).

### 4.5 The MID joins (merchant lane)

1. **`ST.Reference = MID_Master.MID`** — M:1, exact. A MID-valued reference marks a card/merchant receipt. **Guardrail: a MID receipt never matches a non-merchant BSL,** in any lane, on any coincidence of amount.
2. **`BSL addenda MID → MID_Master → campus/DBA/GL`** — extraction requires a merchant marker (MERCHANT SERVICE / BANKCARD / TOUCHNET / CYBERSOURCE / PAYMENTECH) near the number; a bare 10-digit number without a marker is not a validated MID.
3. **Card window:** the card ST precedes the BSL by 1–4 days. Same MID both sides + exact amount + window = Match. Note that merchant deposits also reconcile through the ordinary reference grouping of §4.1 — the validation account showed same-MID receipt groups summing to their settlement under the shared reference (e.g., group reference `8039416949`, a MID, receipts summing $1,134.00).

### 4.6 The scope join

**`BSL.Bank Account (short) = MET.CBE_BANK_ACCOUNT_NAME (long, normalized)`** — 1:1 after the long→short mapping. Every match is same-account; cross-account amount hits are advisory flags for the separate Misdirection project, never Matches here.

### 4.7 The configuration joins

- `Recon Rulesets.MTCH RLE NME → Matching Rules` and `.TOL RLE NME → Tolerance Rules` — M:1; sequence order = firing order.
- `Parse Rules (bank + code + type) → BSL.RECON_REFERENCE` — governs; upstream of every reference join.
- `Transaction Creation Rules.SRCH STRNG → BSL.ADDENDA_TXT → GL CASH/OFFSET` — governs; the auto-creation trigger and the primary stranding defect.

---

## 5. Deterministic gates

### 5.1 Lane precedence (classify in this order)

1. **STATE** — Additional Information starts `STATE-TN` / contains `STATE OF TENN`, or the reference is an Edison payment reference. State lines match Receivables receipts only; the ORT chain never applies.
2. **MERCHANT** — the reference or addenda carries a validated MID or a merchant marker.
3. **GENERAL** — everything else runs the ORT chain.

### 5.2 Type gates (hard rejects)

| ST type | Pairs with BSL | Never pairs with |
|---|---|---|
| Credit Card | ACH ("merchant services") | Check, Miscellaneous |
| Check | Miscellaneous | — |
| EFT | ACH (EFT is often miscategorized — do not reject on the label alone) | — |

### 5.3 Date bands

Signed lag = BSL date − ST date. **Strong 0–3 · Moderate 4–7 · Weak 8–15** (the Oracle tolerance ceiling) — a Match needs the tie inside ±15 with corroboration. 16–30 days survives only as a Candidate with strong reference corroboration. 31+ days: a BSL trailing its ST by 31+ is a stale-ST pairing (money reaches the bank within days) — rejected everywhere; a BSL preceding its ST by 31+ is receipt-entry lag — plausible but Review-only pending manual corroboration. On the ORT clearing side, `CET_CLEARED_DATE` on REC rows agrees with the BSL booking date within ±1 day. Date alone is never sufficient evidence.

---

## 6. Known defects and watchpoints (verified)

1. **Auto-creation strands BSLs.** 905 Transaction Creation Rules; an addenda hit auto-creates and auto-reconciles, leaving the true BSL stranded or a deposit cherry-pick-split.
2. **Auto-rec ignores date.** Oracle matches on amount + reference with `DTE MTCH = N`; equal amounts on different days mis-pair. This is the operative defect the backward engine hunts.
3. **Dual-fire receipts.** One `RECEIPT_ID` spawning two identical open external STs (132 on FHB Master). Dedup by `RECEIPT_ID` before summing; never Match a dual-fire member.
4. **Combined-ID truncation** (§2.3). 5. **Long vs short account names** (§4.6). 6. **`CET_CLEARED_DATE` null on UNR** (§2.3). 7. **Edison Amount repeats** (§2.4). 8. **MID Master header chaos** (§2.5). 9. **Leading-zero corruption** — read every reference with openpyxl as text; pandas float-coerces `0006789599` into `6789599.0`. 10. **Two active-looking rulesets** — "UT New Recon Rule Set" (2026-06-02, AOWENS43) governs; the 2025 JMCGILL1 set is legacy.

---

## 7. Worked example

A GENERAL-lane BSL: 2026-07-02, $1,134.00, reference `8039416949`.

1. Lane: the reference is a MID (`80…`, 10 digits) → MERCHANT.
2. MET: deposit `d:1048221` carries three `r:` receipts whose `CET_REFERENCE_TEXT` is `8039416949`; deduped amounts $500.00 + $400.00 + $234.00 = $1,134.00.
3. Bridge: each `r:` row's `CET_TRANSACTION_ID` names an open ST.
4. Corroboration: the shared reference **is** the BSL reference — gate (b) satisfied; the MID matches both sides.
5. Dates: receipts 2026-06-30, settlement 2026-07-02 — inside the 1–4-day card window.
6. Classification: unique, corroborated, fully open, in-window → **Match**, citing `d:1048221`, the three `r:` values, and the three ST numbers.

If one of the three receipts had already been REC (auto-rec split), the open members would sum short; the line becomes a **Candidate** naming the REC member — and the backward engine receives the group as an unwind lead.
