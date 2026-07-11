# UT Reconciliation — SPN / Receivables Data Relationships

**Scope.** This document explains, end to end, how the data on the Receivables and sponsored-projects (SPN) side of the University of Tennessee's Oracle Cash Management (DASH) reconciliation connects: which files exist, what each field means, which fields join to which, with what transform, and what breaks each join. It is the prose companion to the SPN/Receivables Rosetta Stone (Part 2) workbook and the SPN/Receivables relationship diagram.

**Companion document.** `UT_Recon_ORT_Data_Relationships.md` covers the ORT/ECT side. Both sides share one anchor (the bank statement line) and one doctrine.

**The one-sentence model.** A Receivables receipt is just an ST: it reconciles to a bank line on **Amount + Reference**, exactly as an ECT does; everything else in this document — invoice, contract, award, SPN — exists to **corroborate** that pairing and to supply the **GL account string** when no pairing exists.

---

## 1. The SPN — the anchor concept

**SPN = Sponsored Projects Number**, the identifier UT assigns to a sponsored-projects award.

- The SPN is **not stored in a dedicated column anywhere**. It is embedded in the **Receipt Number** and in the **ST Transaction Number**: `MEP SPN110364 2025-3729`, `SPN 051826 Inv 110011427`, `18SPN062526 7318069`. Extract it with the regex `(?i)SPN\s*\d+`.
- Coverage (verified on the all-status receipts file): of **11,723** receipts, **2,926** carry an SPN token and **3,392** carry a pure-numeric Edison reference (State receipts). The remainder are conversion, campus-batch, or miscellaneous formats (`UTK_CONV_1719`, `UTHSC_8900P_OCT25-221`, plain numerics like `193200`).
- **Why it matters.** The bank never sees an SPN — the bank side carries only dollars, a date, a reference, and sometimes a payer. The SPN chain (award → contract → invoice → receipt) is what turns an exact-amount candidate into an audit-defensible Match and what names the GL string for a manual ECT when no receipt exists.

---

## 2. The cast of entities

| # | Entity | File | Grain (one row =) | Function |
|---|---|---|---|---|
| 1 | **Receivables Receipts (all-status)** | `20260709_Receivables_Receipts_All.xlsx` | One receipt, any status | **The ST for the AR side** — the money that hit the bank. 11,723 receipts. |
| 2 | **Applied/Unapplied Payment Report** | `Applied_-_Unapplied_Payment_-_Report…xlsx` | One invoice application of one receipt | **The bridge**: receipt → invoice → contract, with applied/unapplied amounts. 16,792 rows. |
| 3 | **Contracts → Receivable Invoices (AR-GMS)** | `Contracts_to_Receivable_Invoices…xlsx` | One invoice under a contract | Contract header (number + name carrying the project/fund code) then its invoices. **Hierarchical layout.** |
| 4 | **GMS 001 Sponsored AR Aging** | `RPT_GMS_001…RPT1.xlsx` | One sponsored invoice with aging | Invoice → **Award Number** → SPN customer/account → **Owning Org** → invoiced amount. 2,264 sponsored invoices; 1,346 awards. **Sponsored only.** |
| 5 | **AR Invoices** | `AR_Invoices_AR_Invoices.csv` | One AR invoice header | Invoice master: customer, department, currency, old TRX number. Resolves non-sponsored invoices. |
| 6 | **AR Matched Invoice Receipts (Deposit Receipts, Non-Misc)** | `AR_Matched…csv` | One matched deposit receipt | Receipt → deposit date, batch ID, **GL string (`CONCATENATED_SEGMENTS`)**, bank account. 4,337 rows. |
| 7 | **RPT AR 063 Unapplied Receipts Summary** | `RPT_AR_063…RPT63.xlsx` | One unapplied receipt | Cash received but not applied — a distinct Review cause. |
| 8 | **BSL / Edison / configuration** | Part 1 files | — | The anchor, the State decomposition, and the governing layer (see the ORT document). |

---

## 3. Entity by entity: fields and what each one touches

### 3.1 Receivables Receipts (all-status) — the ST

| Field | What it is | What it connects to | Contingency |
|---|---|---|---|
| **Receipt Number** | Identity **and SPN/Edison carrier** | SPN token (embedded); Edison reference (numeric formats); Applied/Unapplied `Receipt Number`; AR Matched `ACRA_RECEIPT_NUMBER` | Formats vary widely; the SPN token appears only on sponsored receipts |
| **Status / State** | Lifecycle | Match eligibility | `Cleared`/`Confirmed` = applied or reconciled (closed); `Remitted` = sent to bank, open in CM; `Unapplied` = cash not yet applied (open); `Reversed` = void. **A Cleared receipt matching an open bank line signals an auto-rec error to unwind.** Unknown statuses default to closed (conservative). |
| **Customer Name** | Payer | BSL payer; GMS `Account Name (… SPN)`; ST `CPARTY_NAME` | Corroboration for named-payer rules; **secondary to Reference** for the workhorse rules |
| **Entered Amount** | Exact signed cents | BSL Amount (1:1 or as group member) | Dedup receipt-total vs invoice-application splits first |
| **Deposit Date / Receipt Date** | Date support | BSL Date, ±15-day window | Date rarely gates the AR 1:M rule (`Date Match = N` in 19,732 of 21,942 historical reconciliations); on State lines months of entry lag are routine |
| **Reference / Structured Payment Reference** | **The match key** | BSL `RECON_REFERENCE`; Edison invoice number (State receipts) | The receipts-file reference column (col 35 in the verified export) carries the bank-facing reference; the SPR groups a sponsored deposit |
| **Batch / Remittance Batch** | Grouping (weak) | Physical deposit | Frequently empty — corroboration, never the primary group key |
| **Unapplied Amount** | Residual cash | RPT AR 063 | Distinguishes "receipt exists, unapplied" from "no receipt entered" |

### 3.2 Applied/Unapplied Payment Report — the bridge

| Field | What it connects to | How |
|---|---|---|
| **Receipt Number** | Receivables Receipts | 1:1 to the receipt; 1:M to its invoice applications |
| **TRX Number** | AR Invoices; GMS 001 `Invoice Number`; Contracts file | The invoice the receipt paid |
| **Contract Number** | Contracts file; GMS Award | The sponsored contract; occasionally echoed inside the Receipt Number (`REF:261; 2003296`) |
| **Accounted Applied Amount** | Receipt Entered Amount | Sum over the receipt = the applied portion |
| **Accounted Unapplied Amount** | RPT AR 063 | The unapplied residue; explains a receipt that only partly clears a bank line |

### 3.3 Contracts → Receivable Invoices — the project layer

Hierarchical layout: a **contract header row** (`Contract Number`, `Contract Name`) followed by its invoice rows. The parser must read the header, then attribute the following invoice rows to it until the next header. **`Contract Name` carries the project/fund code** — `701351-Plant Funds EG`, `180008-Center of Farm Management` — the same code family that appears in GMS `Owning Org` and in GL segments.

### 3.4 GMS 001 Sponsored AR Aging — the award layer

| Field | Role |
|---|---|
| **Invoice Number** | Join from Applied/Unapplied `TRX Number` |
| **Award Number** | **The SPN grouping root.** 1,346 distinct awards over 2,264 sponsored invoices |
| **Account Name** | The sponsored customer, suffixed `SPN` (`1890 Universities Foundation SPN`, `AAA Foundation for Traffic Saf…`) — payer corroboration |
| **Owning Org** | `180008-Center of Farm Ma…` — drives the GL segment recommendation |
| **Invoiced Amount** | Gross per invoice — decomposes a multi-invoice payment; compare to Applied for partials |

Header begins at row 14 in the verified export. **Sponsored invoices only** — a TRX absent here is non-sponsored and resolves through AR Invoices.

### 3.5 AR Matched Invoice Receipts — the deposit/GL layer

| Field | Role |
|---|---|
| **ACRA_RECEIPT_NUMBER** | Join to Receivables Receipts (partial — see §5.3) |
| **ACRA_DEPOSIT_DATE / ABA_BATCH_ID** | The physical bank deposit the receipts rolled into |
| **CONCATENATED_SEGMENTS** | **The GL account string** — the recommendation for a manual ECT on a Review-only line |
| **Bank Account** | Scope |

### 3.6 RPT AR 063 — the unapplied ledger

One row per receipt whose cash sits unapplied. In Review Notes this distinguishes two different departmental failures: **the receipt exists but no one applied it** (RPT 063 hit) versus **no receipt was ever entered** (no hit anywhere) — different fix, different owner.

---

## 4. The hierarchy, top to bottom

```
SPN AWARD  (GMS Award Number — the SPN root)
  └── CONTRACT  (Contract Number; Contract Name = project/fund code)
        └── INVOICE  (TRX / Invoice Number; gross in GMS 001 / AR Invoices)
              └── RECEIPT  (Receipt Number — THE ST; embeds SPN or Edison ref)
                    └── DEPOSIT  (AR Matched batch + deposit date)
                          └── BANK LINE  (the anchor)
```

A child has no standalone meaning: an invoice application is a slice of a receipt; a receipt is a slice of a deposit; the deposit is what the bank line records. Reconciliation walks this tree **bottom-up from the bank line**; corroboration walks it **top-down from the award**.

---

## 5. The join crosswalk — every key, transform, and hazard

### 5.1 The reconciliation join (the only one that makes a Match)

**`BSL.RECON_REFERENCE = Receipt.Reference / Structured Payment Reference`**, with **`sum(deduped receipts) = BSL.Amount`** on exact signed cents.
Cardinality 1:1 or 1:M. Transform: none (the Parse Rules already produced `RECON_REFERENCE` from the whole Account Servicer Reference). Confirmed by the CM Matching Rules (1:1 SPR rule; 1:M Receivables rule) and by the validation account's history (Amount + Reference = 19,732 of 21,942 reconciliations). The bank never sees the SPN; **it sees the reference**.

### 5.2 The SPN extraction (embedded, not joined)

**`Receipt Number → SPN token`** — regex `(?i)SPN\s*\d+`. 2,926 of 11,723 receipts. Sponsored receipts only; a State receipt instead carries a pure-numeric Edison reference. The **SPN root** for grouping strips only a trailing `-N` sequence from the transaction number — never a space-separated reference (stripping the reference destroys the group key).

### 5.3 The receipt-number join (two numbering systems)

**`Receipt.Receipt Number = Applied/Unapplied.Receipt Number`** — 1:M, exact string, but only **2,683 direct overlaps** out of 7,185/10,806 populations. Two numbering systems coexist (`SPN 04292025_ SPN107635 _ 250` vs `193200`). **Rule: join by receipt number where formats align; otherwise route through the invoice** (`TRX Number`), which both systems share. The same caution applies to `AR Matched.ACRA_RECEIPT_NUMBER`.

### 5.4 The invoice joins

- **`Applied/Unapplied.TRX Number = GMS 001.Invoice Number`** — M:1, exact. Sponsored only; a miss here means non-sponsored, resolve via **`= AR_Invoices.TRX Number`**.
- **`Applied/Unapplied.Contract Number = Contracts file.Contract Number`** — M:1, exact; read the hierarchical header.
- **`GMS 001.Invoice Number → Award Number`** — M:1 within the file; the award is the SPN grouping root.

### 5.5 The State/Edison joins

**`STATE-TN BSL.Reference (strip leading zeros) = Edison_Payments.Reference → invoice set → Edison_Invoices.Gross`** — the grosses sum to the bank line (42 of 45 State lines reconstructed on the worked account: 24 single-invoice, 18 multi-invoice). Each Edison invoice then ties a Receivables receipt: **the receipt's reference or number contains the invoice number** (`SPN 051826 Inv 110011427`). Two rules travel together: **never sum the Edison Payments Amount column** (it repeats the payment total per row — per-invoice grosses live in Edison_Invoices), and **an Edison sum confirms the payment's identity, never which receipt records it** — a State Match still requires the reference tie to a receipt. STATE-TN payments are never card payments; a card-batch receipt never enters the State lane.

### 5.6 The deposit/GL join

**`Receipt.Receipt Number = AR Matched.ACRA_RECEIPT_NUMBER`** — 1:1 where formats align (§5.3). Supplies `ACRA_DEPOSIT_DATE`, `ABA_BATCH_ID`, and `CONCATENATED_SEGMENTS`. The GL recommendation for a Review-only line comes from here, or from GMS `Owning Org` when only the invoice side is known.

### 5.7 The amount join (doctrine)

**`Receipt.Entered Amount ↔ BSL.Amount`** — exact signed cents, 1:1 or group-sum; Oracle's amount tolerance is disabled. **Dedup first**, twice over: (a) duplicate export rows of the same receipt; (b) the receipt-**total** row vs its invoice-application **split** rows under one number — keep the total, drop the splits.

---

## 6. The reconciliation walk (Receivables playbook)

1. **Classify the line.** STATE-TN → Edison playbook; merchant → card lanes (a card-batch receipt never leaves them); otherwise general.
2. **Parse `RECON_REFERENCE`** (whole Account Servicer Reference for the FHB codes).
3. **Assemble the group** — all-status, deduped receipts sharing the reference (or the SPN root, or the deposit). **All-status is mandatory:** a UNR-only pool sums short whenever auto-rec already cleared part of the deposit, which is exactly the case worth catching.
4. **Sum** to the BSL on exact signed cents.
5. **Corroborate** — SPN token, Award/Contract via Applied/Unapplied, SPN customer name, or the reference tie itself. Amount alone never suffices; a unique exact-amount receipt inside ±15 days with **no** payer or reference tie is a **Candidate**, not a Match.
6. **Classify.** Fully open + unique + corroborated + in-window → **Match**. Open-but-competing, or closed members inside the sum (auto-rec split) → **Candidate**, naming the closed members for the backward engine. No group at all → **Review**, distinguishing *unapplied receipt exists* (RPT 063) from *no receipt entered* (departmental backlog), and recommending the GL string (AR Matched `CONCATENATED_SEGMENTS` or GMS `Owning Org`).

---

## 7. Failure modes and their signatures

| Failure | Signature in the data | Handling |
|---|---|---|
| Departmental receipt-entry backlog | Bank line has money; no receipt in any status carries the amount/reference | Review, `RECEIPT_ENTRY_BACKLOG`; recommend the GL string |
| Receipt exists, unapplied | RPT AR 063 row; Unapplied Amount > 0 | Review, distinct cause; name the receipt |
| Auto-rec split | Group's **all-status** sum equals the BSL but some members are `Cleared` against another line | Candidate + unwind lead for the backward engine |
| Dual numbering miss | Receipt joins fail between Applied/Unapplied and Receipts All | Route the join through `TRX Number` |
| Double-count inflation | Same receipt number appears as total + splits, or duplicated rows | Dedup keeping the largest-magnitude (total) row |
| Non-sponsored TRX in GMS lookup | TRX absent from GMS 001 | Resolve through AR Invoices; do not treat absence as an error |
| Edison Amount summed | State line "matches" a multiple of the true payment | Never sum Payments Amount; per-invoice grosses from Edison_Invoices |
| Card receipt drifting into AR lanes | Receipt number carries a standalone `CC` token, type Credit card, or a MID reference | Merchant lanes only; never STATE, never Receivables 1:1 |

---

## 8. Worked example

STATE-TN BSL: 2026-07-01, $61,754.32, reference `0007182042`, Additional Information `STATE-TN…`.

1. Lane: STATE (marker + Edison-format reference).
2. Edison: strip zeros → `7182042`; Edison_Payments shows one payment under that reference covering three invoices; Edison_Invoices grosses $40,000.00 + $18,254.32 + $3,500.00 = **$61,754.32** — the bundle reconstructs the line exactly.
3. Receipts: the all-status receipts file holds `SPN 062926 Inv 110011427` (and siblings) whose references contain the three invoice numbers; deduped entered amounts sum to $61,754.32.
4. Corroboration: the invoice-number tie between Edison and the receipts **is** the reference corroboration; GMS 001 maps the invoices to Award `2101045`, customer `…Universities Foundation SPN`, Owning Org `180008-…`.
5. Dates: receipts entered 2026-08-14 — 44 days **after** the bank line. State exception: entry lag when the BSL precedes the ST carries no ceiling.
6. Classification: bundle sums + receipt reference tie + open receipts → **Match**. Had Edison confirmed the payment but no receipt carried the invoices, the line would go to **Review**: "Edison confirms the payment; no DASH receipt records it; recommended account string from Owning Org `180008-…`."
