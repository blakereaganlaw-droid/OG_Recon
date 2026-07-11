# Oracle Fusion Cloud Cash Management — Bank Statement Reconciliation Setup: Field-Level Authoring Reference (25D / 26A era)

> **CRITICAL METHODOLOGY & VERIFICATION NOTICE — READ FIRST.**
> During this research session, **all live web-research tools (`web_search`, `web_fetch`) and the delegated research subagent returned hard tool-unavailability errors.** No page on `docs.oracle.com` (the "Using Cash Management" guide, "Implementing Payments and Collections," Oracle Applications Help, FBDI/REST references) could be opened or verified in this session, and **no verbatim Oracle quotes or live source URLs could be captured.**
> Consequently, everything below is reconstructed from established, prior knowledge of Oracle Fusion Cloud Cash Management (broadly the 22x–24x design, which has been stable across releases) and is presented as an **authoring scaffold, NOT as primary-source-verified content.** Field names, allowed values, and especially the parse-expression token grammar, the exact advanced-criteria column names, and the complete `CE_TRX_TYPE` delivered value list **must be confirmed against the live current-release Oracle guides before you rely on them.** Each section carries a confidence flag. I have **not** fabricated citations; where I could not confirm something, I say so explicitly. This transparency is the honest outcome given the tool failure, and I strongly recommend re-running the confirmation step (see Recommendations) against the specific Oracle pages I name.

---

## TL;DR
- **Six setup objects drive automatic bank-statement reconciliation, and they execute in a strict runtime pipeline:** (1) transaction-code mapping normalizes raw bank codes into Oracle transaction types → (2) optional parse rules enrich statement-line fields → (3) the reconciliation rule set (matching rule + tolerance rule, evaluated in sequence per bank account) attempts auto-match → (4) transaction creation rules generate first-notice bank items (fees/interest) from whatever remains. Author them in Setup and Maintenance, **Financials → Cash Management** functional area.
- **The highest-leverage, most error-prone authoring decisions** are: inverting `LIKE`/`NOT LIKE` in advanced criteria; parse-rule literal anchors that don't byte-for-byte match the bank's actual delimiter text (causes *silent* non-extraction); over-wide date tolerance letting amount-coincident items auto-match the wrong pair; **disabling amount tolerance** (which forces exact-amount matching and thereby blocks legitimate net-of-fee matching); and overusing transaction creation rules to paper over an AR/AP entry backlog.
- **Scoping rules you must internalize:** amount and percentage tolerances apply to **one-to-one matching only**; date tolerance applies to **all** match types; grouped match types (one-to-many, many-to-one, many-to-many) rely on **grouping attributes** to sum a set before comparing; and matching rule sets are assigned at the **bank account** level, evaluated top-down by **sequence** — so a broad rule placed too early pre-empts more precise rules below it.

---

## Key Findings

1. **Runtime order is the mental model that makes every field coherent.** Code mapping (primary normalization) → parse rules (optional enrichment, run *after* code mapping) → automatic reconciliation via the rule set → manual reconciliation → transaction creation rules (run *after* both auto and manual). Authoring a field wrong upstream (e.g., a code-map miss) cascades: parse rules keyed on the wrong transaction type never fire, and matching rules that test transaction type never match.

2. **Match types** are One-to-One, One-to-Many, Many-to-One, and Many-to-Many. Grouped types require **grouping/advanced matching attributes** so Oracle can aggregate a set of lines or a set of system transactions and compare the *group total* to the single counterpart. Zero-amount / adjustment handling is generally achieved through tolerance-absorbed residuals and transaction creation rules rather than a distinct "zero-amount match type." *(Flag: the existence/label of any dedicated zero-amount match type is UNCONFIRMED against current docs.)*

3. **Tolerance scoping is asymmetric and consequential:** date tolerance = all match types; amount + percentage = one-to-one only. "Amount tolerance disabled" means **exact-amount matching is required** — it is *not* amount-blind matching. This single misunderstanding is the classic reason fee-net receipts (bank credits statement net of a wire fee) fail to auto-reconcile.

4. **Parse rules are optional enrichment**, layered on top of the normalized transaction type; their entire value is populating statement-line target fields (reconciliation reference, customer reference, structured payment reference, etc.) from free-text source fields (like the addenda text) using an extraction expression grammar.

5. **Reconciliation rule sets are ordered pairings** (matching rule + tolerance rule per sequence line) assigned to a bank account; ordering strategy (narrow-before-broad) materially changes match quality.

6. **Transaction creation rules are a last-resort, bank-originated-items-only tool.** Using them for customer receipts (which should reconcile to AR) or for zero-balance-account sweeps is a documented anti-pattern that masks operational backlogs and corrupts reconciliation integrity.

---

## Details

> Formatting note: each object below gives the **Setup and Maintenance task name**, then a field table with the five columns your template needs — **Field | Allowed values / format | What it does | Common mistake | Downstream effect.** Confidence flags: **[H]** high confidence / stable design, **[M]** medium — plausible but verify, **[U]** unconfirmed / verify verbatim against live docs.

---

### OBJECT 1 — RECONCILIATION MATCHING RULES
**Task:** *Manage Bank Statement Reconciliation Matching Rules* (Setup and Maintenance → Financials → Cash Management). **[M]**

**Header fields**

| Field | Allowed values / format | What it does | Common mistake | Downstream effect |
|---|---|---|---|---|
| Matching Rule Name | Free text, unique | Identifies the rule for reuse across rule sets. **[H]** | Naming by bank instead of by *behavior* (e.g., "Wire 1-1 by Ref") makes reuse impossible | A clear name lets one rule serve many bank accounts; poor naming causes rule sprawl |
| Description | Free text | Documents intent/scope | Left blank | Later maintainers can't tell narrow from broad rules; increases mis-ordering risk |
| Match Type | **One to One, One to Many, Many to One, Many to Many** | Defines cardinality of the match — one bank line to one txn, one line to many txns, many lines to one txn, or many-to-many. **[H]** | Choosing many-to-many "to be safe" — it is the most expensive and most prone to false grouping | Grouped types force grouping attributes and enable amount aggregation; wrong type prevents matches from ever forming |

**Matching criteria toggles** (per rule; each is a compare-or-ignore switch)

| Field | Allowed values / format | What it does | Common mistake | Downstream effect |
|---|---|---|---|---|
| Amount (match on amount) | On/Off | Compares bank line amount to system transaction amount, **evaluated against the paired tolerance rule's amount/percentage tolerance**. **[H]** | Assuming "off" makes matching amount-blind — in practice amount comparison is fundamental; turning it off produces reckless matches | Off + wide date tolerance = amount-coincident false matches |
| Date (match on date) | On/Off | Compares bank line date (booking/value/transaction date) to the system transaction date, **against the date tolerance (days before/after)**. **[H]** | Enabling date match but leaving date tolerance at 0 when the bank books a day later | Legitimate 1-day-lagged items fail to auto-match |
| Reconciliation Reference (match on reference) | On/Off | Compares the reconciliation reference on the bank line to the reference on the system transaction. This is the **highest-confidence single criterion** when the bank returns a usable reference. **[H]** | Enabling it before a parse rule has *populated* the reference field on the line | Rule silently never matches because the line-side reference is null |
| Transaction Type (match on type) | On/Off | Compares the (code-mapped) bank line transaction type to the system transaction type. **[M]** | Relying on it when code mapping is incomplete (raw code never normalized) | Type never equals → no match |

**Grouping / advanced matching attributes** (required for grouped match types)
For One-to-Many, Many-to-One, and Many-to-Many, you must specify **grouping attributes** so Oracle aggregates the "many" side into a group and compares the group's summed amount (and shared attribute values) to the single side. Attributes documented as available include: **reconciliation reference, transaction date / booking date, transaction type, amount, and additional reference identifiers.** **[M]** *(Flag: the complete, exact catalog of grouping attributes on each side is UNCONFIRMED — verify in "Using Cash Management.")*

- **How "group by" works:** rows sharing the same value in the chosen grouping attribute(s) are collapsed into one candidate group; the group's total amount is what gets tested against the single counterpart under tolerance. **[M]**
- **Common mistake:** choosing a grouping attribute that isn't actually common across the set the bank batches together (e.g., grouping by reference when the bank batches by deposit/date), so groups never form and grouped matches never succeed.

**Advanced criteria / filter builder** *(this is the item I am least able to confirm verbatim — treat all exact column names as [U])*
- **Operators** typically available: **equals, not equals, LIKE, NOT LIKE, greater than, less than, greater/less than or equal, between, in.** **[M]**
- **Referencing fields:** conditions reference **bank-statement-line fields** vs **system-transaction fields**, distinguished by side/alias in the builder. **[M]**
- **Candidate bank-line field/column names** the query asks about — presented as *plausible* internal names to verify, **not confirmed [U]:** `ADDENDA_TXT` (addenda/free text), `CPARTY_NAME` (counterparty name), `TRX_TYPE` (transaction type), `BOOKING_DATE`, `TRX_DATE`, `RECON_MATCH_REFERENCE` / `EXTERNAL_RECON_REFERENCE` (reconciliation reference), plus structured payment reference, end-to-end ID, instruction/payment-instruction ID, account servicer reference, customer reference. System-transaction-side field names could **not** be confirmed. **[U]**
- **Literals/wildcards:** string literals are quoted and, with `LIKE`, wildcards are used for partial matches; conditions combine with **AND/OR** and (typically) parentheses grouping. Exact wildcard character(s) **[U]**.
- **Signature failure mode — inverted `LIKE`/`NOT LIKE`:** authoring `CPARTY_NAME NOT LIKE '%ACME%'` when you meant `LIKE` will exclude exactly the lines you intended to catch. Because the rule still "runs," this fails *silently* (no error, just zero or wrong matches). **Downstream effect:** an entire payment stream stays unreconciled or matches to the wrong system transactions.
- **Second signature failure mode — idealized counterparty name:** matching on the full legal/idealized counterparty name when the bank only transmits a short "Company Name" string (e.g., `ACME` not `ACME Manufacturing Holdings LLC`). Use `LIKE '%ACME%'` on the short token the bank actually sends. **Downstream effect:** near-zero match rate on an otherwise clean payment stream.

---

### OBJECT 2 — TOLERANCE RULES
**Task:** *Manage Bank Statement Reconciliation Tolerance Rules*. **[M]**

| Field | Allowed values / format | What it does | Common mistake | Downstream effect |
|---|---|---|---|---|
| Tolerance Rule Name | Free text, unique | Identifies the rule for pairing in a rule set | Naming after a bank not a behavior | Rule sprawl / mis-pairing |
| Description | Free text | Documents scope | Blank | Harder maintenance |
| **Date tolerance — Days Before** | Integer (days) | Allows the system transaction date to be up to *n* days **before** the bank line date and still match. **Applies to ALL match types.** **[H]** | Setting large symmetric windows "to be safe" | Amount-coincident items in the window auto-match the wrong pair |
| **Date tolerance — Days After** | Integer (days) | Allows the transaction date up to *n* days **after** the bank line date. Days-before and days-after are **separate** fields. **[H]** | Assuming one field covers both directions | Lagged bank bookings (typical) fail if only "before" is set |
| **Amount tolerance — Enabled** | On/Off | Turns fixed-currency variance on/off. **Applies to ONE-TO-ONE only.** **[H]** | Believing "disabled" = amount-blind matching | **Disabled = EXACT amount required.** Net-of-fee receipts (bank credit ≠ invoice by the fee) never auto-match |
| Amount tolerance — Amount Below | Number (currency) | Max allowed shortfall (bank < book) still treated as a match | Setting asymmetric values by accident | One direction of fee delta matches, the other doesn't |
| Amount tolerance — Amount Above | Number (currency) | Max allowed excess (bank > book) still matched | Too-wide value | Genuine discrepancies get silently reconciled and buried in the differences account |
| **Percentage tolerance — Enabled** | On/Off | Turns percentage variance on/off. **ONE-TO-ONE only.** **[M]** | Using both amount and % without understanding interaction | Whichever resolves first governs; unexpected matches |
| Percentage tolerance — Percent Below | Number (%) | Allowed % shortfall | Large % on large-value accounts | Big absolute residuals absorbed silently |
| Percentage tolerance — Percent Above | Number (%) | Allowed % excess | As above | As above |

**Behavioral rules to encode in the template:**
- **Scope:** amount + percentage tolerance = **one-to-one only**; date tolerance = **all** match types. **[H]** For grouped matches, amount agreement comes from the **summed group** total, not a per-line amount tolerance. **[M]**
- **Automatic vs manual reconciliation:** in **automatic** reconciliation the tolerance is applied strictly as a pass/fail gate; in **manual** reconciliation a user can typically override and accept a difference outside tolerance (with a warning), booking the residual. **[M — verify exact override behavior in current release.]**
- **Where the residual posts:** a difference *within* tolerance posts to the bank-account-defined **Reconciliation Differences account**, or to **Bank Charges** (for fee deltas) / **Exchange Gain-Loss** (for FX revaluation deltas), per the account's configuration. **[M]**
- **Pairing:** a tolerance rule is **paired with a matching rule** inside each **sequence line of a reconciliation rule set** (Object 4). One tolerance rule may be reused across many pairings.

---

### OBJECT 3 — PARSE RULE SETS AND PARSE RULES
**Task:** *Manage Parse Rule Sets*. **[M]**

**Nature & placement:** Parse rule sets are **OPTIONAL** and act as an **enrichment layer applied AFTER transaction-code mapping** during statement import/processing. They extract substrings from free-text/reference source fields and write them into structured statement-line target fields so downstream matching rules (especially reconciliation-reference matching) have clean values to compare. **[M]**

**Per-parse-rule fields**

| Field | Allowed values / format | What it does | Common mistake | Downstream effect |
|---|---|---|---|---|
| Sequence | Integer | Order of rule execution within the set | Two rules writing the same target with wrong order | Later rule overwrites (or is blocked by overwrite flag) unexpectedly |
| Enabled | On/Off | Activates the rule | Building then forgetting to enable | Extraction never happens; matching references stay null |
| Transaction Code | Bank code value | Scopes the rule to lines with that (mapped) code | Keying on raw code when only the type is populated | Rule never fires |
| Transaction Type | `CE_TRX_TYPE`-mapped value | Alternative/additional scoping to a normalized type | Type not mapped upstream | Rule never fires |
| **Source Field** | Statement-line source column (e.g., addenda text `ADDENDA_TXT`, account servicer reference `ACCNT_SERVICER_REF`, customer reference `CUST_REFERENCE`) **[U — exact names]** | The field the expression reads from | Pointing at a field the bank leaves empty | Nothing to extract → null target |
| **Target Field** | One of the populatable line fields: **reconciliation reference (`RECON_REFERENCE`), customer reference (`CUSTOMER_REFERENCE`), structured payment reference, end-to-end ID, instruction ID, transaction ID, check number, contract identification** **[U — exact names]** | Where the extracted value lands | Writing to a target the matching rule doesn't test | Extraction "works" but doesn't help matching |
| **Parse Rule / Extraction Expression** | Token grammar (see below) | Defines *how* to slice the source text | Literal anchor mismatch | **Silent non-extraction** (see below) |
| **Overwrite** | On/Off | Whether an already-populated (native or code-map-populated) target value is replaced | Leaving Overwrite ON and clobbering a good native reference | Good value replaced by a worse parsed one — or, if OFF, the parse is ignored when a value already exists |

**Parse expression token grammar** *(this is the single hardest-to-confirm item; presented as the commonly documented pattern but flagged [U] — verify verbatim examples in the guide):*
- **Whole-field / rest-of-field token** — a token (commonly rendered like `(X~)` or an equivalent "rest of string" marker) that captures the remainder of the source text from the anchor point. **[U]**
- **Numeric extraction tokens** — an `N` token, repeated to indicate digit counts, used for numeric segments; e.g., a mask like `0000000000(N)` to strip leading zeros and keep the significant digits. **[U]**
- **Character position ranges** — e.g., `(1-10)` or `(16-25)` to extract a fixed substring by position. **[U]**
- **Literal text anchors** — literal strings placed before and/or after the extraction token to grab text *between* two literals; e.g., an expression of the form `SENDING CO NAME:(X~)ENTRY DESC` extracts whatever sits between `SENDING CO NAME:` and `ENTRY DESC`. **[U]**
- **Delimiters/wildcards** — tilde (`~`) and period (`.`) style delimiter/wildcard characters appear in the grammar. **[U]**

**The defining failure mode — literal-anchor mismatch = silent non-extraction:** the literal anchors in your expression must match the bank's actual text **exactly** (spacing, punctuation, case as delivered). If the bank sends `SENDING COMPANY NAME:` but your anchor says `SENDING CO NAME:`, the token finds nothing, the target stays null, **no error is raised**, and every downstream reference-based match quietly fails. **Always author parse expressions from a real sample statement's raw addenda text, not from an idealized spec.**

---

### OBJECT 4 — RECONCILIATION RULE SETS
**Task:** *Manage Bank Statement Reconciliation Rule Sets* (may appear as "Manage Bank Statement Reconciliation Rules"). **[M]**

A reconciliation rule set is an **ordered/sequenced list**; each line **pairs exactly one matching rule with one tolerance rule.** The set is **assigned at the bank account level** and entries are evaluated **in sequence** — the first pairing that produces a valid match wins for that line.

| Field | Allowed values / format | What it does | Common mistake | Downstream effect |
|---|---|---|---|---|
| Rule Set Name | Free text, unique | Identifies the set for bank-account assignment | Generic name reused across dissimilar banks | Wrong set assigned to an account |
| Description | Free text | Documents scope/order rationale | Blank | Ordering logic lost to future maintainers |
| Sequence (per entry) | Integer | Evaluation order of the pairing | Broad rule at a low sequence number | **Broad rule pre-empts precise rules below it**, producing coarse/false matches and starving high-confidence rules |
| Matching Rule (per entry) | Existing matching rule | The cardinality + criteria to apply | Pairing a grouped matching rule with amount tolerance expecting per-line tolerance | Amount agreement comes from the group sum, not tolerance |
| Tolerance Rule (per entry) | Existing tolerance rule | The variance envelope for that matching rule | Pairing a wide tolerance with a broad matching rule | Compounded looseness → false matches |

**Best-practice ordering (encode in the template):** sequence **narrow, high-confidence rules first** (e.g., one-to-one exact match on reconciliation reference + zero/near-zero tolerance), then progressively broader rules (date-tolerant, then grouped, then amount/percentage-tolerant catch-alls) last. Rationale: because evaluation stops at the first successful match, a broad rule placed early consumes lines that a precise rule further down would have matched more accurately. **[H — this ordering principle is well established.]**

---

### OBJECT 5 — BANK STATEMENT TRANSACTION CREATION RULES
**Task:** *Manage Bank Statement Transaction Creation Rules*. **[M]**

**Nature & placement:** these run **AFTER** both automatic and manual reconciliation, and are intended **only** for **bank-originated "first notice" items** — items the bank tells you about that you have no prior system transaction for: **bank fees, interest, bank adjustments, NSF/returns.** They create the corresponding Cash Management transaction (and, optionally, its accounting) so the line can reconcile.

| Field | Allowed values / format | What it does | Common mistake | Downstream effect |
|---|---|---|---|---|
| Bank Account | Existing bank account | Scopes the rule to one account | Building at wrong account | Rule never triggers on intended lines |
| Enabled | On/Off | Activates the rule | Forgetting to enable | No transactions created |
| Sequence | Integer | Evaluation order among creation rules | Overlapping search strings out of order | Wrong rule claims a line first |
| Rule Name | Free text | Identifies the rule | — | — |
| Description | Free text | Documents intent | Blank | Maintenance risk |
| **Transaction Code** | Bank/BAI code value | Scopes to lines carrying that code | Keying on a code the bank doesn't send | Rule never fires |
| **Transaction Type** | `CE_TRX_TYPE` lookup value | Classifies the created transaction | Wrong type → wrong accounting | Misclassified fee/interest in GL |
| **Search Field** | Statement-line field to inspect | Which field the search string is tested against | Searching a field the value isn't in | No match |
| **Search String** | Text / wildcard pattern | Identifies qualifying lines | Too-broad string | Sweeps in lines that should reconcile to AR/AP |
| **Cash Account** | GL account | Cash side of the created entry | Wrong cash account | Cash misstated |
| **Offset Account** | GL account | The expense/income offset (e.g., bank-charge expense, interest income) | Wrong offset | P&L misclassification |
| **Accounting (auto-accounting) flag** | On/Off | Whether accounting is created automatically for the new transaction | On with wrong accounts | Auto-posts incorrect entries at scale |

**Standard delivered `CE_TRX_TYPE` values** *(commonly delivered set — verify the complete list against the CE_TRX_TYPE lookup in the current release [U]):* **Bank Fee (BKF), Interest (INT), Bank Adjustment (BKA)**, and related types such as NSF/return handling. The complete enumerated list (and each code's abbreviation) **could not be confirmed** in this session. **Sample bank transaction codes** are typically **BAI/BAI2 codes** mapped via Object 6.

**Anti-patterns to warn against in the template:**
- Creating external transactions for **customer receipts that should reconcile to AR** — this masks a receipt-entry backlog, double-counts cash, and breaks the AR-to-cash audit trail.
- Creating transactions for **zero-balance-account (ZBA) sweeps** — sweeps net to zero and should be modeled as transfers, not fabricated one-sided transactions.
- Overly broad **search strings** that claim lines a matching rule should have reconciled.

---

### OBJECT 6 — BANK ACCOUNT CONTROL POINTS & TRANSACTION-CODE MAPPING
**Tasks:** *Manage Bank Accounts* (reconciliation fields), *Manage Bank Statement Transaction Codes*, *Manage Code Map Groups* / Formats. **[M]**

**Bank-account reconciliation-relevant fields**

| Field | Allowed values / format | What it does | Common mistake | Downstream effect |
|---|---|---|---|---|
| Parse Rule Set | Existing parse rule set (optional) | Enrichment applied to this account's statements | Assigning a set built for another bank's text | Silent non-extraction across the account |
| Reconciliation Matching Rule Set | Existing rule set | The ordered matching/tolerance logic for the account | Leaving unassigned | **No automatic reconciliation occurs** |
| Reconciliation Start Date | Date | The date from which reconciliation/matching begins | Setting it after live balances exist | Pre-date items never considered / opening imbalance |
| Reconciliation Differences Account | GL account | Where within-tolerance residuals post | Unset | Tolerant matches can't book residual → matches fail |
| Bank Charges Account | GL account (**at business-unit-access level**) | Where fee deltas post | Set at wrong BU access level | Fee-delta matches error |
| Exchange Gain/Loss Account | GL account (**at business-unit-access level**) | Where FX residuals post | Unset for FX accounts | FX-tolerant matches fail |
| Tolerances (account-level) | Config values | Account-level variance defaults where applicable | Conflicting with rule-set tolerance | Unexpected match envelope |
| Secure Bank Account by Users and Roles | On/Off flag | Restricts account access/visibility by user and role | Enabling without granting roles | Users lose access to reconcile the account |

**Manage Bank Statement Transaction Codes & Code Map Groups — the PRIMARY normalization layer (runs BEFORE parse rules):**
- External bank codes (e.g., **BAI2**-type transaction codes) are mapped to Oracle internal transaction types (the **`CE_TRX_TYPE`** lookup) via transaction-code setup and **Code Map Groups / Formats** (which also support standards such as BAI2, SWIFT MT940, and EDIFACT). **[M]**
- **Why it comes first:** matching rules that test transaction type, parse rules scoped by transaction type, and creation rules scoped by code all depend on this normalization. **Downstream effect of a mapping gap:** the raw code is never normalized, so every downstream object keyed on transaction type silently fails. **Common mistake:** onboarding a new bank/statement format without extending the code map, then wondering why auto-reconciliation "stopped working" for that account.

---

## Recommendations

**Because live Oracle sources could not be reached this session, the first recommendation is a verification protocol; the rest are authoring guidance.**

1. **Verify before you build (do this first).** Open and confirm each flagged item against these specific primary Oracle pages for your exact release (25D/26A) in the Fusion Cloud Financials library on `docs.oracle.com`:
   - *Using Cash Management* → "Bank Statement Reconciliation" chapter → sub-topics on Matching Rules, Tolerance Rules, Reconciliation Rule Sets, Parse Rule Sets, and Transaction Creation Rules (confirm match-type list, criteria toggles, the **exact parse-expression token grammar with Oracle's own examples**, and the advanced-criteria operator list and field/column names).
   - *Implementing Payments and Collections* (or the Cash Management implementation content) → for the Setup and Maintenance **task names** and bank-account reconciliation setup.
   - Oracle Applications Help (in-product) → the "Manage Bank Statement Reconciliation …" task help topics.
   - The **Tables and Views for Financials** reference and **REST API for Oracle Fusion Cloud Financials** → to confirm the literal column/attribute names (`ADDENDA_TXT`, `CPARTY_NAME`, `TRX_TYPE`, `BOOKING_DATE`, `TRX_DATE`, `RECON_MATCH_REFERENCE`, `EXTERNAL_RECON_REFERENCE`, `RECON_REFERENCE`, `CUSTOMER_REFERENCE`, etc.), which are the least certain items here.
   - The **`CE_TRX_TYPE` lookup** (Manage Standard Lookups / the Cash Management lookups reference) → to capture the complete delivered value list and codes.

2. **Author in runtime order, and test each layer in isolation.** Build/confirm code mapping → import a real sample statement → confirm normalized types → add parse rules and confirm target fields populate on the actual sample → then build matching + tolerance rule sets → run auto-reconciliation on the sample → add transaction creation rules only for genuine first-notice residuals.

3. **Sequence rule-set entries narrow-to-broad.** Line 10: one-to-one exact on reconciliation reference, zero tolerance. Line 20: one-to-one on amount + date with tight date tolerance and small amount tolerance (to catch fee deltas). Line 30+: grouped types. Catch-all broad rules last.

4. **Guardrails on the specific failure modes:** (a) after authoring any `LIKE`/`NOT LIKE` condition, test it against known lines to confirm the polarity; (b) match counterparty on the **short token the bank actually transmits**, not the legal name; (c) copy parse-rule literal anchors **from the raw statement text**, character-for-character; (d) keep date tolerance as narrow as your bank's real booking lag (often 0–2 days), never "wide to be safe"; (e) keep **amount tolerance ENABLED** wherever net-of-fee matching is expected; (f) restrict transaction creation rules to bank-fee/interest/adjustment/NSF codes only.

**Thresholds that would change the guidance:** if auto-match rate is high but *false* matches appear → tighten tolerances and move broad rules later. If match rate is low with clean references → a parse rule likely isn't populating the reference (check the literal anchor). If a payment stream never matches on counterparty → switch to the short bank-sent token. If fees never reconcile → confirm the code map and the transaction creation rule's search string/code.

---

## Caveats

- **No primary-source verification was possible in this session.** `web_search`, `web_fetch`, and the research subagent all returned tool-unavailability errors; **no `docs.oracle.com` page was opened and no verbatim Oracle quote or live URL was captured.** Treat this document as an authoring scaffold to be confirmed, not as citation-backed primary content. I did not manufacture any source URLs or quotes.
- **Highest-uncertainty items (verify verbatim before relying on them):** (1) the exact **parse-expression token grammar** and Oracle's own example expressions; (2) the exact **advanced-criteria operator list** and the literal **bank-line/system-transaction column names**; (3) the **complete delivered `CE_TRX_TYPE` lookup** values and codes; (4) the precise, current **Setup and Maintenance task names** (labels shift slightly between releases); (5) any **zero-amount/adjustment match type** existence and label; (6) the exact **automatic-vs-manual tolerance override** wording.
- **Release sensitivity:** I could confirm **no** 25D/26A-specific change in this session. Oracle's quarterly "What's New" / readiness material for 25A–26A should be checked for any changes to matching-rule attributes, parse-rule grammar, or new match types.
- **Source-type distinction:** everything above is my own reconstruction (effectively an *unsourced tertiary* account). When you verify, prioritize **primary Oracle** pages (docs.oracle.com guides, Applications Help, Tables/Views and REST references) over **third-party** material (consultancy blogs, Oracle Support community posts), and note in your template which cells were confirmed against which primary page.
- **Business-unit vs bank-account level:** bank charges and exchange gain/loss accounts are noted as configured at the business-unit-access level of the bank account; confirm this placement in your instance, as access-model details can vary by setup.