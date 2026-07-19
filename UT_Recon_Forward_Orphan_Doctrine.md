# UT Reconciliation — Forward Matching Under Prior Reconciliations (Orphan Doctrine)

**Scope.** This document teaches a **forward-looking** matching engine — one that
pairs OPEN bank statement lines with OPEN system transactions — that the pool it
draws from has already been worked, sometimes wrongly. Prior reconciliations
(especially automated ones) **consume** transactions and **orphan** others. An
engine that treats "open" as "money not yet in the bank" will force-match
orphans to the wrong lines and fail to explain stranded lines whose true
counterparts were already consumed. The ORT and MET data are the key to telling
these cases apart.

**Audience.** The forward engine (OG_Recon or successor). The backward engine
(Unreconcile2) already hunts these defects in reconciled history; this document
transfers what it learned — verified on four real accounts (Regions UTIPS,
Regions UTIA, FHB AP, FHB UTHSC) — into rules the forward engine can apply
*before* making a match.

**Companions.** `UT_Recon_ORT_Data_Relationships.md` (fields and joins, entity
by entity), `UT_Recon_SPN_Receivables_Data_Relationships.md` (Receivables side),
`UT_Recon_Engine_BUILD_SPEC.md` §10 (the backward contract). Terminology here
follows the ORT relationships document exactly.

---

## 1. The premise the forward engine must drop

The naive premise: *an open ST is a transaction whose money has not yet been
matched, and an open BSL is money no ST accounts for; pair them and both books
close.*

That premise fails on every validated account, because the open pool is shaped
by everything that already happened to it. Two corollaries replace it:

1. **The open pool contains orphans.** An open ST may represent money that is
   *already in the bank and already reconciled* — to an auto-created twin, to a
   same-amount stranger, or to nothing the depositor intended. Its money is
   spoken for; the ST itself was simply never the one consumed.
2. **The open pool is missing consumed counterparts.** A stranded BSL may have
   no open counterpart *because its true counterpart was consumed by an earlier
   (possibly wrong) reconciliation of a different line*. No amount of searching
   the open pool will close that line; the fix is an unwind, not a match.

Scale, from the validated accounts: FHB UTHSC's history is dominated by
automation — roughly **11,000 ESSADMIN (AutoReconcile) groups and 6,500
OIC_SYSTEM_USER (Transaction Creation Rule) groups**, from which the backward
re-audit produced 3,342 unwind recommendations. Regions UTIPS carries **751
OIC-created groups riding transaction codes 165/175 that no surviving
Transaction Creation Rule even covers** — the offending rules were removed
after the damage was done. Orphaning is not an edge case; on these accounts it
is the dominant failure mode.

---

## 2. How orphans are made — the verified factories

Every orphan mechanism below was confirmed against real export sets. The
`Created By` column of the reconciled-group export names the actor and is the
first thing to read.

### 2.1 Transaction Creation Rules — `Created By = OIC_SYSTEM_USER`

The University over-uses Transaction Creation Rules (TCRs). When a bank line
arrives whose transaction code (or addenda `SRCH STRNG`) hits an enabled TCR,
Oracle **auto-creates a brand-new external ST from the bank line and reconciles
the two in one motion**. The auto-created ST posts to the rule's configured
generic CASH/OFFSET GL strings (`CM_Configurations_*` → Transaction Creation
tab).

Meanwhile the department had entered — or will enter — the *same money*
intentionally through ORT: a parked receipt carrying the depositor's real
intended cash/offset GL entities. That receipt becomes an ST too. But its bank
line is already taken by the auto-created twin, so the user's ST **orphans
open, forever**, while the money sits booked to the TCR's generic accounting
instead of the department's intent.

What the forward engine sees later: an open, user-entered ST with exact cents,
plausible date, plausible type — and no bank line for it, because its line was
consumed at arrival. That ST is a lure. Matching it to some *other*
same-amount line compounds the original defect (this is defect
`TCR_ORPHANED_ST`, §10.2 check 11 in the backward spec).

### 2.2 AutoReconcile — `Created By = ESSADMIN`

The AutoReconcile batch matches an open BSL to open STs on **amount + type +
reference** only, with a date window of at most the configured tolerance
(±15 days where date matching is enabled at all; some matching rules run with
`DTE MTCH = N` and ignore date entirely — always read the account's
`CM_Configurations_*` Tolerance export rather than assuming). **Payer, GL
entity, and depositor intent are not part of the comparison.**

Consequence: two STs with the same signed cents, same type, and same reference
inside the window are **indistinguishable to Oracle**. The batch consumes one
— not necessarily the right one. The true owner stays orphaned open
(`AUTOREC_AMBIGUOUS`), and when the consumed competitor posts to a *different
GL entity* than the orphan, one entity's money has been booked to another
(`AUTOREC_ENTITY_CONFLICT` — the worst case, because the ledger is now wrong,
not just the pairing).

Same-amount collisions are common by construction: departmental deposits
cluster at round and repeated amounts (fee schedules, ticket prices, standard
remittances), and on check rails the amounts repeat across unrelated checks.
FHB AP showed the compounding form: **a 4-link cascade of $1,100 checks**, each
wrong consumption stranding the next true owner, which the batch then "solved"
with the next wrong consumption (16 High `CHECK_REF_WRONG_ST` findings total).
One blind pull can wrong-foot the pool for every same-amount item after it.

### 2.3 Dual-fire receipts

One ORT receipt (`RECEIPT_ID`) occasionally spawns **two identical open
external STs** (132 confirmed on FHB Master). Prior reconciliation consumes
one; its twin orphans open. The twin is not new money — it is a duplicate of
money already matched. Dedup by `RECEIPT_ID` before any sum, and never match a
dual-fire twin to anything.

### 2.4 The compounding rule

Factories 2.1–2.3 interact. A TCR consumption removes a line from the pool and
adds a generic ST; that worsens the ESSADMIN same-amount ambiguity for
everything left; a wrong ESSADMIN pull then strands another true owner. The
forward engine must therefore assume the pool's *shape* is partly an artifact
of prior automation — not a clean picture of unmatched money.

---

## 3. Why ORT and MET are the key

Prior reconciliations are recorded in Recon History, but history only says
*what was paired*, never *what was intended*. Intent lives in ORT; identity and
status live in MET. Together they let the engine distinguish a true counterpart
from an orphan.

### 3.1 MET — identity and status truth

- **`CET_TRANSACTION_ID` = ST Transaction Number** is a perfect 1:1 bijection
  (100% coverage on every account tested: 13,522 / 1,603 / 436 external STs).
  Given any candidate ST, MET names its receipt and deposit. Never substitute
  the truncated Combined ID.
- **MET `Status` (UNR/REC)** is an independent record of whether the receipt's
  money is consumed. A candidate ST whose MET row is REC — or whose
  `CET_CLEARED_DATE` is populated (it is null on UNR rows and agrees with the
  BSL booking date ±1 day on REC rows) — is already spoken for, whatever a
  stale ST export says. **Do not consume it.**
- **`CET_DESCRIPTION`** (`d:{DEPOSIT_ID} | r:{RECEIPT_ID} | payer/purpose`)
  supplies the deposit grouping and payer corroboration that amount alone
  never can.

### 3.2 ORT — the intent record

The ORT parked-receipt exports hold what the depositor *meant*:

- **Join: MET `RECEIPT_ID` = ORT `Parked Receipt Item ID`.** This is the
  receipt **ITEM** id, *not* `Parked Receipt ID` — the wrong column joins
  almost nothing; the right one joined 3,647 of 3,790 real rows.
- Each ORT row carries the deposit's intended **CASH GL entity and OFFSET GL
  entity** (the account strings the depositor directed the money to), the
  business unit, and the receipt reference text.
- The generic `ORT` export's entity strings are the ground truth the backward
  engine quotes when it shows that an auto-created ST displaced a user's
  intent. The forward engine uses the same evidence *prospectively*: of two
  same-amount candidates, the one whose ORT intent coheres with the bank
  line's context (entity, payer, deposit id) is the true owner.

### 3.3 The intent test (the tie-breaker automation lacks)

When candidates are indistinguishable on amount/type/reference — precisely the
situation where ESSADMIN guesses — the engine breaks the tie with evidence
Oracle never consulted:

1. Bridge each candidate ST → MET row (`CET_TRANSACTION_ID`).
2. Join MET `RECEIPT_ID` → ORT `Parked Receipt Item ID`.
3. Compare each candidate's cash/offset **entity segments** (leading 2-digit
   prefixes) against what the line's context implies, component-wise —
   compare only segments both sides actually carry (the UTIA MET export has no
   asset segments; its `ENTITY` column supplies the cash-side prefix, and an
   offset-only divergence is still a cross-entity signal).
4. A candidate whose intent diverges is *not* promoted to Match on amount
   coincidence — it is at best a flagged Candidate naming the conflict.

### 3.4 The config side

- **`CFG_TOLERANCE`** supplies the account's real day window (±15 confirmed
  from config on the validated accounts — read it, never assume it).
- **`CFG_TCR`** supplies each enabled TCR's transaction code and CASH/OFFSET
  posting, scoped **strictly** to the account (drop rules whose bank name maps
  elsewhere or nowhere). A bank line whose transaction code is covered by an
  enabled TCR will be consumed *at arrival* by an auto-created ST — the
  forward engine should expect no open counterpart for such lines and expect
  an orphaned user ST nearby instead. Remember the UTIPS gap: consumption in
  history may ride rules **since deleted**, so an uncovered code does not
  prove no TCR ever fired.

---

## 4. Orphan signatures in the open pool

| # | Signature | Likely meaning | Forward action |
|---|---|---|---|
| 1 | Open user-entered ST; a REC group exists with the same signed cents, compatible type, in-window date, `Created By = OIC_SYSTEM_USER` | TCR consumed the ST's line with an auto-created twin | Do not match elsewhere; refer group for unwind (`TCR_ORPHANED_ST`) |
| 2 | Open ST with a same-amount/type/reference REC sibling, `Created By = ESSADMIN`, entities differ | Blind pull consumed the wrong one | Do not consume; refer (`AUTOREC_ENTITY_CONFLICT`) |
| 3 | Open ST whose MET row is REC / `CET_CLEARED_DATE` populated | Money already cleared under another identity | Never consume; reconcile exports disagree — investigate |
| 4 | Two identical open STs sharing one MET `RECEIPT_ID` | Dual-fire twin | Dedup by `RECEIPT_ID`; match at most one, flag the twin |
| 5 | Stranded BSL, no open counterpart at exact cents, line's trx code covered by an enabled TCR (or in a known deleted-rule code like UTIPS 165/175) | Counterpart was auto-created and consumed at arrival | Stop searching the open pool; the fix is upstream |
| 6 | Stranded BSL; an all-REC MET deposit sums exactly to the line | The deposit's receipts were cherry-picked onto other lines | Refer the consuming groups for unwind; do not force partial matches |

---

## 5. Rules for the forward engine

**R1 — Amount is never identity.** Exact signed cents is necessary for every
match and sufficient for none. A same-amount open ST with no reference tie, no
type corroboration, and no intent coherence is as likely an orphan as a
counterpart. (Unknown types never corroborate; amount alone never names a
match — same doctrine the backward engine applies to naming orphans.)

**R2 — Read `Created By` before trusting the neighborhood.** Before consuming
an open ST, scan Recon History for REC groups at the same signed cents within
the configured window. A hit created by `ESSADMIN` or `OIC_SYSTEM_USER` means
automation already worked this amount — run the intent test (§3.3) before
matching, and prefer referral (R9) over a low-evidence match.

**R3 — MET status outranks the ST export.** If MET says REC (or carries a
`CET_CLEARED_DATE`) for a candidate's receipt, the candidate is consumed. Do
not match it, whatever the ST export's recon status says.

**R4 — Corroborate intent, not just identifiers.** For departmental deposits,
walk ST → MET → ORT and compare intended entities with the line's context.
Identifier agreement plus intent divergence is a Candidate with a named
conflict, never a Match.

**R5 — Expect consumption on TCR-covered codes.** For a bank line whose code
an enabled TCR covers, the expected outcome is that Oracle already consumed
it. Finding it open is itself a signal (rule disabled? scope gap?) — and
finding a same-money *user* ST open beside a TCR group is the classic orphan.

**R6 — Respect rail compatibility, both directions.** Customer check deposits
post as `Miscellaneous` bank lines: a CHECK-typed ST against a MISC line is
normal and never disqualifying. ACH-typed STs funding MISC lines are flagged
for review (standing user decision, 2026-07-14), not silently accepted.

**R7 — Never require reference equality on deposit rails (Regions).** Regions
deposit references are **dual-encoded**: the bank reference (a location /
'CHECK DEPOSIT PACKAGE' code like 60000/30010, or a batch serial) and the
check/receipt reference (an Oracle deposit-ticket or receipt number) are
*different identifier systems for the same deposit* — they are supposed to
differ. Do not build a deposit-slip reference-tie check; all 7 candidate
findings from such a check were refuted against BAI2/MET/ORT. The real
deposit-check signal is DATE (late receipt entry).

**R8 — On check rails, the check number is the identity.** Reference = check
number on check rails (FHB AP). Same amount + different check number is a
*conflict*, never a partial match — this is exactly how the $1,100 cascade
propagated. Compare check numbers canonically (strip leading zeros; read
references as text — float coercion corrupts `0006789599`).

**R9 — Refer, don't route around.** When evidence says a prior reconciliation
consumed the wrong item, the forward engine must not "work around" it by
matching the orphan to a second-best line. That converts one defect into two.
Emit the finding (group id, orphan ST id, evidence) for the backward engine
(Unreconcile2), which recommends the unwind and proposes the re-reconciliation
from the live pool. After the unwind, the freed line and ST return to the open
pool and the forward match becomes clean — in the right order.

**R10 — Determinism and hygiene carry over.** Integer cents via `Decimal`
only; dedupe members by Transaction Number keeping the largest magnitude;
canonicalize legacy-reader floats (`683.0` → `'683'`); normalize long MET bank
account names to the short engine name before any deposit lookup; join on
`DEPOSIT_ID`/`RECEIPT_ID` separately, never the truncating Combined ID; treat
OTBI ALL_DATA `Dep Amnt` as an unsigned magnitude.

---

## 6. What is NOT an orphan — false-positive guards

Careful drafting cuts both ways: these patterns *look* like orphans and are
not. Each guard below was adversarially verified on real data.

1. **Old open checks.** Check float is long — FHB AP's paid-check float runs
   to a p99 of **95 days**, and 81 genuinely stale checks were still paid. An
   open check ST months old is usually a slow check, not an orphan. Check-rail
   date logic uses the 180-day STALE threshold, not the deposit window.
2. **Divergent Regions deposit references** (R7). Supposed to differ. Sibling
   references (e.g. batch serials 118004/118001) can even belong to *one*
   deposit — verified: one Lions Club $700 deposit.
3. **Co-members of a 1:M group.** A group legitimately holding two
   same-amount/type/reference members *consumed both* — ambiguity exists only
   against transactions **outside** the reconciled set. Never count a
   co-member as a competitor.
4. **Multi-line groups' "untied" members.** On a multi-line group, members
   that don't tie one line exactly fund the group's *other* line (verified on
   UTIA). One-to-many suspicion applies to **single-line** groups only.
5. **Blank bank references.** Large ACH disbursements post with reference `NA`
   (FHB UTHSC); an empty bank reference is not a reference mismatch — skip
   the comparison, don't fail it.
6. **CHECK ST on a MISC line** (R6). Normal, always, everywhere.

---

## 7. Worked examples (illustrative values)

### 7.1 The TCR orphan

A $2,450.00 ACH lands on code 165. An enabled TCR covers 165, so Oracle
creates ST `EXT-90211` from the line (posting to the rule's generic
CASH/OFFSET strings) and reconciles them — `Created By: OIC_SYSTEM_USER`.
Two days later the department's ORT deposit posts: receipt item `198564`,
$2,450.00, intended cash entity `01`, becomes user ST `EXT-90544`.

Forward engine, weeks later, sees `EXT-90544` open and a *different* open
$2,450.00 bank line from another payer. R1/R2 fire: history holds a REC group
at 245000 signed cents, `OIC_SYSTEM_USER`, in-window. MET bridges `EXT-90544`
→ `r:198564`; ORT (`Parked Receipt Item ID` = `198564`) shows intent entity
`01` — matching the *consumed* line's context, not the new line's. Verdict:
`EXT-90544` is the TCR group's orphan. Refer for unwind; do not attach it to
the new line.

### 7.2 The ESSADMIN entity conflict

Two $500.00 EFT STs share reference `1375681` within ±15d: `EXT-71001`
(entity `01`) and `EXT-71002` (entity `07`). AutoReconcile (`ESSADMIN`)
consumed `EXT-71002` against a line whose ORT deposit directed cash to entity
`01`. The forward engine later finds `EXT-71001` open. Intent test: the
consumed competitor's entity (`07`) conflicts with the deposit's intent
(`01`) — signature #2. The open ST is the *true owner*, orphaned. Refer
(`AUTOREC_ENTITY_CONFLICT`); matching `EXT-71001` to any other line would bury
the cross-entity booking.

### 7.3 The stranded line with no counterpart

A $1,134.00 MISC line strands: no open ST at 113400 cents ties. MET shows an
all-REC deposit `d:1048221` whose deduped receipts sum to exactly $1,134.00 —
its three STs were consumed piecemeal onto other lines (cherry-pick split).
Signature #6: stop searching the open pool; the three consuming groups are
unwind referrals, and the line will close cleanly only after they release.

---

## 8. Join quick reference

| Join | Direction / cardinality | Transform | Hazard |
|---|---|---|---|
| ST `Transaction Number` = MET `CET_TRANSACTION_ID` | 1:1 bijection | none | Never use Combined ID (truncates) |
| MET `RECEIPT_ID` = ORT `Parked Receipt Item ID` | M:1 | `N()` canonicalize (floats from legacy readers) | The similarly-named `Parked Receipt ID` is the WRONG column |
| MET `DEPOSIT_ID` groups `RECEIPT_ID` | 1:M parent→child | — | Dedup by `RECEIPT_ID` (dual-fire) then Transaction Number (total vs split rows) before summing |
| Recon History group ↔ `Created By` | attribute | `_norm_header` | `ESSADMIN` = AutoReconcile; `OIC_SYSTEM_USER` = TCR; else human |
| BSL trx code → `CFG_TCR` (enabled, account-scoped) | M:1 | strict account scope | Deleted rules leave uncovered consumption (UTIPS 165/175) |
| Day window → `CFG_TOLERANCE` | config | max(`DAYS BFR`,`DAYS AFTR`) over date-enabled rules | Never assume ±15 — read it |
| Bank account: BSL short name = MET long name | 1:1 | long→short normalization | Skipped normalization = every deposit lookup silently empty |
| Check reference (check rails) | identity | strip leading zeros, text-only reads | Same amount + different check number = conflict, not candidate |

---

*Doctrine sources: `CLAUDE.md` (per-account verified facts, 2026-07),
`UT_Recon_Engine_BUILD_SPEC.md` §10.2 checks 9/11/12,
`UT_Recon_ORT_Data_Relationships.md` §2.3/§4.2/§6, and the
`unreconcile_engine.py` checks `_check_tcr_orphans` / `_check_autorec_flaw` /
`_load_ort_index`. Spec wins over prose; where this document and the engine
disagree, fix one of them loudly.*
