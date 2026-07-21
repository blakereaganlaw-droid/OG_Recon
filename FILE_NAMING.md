# File naming & format guide

The router classifies every file **by its name** — the name IS the routing
instruction. The engine announces, the moment it routes, every file it will
ignore (reason + suggested rename, on the console and in the run log /
manifest) — nothing is ever dropped silently. Following this guide makes
every upload unambiguous on the first try.

## The recommended pattern

```
YYYYMMDD_<Source>_<Account>_<Role>[_<Status>].<ext>
20260719_Oracle_CM_FHB_Master_BSL_UNR.xlsx
20260719_Oracle_CM_FHB_Master_ST_UNR.xlsx
20260719_Oracle_OTBI_MET_All.xlsx
20260719_Oracle_Payables_Payments.xlsx
20260718_FHB_Master_BAI2.txt
```

1. **Date stamp first, `YYYYMMDD`.** When two files claim the same role,
   the newest stamp wins; two *equal* stamps fail loud (except MET — see
   pagination below). An undated file always loses to a dated one. An
   8-digit run that isn't a plausible calendar date (a merchant number
   like `99999999`) is never mistaken for a date.
2. **Account token in the name** for account-specific exports (BSL, ST,
   single-account MET, BAI2). One upload = one account; two different
   account tokens in one upload fail loud before anything runs.
3. **Role keyword somewhere in the name** — this is what routes the file.

## Role tokens (complete)

| Export | Name must contain | Loaded? | Example |
|---|---|---|---|
| Open bank statement lines | `BSL` (not `All`, `All_Data`, `Enriched`, `Reconciled`) | **yes — required** | `20260719_Oracle_CM_FHB_Master_BSL_UNR.xlsx` |
| All-accounts open BSL | `All` + `BSL` | yes (reverse-misdirected search only) | `20260719_Oracle_OTBI_All_BSL_UNR.xlsx` |
| Open system transactions | `_ST_` / `_ST.` / `Account_ST` | yes | `20260719_Oracle_CM_FHB_Master_ST_UNR.xlsx` |
| MET / ORT bridge (OTBI) | `MET` or `Oracle_OTBI` | yes | `20260719_Oracle_OTBI_MET_All.xlsx` |
| Receivables receipts | `Receivables_Receipts` / `Receipts_All` / `Oracle_Receipts` | yes | `20260718_Oracle_Receivables_Receipts.xlsx` |
| AP payments | `Payables` **and** `Payments` (both) | yes | `20260719_Oracle_Payables_Payments.xlsx` |
| Raw BAI2 transmission | `BAI` token **OR recognized by content** (a `.txt` opening with an `01,` header + `02,`/`16,` records — e.g. `BAIEXP_07202026_071541.txt`) | yes (enrichment) | `20260718_FHB_UTHSC_BAI2.txt` |
| BAI2 spreadsheet | `BAI` token **OR recognized by content** (a csv/xlsx whose header carries a `BAI Code` + date/bank-reference column — e.g. `20260720_FHB_Master.xlsx`) | yes (enrichment) | `20260715_FHB_UTC_BAI2.csv` |
| Edison payments / invoices | `Edison_Payments` / `Edison_Invoices` | yes (Review annotation) | — |
| GMS Sponsored AR Aging | `GMS_001` / `Sponsored_AR_Aging` | yes (Sponsored-Projects annotation) | `RPT_GMS_001__Sponsored_AR_Aging_Report.xlsx` |
| GMS Sponsored Award Profile | `Sponsored_Award_Profile` / `RPT_GMS_002` | yes (SPN/award/sponsor bridge) | `RPT_GMS_002__Sponsored_Award_Profile_Report.csv` |
| GMS Project Contract & Funding | `Project_Contract_And_Funding` / `RPT_GMS_035` | yes (SPN→award→sponsor bridge) | `RPT_GMS_035__Project_Contract_and_Funding_Source_Report.xlsx` |
| AR Unapplied by Customer | `AR_Unapplied_By_Customer` / `RPT_AR_059` | yes (unapplied-receipt annotation) | `RPT_AR_059__AR_Unapplied_By_Customer_Report.xlsx` |
| AR Unapplied Receipts Summary | `AR_063` / `RPT_AR_063` / `Unapplied_Receipts_Summary` | yes (customer unapplied context) | `RPT_AR_063__Unapplied_Receipts_Summary_Report.xlsx` |
| Active Award Conversion Report | `Award_Conversion` / `Oracle_Award_Conversion` | yes (SPN annotation — resolves the SPONSOR's own contract number to our SPN/award) | `Active_Oracle_Award_Conversion_Report.txt` (pipe-delimited) |
| MID master | `MID_Master` | yes | `UT_MID_Master_Consolidated.xlsx` |
| ORT misc / AR | `ORT`+`Misc` / `ORT`+`_AR` | yes (data-feed-error sweep + per-receipt `REFERENCE_TEXT` cross-reference) | `20260719_ORT_Misc_All.xlsx` |
| ORT departments | `ORT_Department` / `Department_Info` | yes (MID directory) | `ORT_Department_MIDs.xlsx` |
| Chart of Accounts bundle | `Chart_Of_Accounts` / `AcctCombos` / `ComboSets` / `CombosTech` / `Segments` / `GL_Departments` | yes (decode labels; multi-file union) | `AcctCombos_base.csv` + shards |
| CM config: creation rules | `Transaction_Creation_Rules` | yes (config audit) | `CM_Configurations_Transaction_Creation_Rules.xlsx` |
| CM config: parse / matching / tolerance / rulesets | `Parse_Rules` / `Matching_Rules` / `Tolerance_Rules` / `Recon_Rulesets` | yes (config audit) | `CM_Configurations_Parse_Rules.xlsx` |
| Applied/Unapplied receipts | `Applied` or `Unapplied` | yes | — |
| AR matched receipts feed | `AR_Matched` / `Deposit_Receipts` | yes | `AR_Matched_Invoice_Receipts_...csv` |
| Enriched BSL workbook | `Enriched` / `Crossref` | yes (enrichment) | — |
| **Reconciled forensic exports** | `Reconciled` / `Reconciliation_Report` | recognized; `Reconciled_*` (Exported sheet) feeds the ADVISORY recon-history/orphan audit — never a pool source | `..._Reconciled_..._BSL.xlsx` |
| Lifecycle workbook | `All_Data` | recognized, never loaded (Unreconcile2) | `FHB_UTC_All_Data.xlsx` |
| Other GMS / AR reference reports | `RPT_GMS_0…` (other), `AR_Invoices`, `Contracts_To_Receivable_Invoices`, `Relationship_Map`, `Rosetta` | recognized, **not used by the forward engine** (you'll get a console NOTE) | — |

Account tokens: `FHB_Master`, `FHB_UTC` (also `FHB_UT_Chatt`), `FHB_UTHSC`,
`FHB_UTIA`, `FHB_UTM`, `FHB_UTSO`, `FHB_AP` (or `FHB_Accounts_Payable`),
`Regions_Master`, `Regions_UTM`, `Regions_UTIA`, `Regions_UTIPS`,
`Regions_UTSI`, `Regions_UTHSC`, and the five Student Refund depositories:
`FHB_Student_Refund_UTK` / `_UTC` (or `_UT_Chatt`) / `_UTHSC` / `_UTM` /
`_UTSO`. Student Refund tokens always win over the bare campus token, so
`FHB_Student_Refund_UTC_...` never lands in the FHB_UTC scope.

## Formats

- **`.txt` is accepted for raw BAI2 transmissions and for the Active Award
  Conversion Report** (pipe-delimited; the reader sniffs the separator). Any
  other `.txt` is skipped with a rename suggestion.
- **`.xlsx` preferred for correctness; `.csv` STRONGLY preferred for the
  big exports (speed).** Reading is the single largest cost of a run, and
  it is dominated by openpyxl's XML parsing of large `.xlsx` files. The
  same data as `.csv` reads **~6–7× faster** (measured on real Master:
  `ORT_Misc` 155k rows = **34s as `.xlsx` vs 5s as `.csv`**). For the
  three heavy exports — **`ORT_Misc`, `Payables_Payments`, and the `MET`**
  (each tens-to-hundreds of thousands of rows) — export as `.csv` and the
  whole run roughly halves. Reference-bearing columns are safe in `.csv`
  (the reader never float-coerces them); the small exports (BSL, ST,
  Edison, GMS/AR, MID) can stay `.xlsx`.
- `.xlsm`, `.xlsb` (needs `pyxlsb`), and `.csv`
  work. **`.txt` is accepted ONLY for raw BAI2 transmissions** and only
  when the name carries a `BAI` token. **`.xls` (legacy Excel) is not
  readable** — you'll be told immediately; re-export as `.xlsx`/`.csv`.
- Export **values, not formulas**.
- Never open-and-resave a reference-bearing export through tools that
  "help": auto-formatting can float-coerce `0006789599` into `6789599.0`
  and destroy the join key. Raw exports straight from Oracle/the bank are
  ideal (`.xlsb` integral-float artifacts are auto-repaired).
- Claude-web upload prefixes (`933782d6-Name.xlsx`) are stripped at
  staging; a plausible leading `YYYYMMDD` is kept.
- Skip exotic characters; letters, digits, `_`, `-`, `.` are safe
  everywhere.

## Actual and potential problems (and the fix)

| Problem | What happens | Fix |
|---|---|---|
| **`Reconciled_..._BSL.xlsx` in an upload** | Previously would have bound the open-BSL role and poisoned the run with already-reconciled lines. Now routes to the `RECONCILED` role (advisory recon-history audit only). | Keep `Reconciled` in the name (it's the protection); don't strip it. |
| **All-accounts BSL named without `All`** | Would bind as THIS account's open BSL → thousands of foreign lines. | Always keep `All_BSL` together (`Oracle_OTBI_All_BSL_UNR`). |
| **Raw bank file named `BAIEXP_….txt` / `…_FHB_Master.xlsx`** | No `BAI` token, but recognized by CONTENT (`01/02/16` records, or a `BAI Code` header) and routed as BAI2 (announced as a content NOTE). | Works as-is; a `_BAI2` token still makes it unambiguous. |
| **Same date stamp on two exports of one role** | Run stops and asks you to disambiguate — nothing silently ignored. | Bump the newer file's date stamp. Exception: **page shards** of BSL/ST/Receipts/Payments/ALL_BSL/MET (`..._X.csv` + `..._X_2.csv`, same date) are unioned automatically — with a fail-loud guard if the "pages" carry duplicate rows (a re-upload is not a page). |
| **Undated refresh beside a dated file** | The dated file wins even if the undated one is newer. | Always date-stamp refreshes. |
| **`UT Chatt` spelling** | Now maps to UTC (previously invisible to the mixed-account guard). | Prefer `UTC` for consistency. |
| **BAI2 window vs open lines** | ALL staged BAI2 files now union (same transaction deduped by bank reference, richer addenda kept) — a July file plus an older file covers both windows. | Stage every BAI2 file spanning the open-line window. |
| **`_ST_` requirement** | A file named `...FHB_Master_ST.xlsx` works (`_ST.`), but `...Master_STATUS.xlsx` will not route as ST (deliberate — `All_Status`/`Rosetta_Stone` protection). | Use `_ST_UNR` / `_ST.` forms. |
| **Generic names** (`export.xlsx`, `data.csv`) | No rule matches → announced and skipped. | Use the recommended pattern. |
| **Mixed-account uploads** | Preflight fails loud listing the conflicting tokens. | One account per upload (all-accounts MET/ALL_BSL/receipts/payments/config/CoA files are account-neutral and always fine to include). |

## What the engine tells you, immediately

At routing time (console + run log `files_ignored_by_name` + manifest):

```
IGNORED (file name): notes.txt — .txt is accepted only for raw BAI2 transmissions  FIX: export as .xlsx/.csv, or add a _BAI2 token if it is a raw bank file
IGNORED (file name): legacy.xls — legacy .xls format is not readable  FIX: re-export as .xlsx or .csv
IGNORED (file name): mystery_export.xlsx — no router rule matches this name  FIX: rename with the export's role token (see FILE_NAMING.md)
NOTE (file name): FHB_UTC_All_Data.xlsx recognized as ALL_DATA — lifecycle workbook — Unreconcile2's fuel; the forward engine will not read it.
```

## Known gaps (data we don't ingest yet)

- **GMS aging, AR invoices, sponsor map, contracts, unapplied summary** —
  routed but unused by the forward passes; they would enrich Review
  annotations if ever needed.
- **`Reconciliation_Report_*` renderings** — routed to RECONCILED but
  skipped by the recon-history audit (they lack the `Created By` actor
  column; stage the `Reconciled_*` Exported-sheet exports instead).

(Previously-listed gaps now BUILT: Recon History in-engine (R2 orphan
audit), multi-BAI2 union across statement windows, pagination union for
BSL/ST/Receipts/Payments/ALL_BSL, and the audit's C11 ST-id membership
check.)
