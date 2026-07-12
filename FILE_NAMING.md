# File naming & format guide

The router classifies every file by its **name**, so names matter. The engine
is hardened against most accidents (short tokens like `ar`/`st` only match
whole name segments), but following this guide makes every upload
unambiguous.

## The recommended pattern

```
YYYYMMDD_<Source>_<Account>_<Role>[_<Status>].xlsx
20260710_Oracle_CM_FHB_Master_BSL_UNR.xlsx
20260710_Oracle_CM_FHB_Master_ST_UNR.xlsx
20260710_Oracle_OTBI_MET_All_Accounts_All_Status.xlsx
20260710_FHB_Master_BAI2.xlsx
```

1. **Date stamp first, `YYYYMMDD`.** When two files claim the same role, the
   newest stamp wins; two *equal* stamps fail loud (nothing is silently
   ignored). An undated file always loses to a dated one.
2. **Account token in the name** for account-specific exports (BSL, ST,
   BAI2): `FHB_Master`, `FHB_UTC`, `FHB_UTIA`, `Regions_UTM`, … One upload =
   one account; two different account tokens in one upload fail loud.
3. **Role keyword somewhere in the name** — this is what routes the file:

| File you're exporting | Must contain | Example |
|---|---|---|
| Bank statement lines | `BSL` (not `All_Data`) | `..._FHB_Master_BSL_UNR.xlsx` |
| System transactions | `_ST_` or `_ST.` or `Account_ST` | `..._FHB_Master_ST_UNR.xlsx` |
| MET / ORT bridge (OTBI) | `MET` or `Oracle_OTBI` | `..._MET_All_Accounts_All_Status.xlsx` |
| BAI2 bank detail | `BAI` | `..._FHB_Master_BAI2.xlsx` |
| Receivables receipts | `Receivables_Receipts` or `Receipts_All` | `20260709_Receivables_Receipts_All.xlsx` |
| Edison payments / invoices | `Edison_Payments` / `Edison_Invoices` | — |
| MID master | `MID_Master` | `UT_MID_Master_Consolidated.xlsx` |
| ORT misc / AR receipts | `ORT` + `Misc` / `ORT` + `_AR` | `ORT_Misc_All.xlsb` |
| ORT departments | `ORT_Department` | `ORT_Departments.xlsx` |
| Chart of accounts | `Chart_Of_Accounts` | `ORT_Chart_Of_Accounts.xlsx` |
| Applied/Unapplied | `Applied` or `Unapplied` | — |
| GMS sponsored aging | `GMS_001` or `Sponsored_Aging` | — |
| Lifecycle workbook | `All_Data` | `FHB_UTC_All_Data.xlsx` (Unreconcile2 only) |

## Formats

- **`.xlsx` preferred.** `.xlsb` works (needs `pyxlsb`); `.csv` works.
- Export **values, not formulas** — the engine reads computed values only.
- **Never open-and-resave a reference-bearing export through tools that
  "help"**: pandas/Excel auto-formatting can float-coerce `0006789599` into
  `6789599.0` and destroy the join key. Raw exports straight from Oracle/the
  bank are ideal.
- Leading-zero references, integer transaction numbers, and `NA`
  placeholders are all handled — no cleanup needed on your side.

## Things to avoid

- Don't put `All_Data` in a file's name unless it **is** the lifecycle
  workbook (it will be recognized and deliberately ignored by the forward
  engine).
- Don't reuse the exact same date stamp on two different exports of the same
  role in one folder — the run stops and asks you to disambiguate.
- One account per upload. The all-accounts MET export is the one exception —
  it's filtered to the run's account automatically.
- Skip exotic characters; letters, digits, `_`, `-`, `.` are safe everywhere.
