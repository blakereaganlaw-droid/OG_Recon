#!/usr/bin/env python3
"""
UT Cash Management Reconciliation Engine
========================================

A production-grade, deterministic reconciliation engine for University of
Tennessee bank accounts in Oracle Cash Management (DASH).

Two engines in one program:
  * FORWARD  — match open bank statement lines (BSL) to open system
               transactions (ST), classifying each line Match / Candidate /
               Review.
  * BACKWARD — re-audit already-reconciled groups against doctrine and
               recommend unwinding the unsound ones.

Temperament (binding, see BUILD SPEC sections 0 & 15):
  * Determinism first — same inputs, identical outputs, no float, no clock.
  * Fail loud, never silent — a missing required column/file/relationship
    raises a named exception (InvalidSourceData / MissingRequiredFile /
    AmbiguousColumn) naming the file, the role, and the candidates.
  * Every input row is accounted for exactly once.
  * The pipeline is a fixed sequence of stages; a later stage never overrides
    an earlier one; availability is re-derived inside each loop.

Dependencies: Python 3.10+ standard library + openpyxl + decimal only.
No pandas (it float-coerces zero-padded references and merchant IDs).

This module is self-contained (portable to Grok / ChatGPT / Perplexity as
either executable code or an operating procedure). The independent audit lives
in recon_audit.py and imports nothing from here.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from dataclasses import dataclass, field, asdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

# openpyxl is imported lazily inside the workbook writer / loader helpers so
# that the primitives and pure logic can be imported and unit-tested without it.


# ======================================================================
# Named exceptions (Section 0.2)
# ======================================================================

class ReconError(Exception):
    """Base class for every engine error."""


class InvalidSourceData(ReconError):
    def __init__(self, file, role, detail):
        self.file, self.role, self.detail = file, role, detail
        super().__init__(f"InvalidSourceData(file={file!r}, role={role!r}, detail={detail})")


class MissingRequiredFile(ReconError):
    def __init__(self, role, detail=""):
        self.role, self.detail = role, detail
        super().__init__(f"MissingRequiredFile(role={role!r}, detail={detail})")


class AmbiguousColumn(ReconError):
    def __init__(self, file, role, candidates):
        self.file, self.role, self.candidates = file, role, candidates
        super().__init__(
            f"AmbiguousColumn(file={file!r}, role={role!r}, candidates={candidates})"
        )


# ======================================================================
# Section 7 — Normalization primitives (implement first; unit-tested)
# ======================================================================

_MID_RE = re.compile(r"^(80\d{8}|2000\d{6})$")
_HEARTLAND_MID = "6500000097"
_SPN_RE = re.compile(r"SPN\s*\d+", re.IGNORECASE)
_DIGIT_RUN_RE = re.compile(r"\d+")
_NON_ALNUM_RE = re.compile(r"[^A-Za-z0-9]+")
_ALPHA_TOKEN_RE = re.compile(r"[A-Za-z]{4,}")

# Generic stopwords removed from payer token sets.  Explicitly includes the
# BAI2 field labels (never tokenize them as payer text) and channel words.
PAYER_STOPWORDS = {
    # BAI2 / addenda structural labels
    "COMPANY", "CUSTOMER", "ENTRY", "DESCRIPTION", "DESC", "NAME", "ID",
    "SENDING", "REF", "REFERENCE", "ADDENDA", "INFO", "INFORMATION", "NUMBER",
    "ACCOUNT", "TRACE", "BATCH", "INDIVIDUAL", "IDENTIFICATION",
    # channel / generic words
    "ACH", "WIRE", "DEPOSIT", "DEPOSITS", "PAYMENT", "PAYMENTS", "MERCHANT",
    "SERVICE", "SERVICES", "STATE", "TENNESSEE", "TENN", "UNIVERSITY",
    "CREDIT", "DEBIT", "CHECK", "MISC", "MISCELLANEOUS", "TRANSFER", "ZBA",
    "SETTLEMENT", "BANKCARD", "CARD", "FUNDS", "TRANSACTION", "TRAN", "EFT",
    "THE", "AND", "FOR", "FROM", "WITH", "INC", "LLC", "CORP", "CO",
}


def N(v) -> str:
    """Null-safe string: '' if None else str(v).strip()."""
    return "" if v is None else str(v).strip()


def cents(v):
    """Signed integer cents via Decimal.  Handles '$', thousands commas, and
    (parentheses) = negative.  Returns None if unparseable (never raises)."""
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        return int(Decimal(v) * 100)
    if isinstance(v, float):
        # Route through str to avoid binary float noise; spec forbids float
        # math but Excel hands us floats, so normalize deterministically.
        s = repr(v)
    else:
        s = str(v).strip()
    if s == "":
        return None
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()
    s = s.replace("$", "").replace(",", "").replace(" ", "")
    if s.startswith("-"):
        neg = True
        s = s[1:]
    elif s.startswith("+"):
        s = s[1:]
    if s == "":
        return None
    try:
        d = (Decimal(s) * 100).quantize(Decimal("1"))
    except (InvalidOperation, ValueError):
        return None
    c = int(d)
    return -c if neg else c


_DATE_FORMATS = ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%Y/%m/%d", "%d-%b-%Y", "%d-%b-%y")
# Excel serial epoch (1899-12-30 accounts for the Lotus 1900 leap-year bug).
_EXCEL_EPOCH = date(1899, 12, 30)


def parse_date(v):
    """Return a datetime.date, or None if unparseable.

    Accepts date, datetime (checked FIRST because datetime subclasses date),
    ISO strings incl. '...000+00:00', US formats, and Excel serial numbers.
    """
    if v is None:
        return None
    # datetime BEFORE date (Section 6 guard).
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        # Excel serial date.  Reject implausible small/large numbers.
        n = int(v)
        if 1 <= n <= 60000:
            try:
                return _EXCEL_EPOCH.fromordinal(_EXCEL_EPOCH.toordinal() + n)
            except (OverflowError, ValueError):
                return None
        return None
    s = str(v).strip()
    if s == "":
        return None
    # ISO with time / timezone, e.g. 2024-03-05T00:00:00.000+00:00
    m = re.match(r"^(\d{4}-\d{2}-\d{2})[T ]", s)
    if m:
        try:
            return date.fromisoformat(m.group(1))
        except ValueError:
            pass
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    # Bare numeric string could be an Excel serial exported as text.
    if s.isdigit():
        n = int(s)
        if 1 <= n <= 60000:
            return _EXCEL_EPOCH.fromordinal(_EXCEL_EPOCH.toordinal() + n)
    return None


def znorm(s) -> str:
    """Uppercase, strip non-alphanumerics (for reference equality)."""
    return _NON_ALNUM_RE.sub("", N(s)).upper()


def digit_runs(s, n=5) -> set:
    """Set of maximal digit substrings of length >= n (for reference-tie)."""
    return {r for r in _DIGIT_RUN_RE.findall(N(s)) if len(r) >= n}


def payer_tokens(s) -> set:
    """Set of alpha tokens length>=4, minus the generic stopword set."""
    return {t.upper() for t in _ALPHA_TOKEN_RE.findall(N(s))
            if t.upper() not in PAYER_STOPWORDS}


def spn_of(s) -> str:
    """The SPN token (regex (?i)SPN\\s*\\d+) found in a receipt/transaction
    number, else ''.  Normalized to 'SPN' + digits (no internal space)."""
    m = _SPN_RE.search(N(s))
    if not m:
        return ""
    return "SPN" + re.sub(r"\D", "", m.group(0))


def is_mid(s) -> bool:
    """True if the znorm digits match the MID pattern and != Heartland id."""
    z = znorm(s)
    if z == _HEARTLAND_MID:
        return False
    return bool(_MID_RE.match(z))


def sibling(a, b) -> bool:
    """True if a,b are equal-length pure-numeric of length>=7 differing only
    in the last 1-2 digits (a conflict, never a tie)."""
    za, zb = znorm(a), znorm(b)
    if not (za.isdigit() and zb.isdigit()):
        return False
    if len(za) != len(zb) or len(za) < 7:
        return False
    if za == zb:
        return False
    # Prefixes must match up to the last two digits.
    if za[:-2] != zb[:-2]:
        return False
    # And they differ somewhere in the final 1-2 positions (guaranteed since
    # za != zb and prefixes to -2 are equal).
    return True


def signed_lag(bsl_date, st_date):
    """(BSL date - ST date).days, or None if either is missing."""
    if bsl_date is None or st_date is None:
        return None
    return (bsl_date - st_date).days


# ---- Reference equality / tie (Section 7 trailer) --------------------

def reference_equal(a, b) -> bool:
    """znorm equality OR full containment of one znorm token inside the other
    for length >= 6.  A sibling pair is a conflict, never equal."""
    za, zb = znorm(a), znorm(b)
    if not za or not zb:
        return False
    if sibling(a, b):
        return False
    if za == zb:
        return True
    if len(za) >= 6 and za in zb:
        return True
    if len(zb) >= 6 and zb in za:
        return True
    return False


def reference_tie(a, b) -> bool:
    """Weaker corroboration: a shared digit_run of length >= 5 (truncation
    allowed).  A sibling pair never ties."""
    if sibling(a, b):
        return False
    ra, rb = digit_runs(a, 5), digit_runs(b, 5)
    if ra & rb:
        return True
    # Truncation allowed: one run is a prefix of the other (length >= 5).
    for x in ra:
        for y in rb:
            if len(x) >= 5 and len(y) >= 5 and (x.startswith(y) or y.startswith(x)):
                return True
    return False


# ======================================================================
# Section 4 — File router.  Identify every file by NAME, then bind by header.
# ======================================================================

@dataclass
class RouterRule:
    role: str
    contains: list          # all tokens must appear (case-insensitive)
    excludes: list          # none may appear
    any_of: list            # at least one must appear (empty = no constraint)
    sheet: str
    required: bool


# First matching rule wins (Section 4.1).  `any_of` handles "several filename
# shapes"; `excludes` handles the "not All_Data" style negative constraints.
ROUTER_TABLE = [
    RouterRule("BSL", ["bsl"], ["all_data"], [], "first", True),
    RouterRule("ST", [], [], ["_st", "account_st"], "first", False),
    RouterRule("ALL_DATA", ["all_data"], [], [], "multi", False),
    RouterRule("MET", ["met"], [], [], "first", False),
    RouterRule("RECEIPTS", [], [], ["receivables_receipts", "receipts_all"], "Export to Excel", False),
    RouterRule("ORT_AR", ["ort", "ar"], [], [], "Report", False),
    RouterRule("ORT_MISC", ["ort", "misc"], [], [], "Report", False),
    RouterRule("BAI2", ["bai"], [], [], "first", False),
    RouterRule("EDISON_PAY", ["edison_payments"], [], [], "first", False),
    RouterRule("EDISON_INV", ["edison_invoices"], [], [], "first", False),
    RouterRule("MID_MASTER", ["mid_master"], [], [], "all", False),
    RouterRule("ENRICHED", ["enriched", "crossref"], [], [], "first", False),
    RouterRule("APPLIED_UNAPPLIED", ["applied", "unapplied"], [], [], "first", False),
    RouterRule("CONTRACTS_INV", ["contracts_to_receivable_invoices"], [], [], "first", False),
    RouterRule("GMS_AGING", [], [], ["gms_001", "sponsored_aging"], "first", False),
    RouterRule("AR_INVOICES", ["ar_invoices"], [], [], "first", False),
    RouterRule("AR_MATCHED", [], [], ["ar_matched", "deposit_receipts"], "first", False),
    RouterRule("AR_UNAPPLIED_SUMMARY", [], [], ["ar_063", "unapplied_receipts_summary"], "first", False),
    RouterRule("DEPT_INFO", ["ort_department_info"], [], [], "first", False),
    RouterRule("GMS_SPONSOR_MAP", ["rpt_gms_00"], [], [], "first", False),
    RouterRule("CFG_MATCHING", ["matching_rules"], [], [], "first", False),
    RouterRule("CFG_PARSE", ["parse_rules"], [], [], "first", False),
    RouterRule("CFG_TOLERANCE", ["tolerance_rules"], [], [], "first", False),
    RouterRule("CFG_RULESETS", ["recon_rulesets"], [], [], "first", False),
    RouterRule("CFG_TCR", ["transaction_creation_rules"], [], [], "first", False),
    RouterRule("RELATIONSHIP_MAP", [], [], ["relationship_map", "rosetta"], "reference", False),
]

# Roles the pipeline treats as hard requirements at the top level (Section 4.1
# "Required?" column reads "yes"/conditional).  BSL is the sole unconditional
# requirement; ST is required only when no ALL_DATA is present.
HARD_REQUIRED_ROLES = {"BSL"}

# Account tokens recognized in filenames (Section 1.2 + 4.1).
_ACCOUNT_TOKENS = [
    ("FHB_MASTER", ["fhb", "master"]),
    ("FHB_UTHSC", ["fhb", "uthsc"]),
    ("FHB_UTIA", ["fhb", "utia"]),
    ("FHB_UTC", ["fhb", "utc"]),
    ("FHB_UTM", ["fhb", "utm"]),
    ("FHB_UTSO", ["fhb", "utso"]),
    ("FHB_AP", ["fhb", "ap"]),
    ("REGIONS_UTIA", ["regions", "utia"]),
    ("REGIONS_UTIPS", ["regions", "utips"]),
    ("REGIONS_UTM", ["regions", "utm"]),
]

_YYYYMMDD_RE = re.compile(r"(\d{8})")


def _tokens_present(fname_lower, tokens):
    return all(t in fname_lower for t in tokens)


def classify_file(filename) -> str | None:
    """Return the internal role key for a filename, or None if unmatched.
    First matching router rule wins (Section 4.1)."""
    low = filename.lower()
    for rule in ROUTER_TABLE:
        if rule.contains and not _tokens_present(low, rule.contains):
            continue
        if rule.excludes and any(t in low for t in rule.excludes):
            continue
        if rule.any_of and not any(t in low for t in rule.any_of):
            continue
        if not rule.contains and not rule.any_of:
            continue  # a rule with no positive tokens matches nothing
        return rule.role
    return None


def infer_account(filename) -> str | None:
    """Normalize the account token from a BSL/ALL_DATA filename to short form."""
    low = filename.lower()
    for short, tokens in _ACCOUNT_TOKENS:
        if _tokens_present(low, tokens):
            return short
    return None


def _leading_date_key(filename):
    m = _YYYYMMDD_RE.search(filename)
    return m.group(1) if m else "00000000"


@dataclass
class RoutedFile:
    role: str
    path: str
    filename: str
    sheet_hint: str


def route_folder(input_dir) -> dict:
    """Scan input_dir, classify each file, resolve multi-file-per-role by
    newest-YYYYMMDD (union+dedup handled downstream).  Returns
    {role: [RoutedFile, ...]} plus asserts hard-required roles present."""
    if not os.path.isdir(input_dir):
        raise MissingRequiredFile("INPUT_DIR", f"{input_dir} is not a directory")
    by_role = {}
    rule_by_role = {r.role: r for r in ROUTER_TABLE}
    for name in sorted(os.listdir(input_dir)):
        path = os.path.join(input_dir, name)
        if not os.path.isfile(path):
            continue
        if not name.lower().endswith((".xlsx", ".xlsm", ".xlsb", ".csv")):
            continue
        role = classify_file(name)
        if role is None:
            continue
        rf = RoutedFile(role, path, name, rule_by_role[role].sheet)
        by_role.setdefault(role, []).append(rf)

    # Newest-wins ordering within a role (leading YYYYMMDD desc); ties keep all
    # for union+dedup by the loaders.
    for role, files in by_role.items():
        files.sort(key=lambda f: _leading_date_key(f.filename), reverse=True)

    for role in HARD_REQUIRED_ROLES:
        if role not in by_role:
            raise MissingRequiredFile(role, f"no file in {input_dir} matched role {role}")
    return by_role


# ======================================================================
# Section 5 — Column binding (content first, header as tiebreak)
# ======================================================================

@dataclass
class RoleSpec:
    required: bool
    header_aliases: list
    predicate: object  # fn(cell) -> bool


def _norm_header(h) -> str:
    return re.sub(r"[^A-Z0-9]", "", N(h).upper())


# ---- content predicates (Section 5.2) --------------------------------

def pred_date(v):
    return parse_date(v) is not None


def pred_signed_amount(v):
    return cents(v) is not None


def pred_reference(v):
    s = N(v)
    if not s:
        return False
    if parse_date(v) is not None:
        return False
    return bool(re.search(r"[A-Za-z0-9]", s))


_STATUS_VOCAB = {
    "REC", "UNR", "VOID", "APP", "CLEARED", "CONFIRMED", "REMITTED",
    "UNAPPLIED", "REVERSED", "OP", "CL", "OPEN", "CLOSED", "APPLIED", "MATCHED",
}


def pred_status(v):
    return _norm_header(v) in _STATUS_VOCAB


def pred_customer(v):
    s = N(v)
    if not s or not re.search(r"[A-Za-z]", s):
        return False
    digits = sum(c.isdigit() for c in s)
    return digits / max(1, len(s)) < 0.4


_TXN_TYPE_VOCAB = {
    "ACH", "CHK", "MSC", "EFT", "BKA", "BKF", "ZBA", "CREDITCARD", "CHECK",
    "MISCELLANEOUS", "WIRE", "DEPOSIT", "CREDIT", "DEBIT",
}


def pred_txn_type(v):
    return _norm_header(v) in _TXN_TYPE_VOCAB


_GL_RE = re.compile(r"^\d{2}-\d{7}-\d{6}-\d{6}")


def pred_gl_string(v):
    return bool(_GL_RE.match(N(v)))


def pred_mid(v):
    return is_mid(v)


def pred_number(v):
    s = N(v)
    return bool(s) and bool(re.search(r"\d", s))


def pred_any(v):  # non-empty
    return N(v) != ""


def bind_columns(rows, role_specs, filename="<rows>", header_scan=12, sample=50):
    """Bind each role to a column index by scanning content first, header as
    tiebreak (Section 5.1).  Returns (mapping {role: col_index}, header_index).

    rows: list of tuples already read from the sheet.
    """
    if not rows:
        # Every required role fails.
        for role, spec in role_specs.items():
            if spec.required:
                raise InvalidSourceData(filename, role, "empty sheet")
        return {}, 0

    all_aliases = []
    for spec in role_specs.values():
        all_aliases.extend(_norm_header(a) for a in spec.header_aliases)
    max_arity = len(role_specs)

    # 1. Locate the header row.
    header_index = 0
    best_hits = -1
    scan = min(header_scan, len(rows))
    for i in range(scan):
        row = rows[i]
        nonempty = sum(1 for c in row if N(c) != "")
        hits = 0
        for c in row:
            hc = _norm_header(c)
            if hc and any(hc == a or (a and a in hc) for a in all_aliases):
                hits += 1
        # Prefer a row with enough columns AND >=2 alias hits.
        if nonempty >= min(max_arity, 2) and hits >= 2 and hits > best_hits:
            best_hits = hits
            header_index = i
    header = rows[header_index]
    ncols = max(len(r) for r in rows)
    data_rows = rows[header_index + 1: header_index + 1 + sample]

    # 2/3. Score every column for every role; content decides, header breaks ties.
    mapping = {}
    for role, spec in role_specs.items():
        alias_norms = [_norm_header(a) for a in spec.header_aliases]
        scored = []
        for col in range(ncols):
            hcell = header[col] if col < len(header) else ""
            hnorm = _norm_header(hcell)
            if hnorm and hnorm in alias_norms:
                header_score = 3
            elif hnorm and any(a and a in hnorm for a in alias_norms):
                header_score = 2
            else:
                header_score = 0
            sampled = [r[col] for r in data_rows if col < len(r) and N(r[col]) != ""]
            if sampled:
                content_score = sum(1 for c in sampled if spec.predicate(c)) / len(sampled)
            else:
                content_score = 0.0
            scored.append((content_score, header_score, col))
        scored.sort(key=lambda t: (t[0], t[1]), reverse=True)
        best = scored[0]
        # Ambiguity: tie on both content and header for a required role.
        if spec.required and len(scored) > 1:
            second = scored[1]
            if (best[0], best[1]) == (second[0], second[1]) and best[0] > 0:
                raise AmbiguousColumn(filename, role, [best[2], second[2]])
        if spec.required and best[0] == 0:
            raise InvalidSourceData(
                filename, role,
                f"no column satisfied content predicate (header={_norm_header(header[best[2]]) if best[2] < len(header) else ''!r})")
        mapping[role] = best[2]
    return mapping, header_index


# ---- Role specs per file (Section 5.3) -------------------------------

def _rs(required, aliases, pred):
    return RoleSpec(required, aliases, pred)


BSL_ROLES = {
    "date": _rs(True, ["Transaction Date", "Booking Date", "Value Date", "Date", "Post Date"], pred_date),
    "amount": _rs(True, ["Amount", "Transaction Amount", "Signed Amount"], pred_signed_amount),
    "line_key": _rs(True, ["Line Number", "Statement Line", "Line", "Bank Statement Line", "Sequence"], pred_number),
    "reference": _rs(False, ["Reference", "Bank Reference", "Recon Reference"], pred_reference),
    "account_servicer_reference": _rs(False, ["Account Servicer Reference", "Servicer Reference"], pred_reference),
    "customer_reference": _rs(False, ["Customer Reference"], pred_reference),
    "additional_info": _rs(False, ["Additional Information", "Additional Info", "Addenda"], pred_any),
    "transaction_type": _rs(False, ["Transaction Type", "Type"], pred_any),
    "transaction_code": _rs(False, ["Transaction Code", "Bank Transaction Code", "Code"], pred_any),
}

ST_ROLES = {
    "date": _rs(True, ["Transaction Date", "Date", "Accounting Date", "GL Date"], pred_date),
    "amount": _rs(True, ["Amount", "Transaction Amount", "Signed Amount"], pred_signed_amount),
    "transaction_number": _rs(True, ["Transaction Number", "Trx Number", "Transaction Num", "Number"], pred_reference),
    "source": _rs(True, ["Source", "Transaction Source", "Origin"], pred_any),
    "reference": _rs(False, ["Reference", "Recon Match Reference", "Match Reference"], pred_reference),
    "structured_payment_reference": _rs(False, ["Structured Payment Reference", "Strc Pay Ref"], pred_reference),
    "counterparty": _rs(False, ["Counterparty", "Customer", "Payer", "Customer Name"], pred_customer),
    "transaction_type": _rs(False, ["Transaction Type", "Type"], pred_any),
}

RECEIPTS_ROLES = {
    "receipt_number": _rs(True, ["Receipt Number", "Receipt Num", "Receipt"], pred_reference),
    "status": _rs(True, ["Status", "Receipt Status", "State"], pred_status),
    "amount": _rs(True, ["Receipt Amount", "Entered Amount", "Amount"], pred_signed_amount),
    "customer_name": _rs(True, ["Customer Name", "Customer", "Payer"], pred_customer),
    "receipt_date": _rs(False, ["Receipt Date", "Date"], pred_date),
    "deposit_date": _rs(False, ["Deposit Date", "Cleared Date"], pred_date),
    "batch_number": _rs(False, ["Batch Number", "Receipt Batch"], pred_reference),
    "reference": _rs(False, ["Reference", "Remittance Reference"], pred_reference),
    "unapplied_amount": _rs(False, ["Unapplied Amount"], pred_signed_amount),
}

MET_ROLES = {
    "trx_id": _rs(True, ["CET Transaction ID", "Transaction Number", "Trx Id", "Transaction Id"], pred_reference),
    "amount": _rs(True, ["Amount", "Transaction Amount"], pred_signed_amount),
    "transaction_date": _rs(True, ["Transaction Date", "Date"], pred_date),
    "status": _rs(True, ["Status", "State"], pred_any),
    "description": _rs(True, ["CET Description", "Description", "Desc"], pred_any),
    "cleared_date": _rs(False, ["Cleared Date"], pred_date),
    "offset": _rs(False, ["Offset Concatenated Segments", "Offset", "GL"], pred_any),
}

RECON_HISTORY_ROLES = {
    "recon_grp": _rs(True, ["Recon Grp", "Recon Group", "Reconciliation Group"], pred_reference),
    "amount": _rs(True, ["Amount"], pred_signed_amount),
    "auto_flag": _rs(True, ["Auto Recon Flag", "Auto Recon", "Auto Flag"], pred_any),
    "rule_name": _rs(True, ["Rule Name", "Matching Rule Name"], pred_any),
    "match_type": _rs(True, ["Match Type"], pred_any),
    "type_match": _rs(True, ["Type Match"], pred_any),
    "amount_match": _rs(True, ["Amount Match"], pred_any),
    "date_match": _rs(True, ["Date Match"], pred_any),
    "ref_match": _rs(True, ["Ref Match", "Reference Match"], pred_any),
    "recon_src": _rs(False, ["Recon Src", "Recon Source"], pred_any),
}

BANK_STATEMENT_LINES_ROLES = {
    "date": _rs(True, ["Transaction Date", "Booking Date", "Date"], pred_date),
    "amount": _rs(True, ["Amount", "Transaction Amount"], pred_signed_amount),
    "line_key": _rs(True, ["Line Number", "Statement Line", "Line", "Sequence"], pred_number),
    "rec_status": _rs(False, ["Rec Status", "Reconciliation Status", "Status"], pred_any),
    "rec_grp": _rs(False, ["Rec Grp", "Recon Grp", "Rec Group"], pred_reference),
    "reference": _rs(False, ["Reference", "Account Servicer Reference"], pred_reference),
    "additional_info": _rs(False, ["Additional Information", "Additional Info"], pred_any),
    "matching_rule_name": _rs(False, ["Matching Rule Name", "Rule Name"], pred_any),
    "match_type": _rs(False, ["Match Type"], pred_any),
}

MISC_RECEIPTS_ROLES = {
    "rec_num": _rs(True, ["Rec Num", "Receipt Number", "Receipt Num"], pred_reference),
    "rec_amnt": _rs(True, ["Rec Amnt", "Receipt Amount", "Amount"], pred_signed_amount),
    "rec_status": _rs(False, ["Rec Status", "Status"], pred_any),
    "rec_ref": _rs(False, ["Rec Ref", "Reference"], pred_reference),
    "rec_grp": _rs(False, ["Rec Grp", "Recon Grp"], pred_reference),
    "dep_num": _rs(False, ["Dep Num", "Deposit Number"], pred_reference),
    "offset": _rs(False, ["Offset", "GL"], pred_any),
    "recommended_bank_line": _rs(False, ["Recommended Bank Line"], pred_any),
}

EDISON_PAY_ROLES = {
    "reference": _rs(True, ["Reference", "Payment Reference"], pred_reference),
    "invoice_number": _rs(True, ["Invoice Number", "Invoice"], pred_reference),
    "payment_date": _rs(False, ["Payment Date", "Date"], pred_date),
    "amount": _rs(True, ["Amount", "Payment Amount", "Payment Total"], pred_signed_amount),
}

EDISON_INV_ROLES = {
    "invoice_number": _rs(True, ["Invoice Number", "Invoice"], pred_reference),
    "gross_amount": _rs(True, ["Gross Amount", "Gross", "Amount"], pred_signed_amount),
}

APPLIED_UNAPPLIED_ROLES = {
    "trx_number": _rs(True, ["Trx Number", "Transaction Number"], pred_reference),
    "receipt_number": _rs(True, ["Receipt Number"], pred_reference),
    "contract_number": _rs(False, ["Contract Number"], pred_reference),
    "customer_name": _rs(False, ["Customer Name", "Customer"], pred_customer),
    "accounted_applied_amount": _rs(False, ["Accounted Applied Amount", "Applied Amount"], pred_signed_amount),
    "accounted_unapplied_amount": _rs(False, ["Accounted Unapplied Amount", "Unapplied Amount"], pred_signed_amount),
}


# ======================================================================
# Workbook / CSV loading (Section 3)
# ======================================================================

def _read_sheet_rows(path, sheet_hint):
    """Read rows from an xlsx sheet (data_only, read_only=False per Section 3).
    Returns (rows, sheet_title).  sheet_hint 'first' or a title substring."""
    from openpyxl import load_workbook
    try:
        wb = load_workbook(path, read_only=False, data_only=True)
    except Exception as e:
        raise InvalidSourceData(os.path.basename(path), "WORKBOOK", f"load failed: {e}")
    ws = _pick_sheet(wb, sheet_hint)
    rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
    title = ws.title
    wb.close()
    return rows, title


def _pick_sheet(wb, sheet_hint):
    if sheet_hint in ("first", "multi", "all", "reference", "stream", None):
        return wb.worksheets[0]
    # substring match (case-insensitive)
    low = sheet_hint.lower()
    for ws in wb.worksheets:
        if low in ws.title.lower():
            return ws
    return wb.worksheets[0]


def _read_csv_rows(path):
    import csv
    with open(path, newline="", encoding="utf-8", errors="replace") as fh:
        rows = list(csv.reader(fh))
    if rows:
        # Strip leading UTF-8 BOM from the first header cell.
        first = rows[0]
        if first and first[0].startswith("﻿"):
            first[0] = first[0].lstrip("﻿")
    return [tuple(r) for r in rows]


def read_rows(routed_file: RoutedFile):
    """Dispatch to xlsx or csv loader based on extension.  Wraps failures in
    InvalidSourceData (Section 15.5)."""
    ext = os.path.splitext(routed_file.path)[1].lower()
    try:
        if ext == ".csv":
            return _read_csv_rows(routed_file.path), "csv"
        return _read_sheet_rows(routed_file.path, routed_file.sheet_hint)
    except InvalidSourceData:
        raise
    except Exception as e:
        raise InvalidSourceData(routed_file.filename, routed_file.role, f"read failed: {e}")


def read_named_sheet(path, title_substr, role_specs):
    """Read a specific sheet by title substring (for ALL_DATA multi-sheet).
    Returns (rows, mapping, header_index) or (None, None, None) if absent."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=False, data_only=True)
    low = title_substr.lower()
    target = None
    for ws in wb.worksheets:
        if low in ws.title.lower():
            target = ws
            break
    if target is None:
        wb.close()
        return None, None, None
    rows = [tuple(r) for r in target.iter_rows(values_only=True)]
    wb.close()
    mapping, hi = bind_columns(rows, role_specs, filename=os.path.basename(path))
    return rows, mapping, hi


# ======================================================================
# Section 8 — Candidate ST pool (all-status, deduped)
# ======================================================================

# Status vocabularies for availability (Section 8.2).
OPEN_RECEIPT_STATUSES = {"UNR", "REMITTED", "CONFIRMED", "UNAPPLIED", "OPEN", "OP"}
CLOSED_RECEIPT_STATUSES = {"CLEARED", "APP", "APPLIED", "REVERSED", "VOID", "REC", "CLOSED", "CL"}
JOURNAL_SOURCES = {"GL", "JOURNAL", "JE", "MANUAL"}


@dataclass
class PoolEntry:
    id: str
    amount_cents: int
    date: object            # datetime.date or None
    reference: str
    znref: str
    digits: set
    payer_tokens: set
    counterparty: str
    source: str             # AR / EXT / AP / PAY / GL ...
    status: str
    available: bool
    spn: str
    is_mid: bool
    origin: str = ""        # which file/role produced it (for run log)
    deposit_id: str = ""    # ORT d:
    receipt_id: str = ""    # ORT r:


def _mk_entry(id, amount_cents, dt, reference, counterparty, source, status,
              available, origin, deposit_id="", receipt_id=""):
    ref = N(reference)
    return PoolEntry(
        id=N(id),
        amount_cents=amount_cents,
        date=dt,
        reference=ref,
        znref=znorm(ref),
        digits=digit_runs(ref),
        payer_tokens=payer_tokens(counterparty) | payer_tokens(reference),
        counterparty=N(counterparty),
        source=N(source).upper(),
        status=_norm_header(status),
        available=available,
        spn=spn_of(id) or spn_of(reference),
        is_mid=is_mid(reference),
        origin=origin,
        deposit_id=N(deposit_id),
        receipt_id=N(receipt_id),
    )


def _dedup_keep_largest(entries, keyfn):
    """Dedup by keyfn keeping the largest-magnitude amount (the total, not a
    split).  When the kept total lacks a counterparty, borrow it from a dropped
    split (Section 8.1)."""
    groups = {}
    for e in entries:
        groups.setdefault(keyfn(e), []).append(e)
    out = []
    for key, members in groups.items():
        members_sorted = sorted(
            members,
            key=lambda e: (abs(e.amount_cents), e.id),
            reverse=True,
        )
        keep = members_sorted[0]
        if not keep.counterparty:
            for m in members_sorted[1:]:
                if m.counterparty:
                    keep.counterparty = m.counterparty
                    keep.payer_tokens = keep.payer_tokens | payer_tokens(m.counterparty)
                    break
        out.append(keep)
    return out


def build_pool(loaded: dict, account: str, runlog: dict) -> list:
    """Build the union pool from loaded data sources (Section 8).

    `loaded` maps role -> {'rows', 'map', 'header_index'} for bound files.
    Returns a list of PoolEntry.  Records pool sizes in runlog.
    """
    pool = []
    counts = {}

    # 1. Open non-Receivables STs from ST (exclude Journal from eligibility but
    #    keep for backward engine; dedup keep-largest by transaction number).
    st = loaded.get("ST")
    if st:
        rows, m, hi = st["rows"], st["map"], st["header_index"]
        raw = []
        for r in rows[hi + 1:]:
            amt = cents(_cell(r, m.get("amount")))
            if amt is None:
                continue
            source = N(_cell(r, m.get("source"))).upper()
            txn = N(_cell(r, m.get("transaction_number")))
            if not txn:
                continue
            src_norm = _classify_source(source)
            e = _mk_entry(
                id=txn,
                amount_cents=amt,
                dt=parse_date(_cell(r, m.get("date"))),
                reference=_cell(r, m.get("reference")) or _cell(r, m.get("structured_payment_reference")),
                counterparty=_cell(r, m.get("counterparty")),
                source=src_norm,
                status="UNR",
                available=(src_norm not in JOURNAL_SOURCES),  # journals kept but ineligible
                origin="ST",
            )
            # Do not add Receivables rows from the ST export (Section 8.2).
            if src_norm == "AR":
                continue
            raw.append(e)
        deduped = _dedup_keep_largest(raw, lambda e: e.id)
        pool.extend(deduped)
        counts["ST"] = len(deduped)

    # 2. All-status Receivables receipts from RECEIPTS (authoritative).
    rc = loaded.get("RECEIPTS")
    if rc:
        rows, m, hi = rc["rows"], rc["map"], rc["header_index"]
        raw = []
        for r in rows[hi + 1:]:
            amt = cents(_cell(r, m.get("amount")))
            recno = N(_cell(r, m.get("receipt_number")))
            if amt is None or not recno:
                continue
            status = _norm_header(_cell(r, m.get("status")))
            available = status in OPEN_RECEIPT_STATUSES
            e = _mk_entry(
                id=recno,
                amount_cents=amt,
                dt=parse_date(_cell(r, m.get("receipt_date")) or _cell(r, m.get("deposit_date"))),
                reference=_cell(r, m.get("reference")) or recno,
                counterparty=_cell(r, m.get("customer_name")),
                source="AR",
                status=status or "UNR",
                available=available,
                origin="RECEIPTS",
            )
            raw.append(e)
        deduped = _dedup_keep_largest(raw, lambda e: e.id)
        pool.extend(deduped)
        counts["RECEIPTS"] = len(deduped)

    # 3. ORT receipts from MET for the ECT chain (index by d: and reference).
    met = loaded.get("MET")
    if met:
        rows, m, hi = met["rows"], met["map"], met["header_index"]
        raw = []
        for r in rows[hi + 1:]:
            amt = cents(_cell(r, m.get("amount")))
            trx = N(_cell(r, m.get("trx_id")))
            if amt is None or not trx:
                continue
            desc = N(_cell(r, m.get("description")))
            dep_id, rec_id, _payer = parse_met_description(desc)
            cleared = parse_date(_cell(r, m.get("cleared_date")))
            status_raw = _norm_header(_cell(r, m.get("status")))
            # cleared_date present => closed (Section 8.3).
            closed = cleared is not None or status_raw in CLOSED_RECEIPT_STATUSES
            e = _mk_entry(
                id=trx,
                amount_cents=amt,
                dt=parse_date(_cell(r, m.get("transaction_date"))),
                reference=_met_reference(desc) or trx,
                counterparty=_payer,
                source="EXT",
                status=status_raw or ("REC" if closed else "UNR"),
                available=not closed,
                origin="MET",
                deposit_id=dep_id,
                receipt_id=rec_id,
            )
            raw.append(e)
        deduped = _dedup_keep_largest(raw, lambda e: e.id)
        pool.extend(deduped)
        counts["MET"] = len(deduped)

    runlog["pool_sizes"] = counts
    runlog["pool_total"] = len(pool)
    # Status breakdown for the run log.
    by_status = {}
    for e in pool:
        by_status[e.status] = by_status.get(e.status, 0) + 1
    runlog["pool_by_status"] = by_status
    return pool


def _classify_source(source):
    """Normalize a raw source string to one of AR/EXT/AP/PAY/GL (Section 8)."""
    s = _norm_header(source)
    if not s:
        return "EXT"
    if "AR" in s or "RECEIV" in s:
        return "AR"
    if "AP" in s or "PAYABLE" in s:
        return "AP"
    if "PAY" in s:
        return "PAY"
    if "JOURNAL" in s or s == "GL" or "MANUAL" in s or s == "JE":
        return "GL"
    if "EXT" in s or "ECT" in s or "ORT" in s:
        return "EXT"
    return "EXT"


def _cell(row, idx):
    if idx is None or idx >= len(row):
        return None
    return row[idx]


# ---- MET description parser (Section 5.4) ----------------------------

def parse_met_description(desc):
    """Split CET_DESCRIPTION on '|'.  seg0->d: deposit id; seg1->r: receipt id;
    seg2+ joined -> payer/purpose.  Returns (deposit_id, receipt_id, payer)."""
    parts = [p.strip() for p in N(desc).split("|")]
    dep = rec = ""
    if len(parts) >= 1:
        m = re.search(r"d:\s*(\d+)", parts[0], re.IGNORECASE)
        if m:
            dep = m.group(1)
    if len(parts) >= 2:
        m = re.search(r"r:\s*(\d+)", parts[1], re.IGNORECASE)
        if m:
            rec = m.group(1)
    payer = " | ".join(parts[2:]).strip() if len(parts) > 2 else ""
    # Fallback: search whole string for d:/r: if the split missed them.
    if not dep:
        m = re.search(r"d:\s*(\d+)", N(desc), re.IGNORECASE)
        if m:
            dep = m.group(1)
    if not rec:
        m = re.search(r"r:\s*(\d+)", N(desc), re.IGNORECASE)
        if m:
            rec = m.group(1)
    return dep, rec, payer


def _met_reference(desc):
    """Best-effort reference extracted from the payer/purpose segment."""
    _, _, payer = parse_met_description(desc)
    return payer


# ======================================================================
# BSL model + lane classification (Section 9, P1)
# ======================================================================

LANE_STATE = "STATE"
LANE_MERCHANT = "MERCHANT"
LANE_GENERAL = "GENERAL"

# FHB transaction codes whose whole Account Servicer Reference is the match key.
FHB_WHOLE_REF_CODES = {"142", "174", "175", "165", "244", "495", "451", "699",
                       "475", "357", "631", "661"}

_MERCHANT_KEYWORDS = ("MERCHANT SERVICE", "BANKCARD", "TOUCHNET", "CYBERSOURCE",
                      "PAYMENTECH")
_STATE_KEYWORDS = ("STATE-TN", "STATE OF TENN")


@dataclass
class BSL:
    line_key: str
    date: object
    amount_cents: int
    recon_reference: str
    reference_raw: str
    additional_info: str
    transaction_type: str
    transaction_code: str
    lane: str = LANE_GENERAL
    ref_digits: set = field(default_factory=set)
    payer_tokens: set = field(default_factory=set)
    mid: str = ""
    line_info: str = ""


def build_recon_reference(reference, account_servicer_reference, transaction_code):
    """RECON_REFERENCE is the whole Account Servicer Reference for the FHB
    whole-ref transaction codes (Parse Rules (X~)); fall back to reference."""
    asr = N(account_servicer_reference)
    code = N(transaction_code)
    if code in FHB_WHOLE_REF_CODES and asr:
        return asr
    if asr:
        return asr
    return N(reference)


def extract_ach_payer(additional_info):
    """Parse ACH payer from addenda 'SENDING CO NAME: <X> ENTRY DESC' and from
    'Company Name: <X>' in an enriched addenda blob.  Never tokenize BAI2
    labels as payer text."""
    s = N(additional_info)
    m = re.search(r"SENDING\s+CO\s+NAME:\s*(.+?)\s+ENTRY\s+DESC", s, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    m = re.search(r"Company\s+Name:\s*(.+?)(?:$|\s{2,}|\|)", s, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return ""


def _mid_from_text(additional_info):
    """Extract a Customer ID MID from addenda if present."""
    for tok in re.findall(r"\d{10}", N(additional_info)):
        if is_mid(tok):
            return znorm(tok)
    return ""


def classify_lane(bsl: BSL) -> str:
    """P1 lane selector: STATE before MERCHANT before GENERAL."""
    info = bsl.additional_info.upper()
    ref = bsl.recon_reference.upper()
    # STATE
    if info.startswith("STATE-TN") or any(k in info for k in _STATE_KEYWORDS):
        return LANE_STATE
    # MERCHANT
    if bsl.mid or is_mid(bsl.recon_reference):
        return LANE_MERCHANT
    if any(k in info for k in _MERCHANT_KEYWORDS) or any(k in ref for k in _MERCHANT_KEYWORDS):
        return LANE_MERCHANT
    return LANE_GENERAL


def make_bsl(line_key, dt, amount_cents, reference, account_servicer_reference,
             additional_info, transaction_type, transaction_code):
    recon_ref = build_recon_reference(reference, account_servicer_reference, transaction_code)
    ach_payer = extract_ach_payer(additional_info)
    mid = _mid_from_text(additional_info) or (znorm(recon_ref) if is_mid(recon_ref) else "")
    bsl = BSL(
        line_key=N(line_key),
        date=dt,
        amount_cents=amount_cents,
        recon_reference=recon_ref,
        reference_raw=N(reference),
        additional_info=N(additional_info),
        transaction_type=N(transaction_type),
        transaction_code=N(transaction_code),
        ref_digits=digit_runs(recon_ref),
        payer_tokens=payer_tokens(recon_ref) | payer_tokens(ach_payer) | payer_tokens(additional_info),
        mid=mid,
    )
    bsl.lane = classify_lane(bsl)
    # Human-readable line info (no formulas; static).
    info = N(additional_info) or N(reference) or N(transaction_type)
    bsl.line_info = f"{bsl.line_key} {info}".strip()[:200]
    return bsl


# ======================================================================
# Date doctrine (Section 11)
# ======================================================================

DATE_CEILING = 15


def date_band(lag):
    """Non-State band label for a signed lag (BSL - ST)."""
    if lag is None:
        return "UNKNOWN"
    a = abs(lag)
    if a <= 3:
        return "STRONG"
    if a <= 7:
        return "MODERATE"
    if a <= 15:
        return "WEAK"
    if a <= 30:
        return "SUSPICIOUS"
    return "REJECT"


def date_ok_general(lag):
    return lag is not None and abs(lag) <= DATE_CEILING


def date_ok_state(lag):
    """State exception: no ceiling when BSL precedes ST (lag <= 0); demote when
    ST precedes BSL by more than 20 days (lag > 20)."""
    if lag is None:
        return False
    if lag <= 0:
        return True
    return lag <= 20


def date_ok_merchant(lag):
    """Card window: ST precedes BSL by 1-4 days (lag in 1..4)."""
    return lag is not None and 1 <= lag <= 4


# ======================================================================
# Consumption ledger (Section 2.7, 8.4)
# ======================================================================

class Ledger:
    """One consumption ledger spanning Matches and Candidates.  An ST named in
    a Candidate is barred from later Match promotion."""

    def __init__(self):
        self._consumed = set()       # matched (hard-consumed)
        self._candidate_barred = set()

    def is_available(self, entry: PoolEntry):
        return entry.available and entry.id not in self._consumed and entry.id not in self._candidate_barred

    def consume_match(self, ids):
        for i in ids:
            self._consumed.add(i)

    def consume_candidate(self, ids):
        for i in ids:
            self._candidate_barred.add(i)

    def consumed(self):
        return set(self._consumed)

    def barred(self):
        return set(self._candidate_barred)


# ======================================================================
# Classification result model
# ======================================================================

MATCH = "Match"
CANDIDATE = "Candidate"
REVIEW = "Review"

CONF_HIGH = "High"
CONF_MEDIUM = "Medium"
CONF_LOW = "Low"


@dataclass
class Placement:
    bsl: BSL
    kind: str               # Match / Candidate / Review
    confidence: str
    st_entries: list        # list[PoolEntry]
    codes: list             # exception / cause codes
    explanation: str
    pass_name: str
    deposit_ids: list = field(default_factory=list)
    receipt_ids: list = field(default_factory=list)


# ======================================================================
# Section 9 — FORWARD pipeline (fixed order; later never overrides earlier)
# ======================================================================

def _open_entries(pool, ledger):
    return [e for e in pool if ledger.is_available(e)]


def _amount_matches(bsl, entry):
    return entry.amount_cents == bsl.amount_cents


def forward_reconcile(bsls, pool, loaded, account, runlog):
    """Run P0-P10 over the account's open BSLs.  Returns list[Placement].
    P0 (load/validate) has already run by the time we get here."""
    ledger = Ledger()
    placements = {}          # line_key -> Placement (each BSL placed once)
    pass_counts = {}

    def place(bsl, kind, confidence, entries, codes, explanation, pass_name):
        if bsl.line_key in placements:
            return  # a later pass never overrides an earlier one
        ids = [e.id for e in entries]
        if kind == MATCH:
            ledger.consume_match(ids)
        elif kind == CANDIDATE:
            ledger.consume_candidate(ids)
        p = Placement(
            bsl=bsl, kind=kind, confidence=confidence, st_entries=list(entries),
            codes=list(codes), explanation=explanation, pass_name=pass_name,
            deposit_ids=[e.deposit_id for e in entries if e.deposit_id],
            receipt_ids=[e.receipt_id for e in entries if e.receipt_id],
        )
        placements[bsl.line_key] = p
        pass_counts[pass_name] = pass_counts.get(pass_name, 0) + 1

    def unplaced():
        return [b for b in bsls if b.line_key not in placements]

    # ---- P2 DATA_FEED_ERROR sweep -----------------------------------
    feed_errors = _p2_data_feed_errors(loaded)
    runlog.setdefault("p2_data_feed_errors", len(feed_errors))

    # ---- P3 Exact reference 1:1 (Amount + Reference) ----------------
    for bsl in unplaced():
        cands = []
        for e in _open_entries(pool, ledger):
            if not _amount_matches(bsl, e):
                continue
            if not reference_equal(bsl.recon_reference, e.reference) and \
               not reference_equal(bsl.recon_reference, e.id):
                continue
            lag = signed_lag(bsl.date, e.date)
            if not date_ok_general(lag):
                continue
            # Guardrails: journal never matches; MID-into-non-merchant conflict.
            if e.source in JOURNAL_SOURCES:
                continue
            if e.is_mid and bsl.lane != LANE_MERCHANT:
                continue
            cands.append(e)
        if len(cands) == 1:
            e = cands[0]
            lag = signed_lag(bsl.date, e.date)
            place(bsl, MATCH, CONF_HIGH, [e],
                  [], f"Exact amount + reference tie (lag {lag}d, band {date_band(lag)}); source {e.source}.",
                  "P3_exact_1to1")
        elif len(cands) > 1:
            place(bsl, CANDIDATE, CONF_MEDIUM, _sorted(cands),
                  ["MULTIPLE_EQUAL_CANDIDATES"],
                  f"{len(cands)} open entries match amount+reference; conservative Candidate.",
                  "P3_exact_1to1")

    # ---- P4 1:M ECT / ORT reference group (the workhorse) -----------
    for bsl in unplaced():
        if bsl.lane == LANE_STATE:
            continue
        group, closed_members, competing = _p4_reference_group(bsl, pool, ledger)
        if group is None:
            continue
        total = sum(e.amount_cents for e in group)
        plausible = any(date_ok_general(signed_lag(bsl.date, e.date)) for e in group)
        if total == bsl.amount_cents and plausible:
            if closed_members:
                place(bsl, CANDIDATE, CONF_MEDIUM, _sorted(group),
                      ["POSSIBLE_AUTO_REC_SPLIT"],
                      "Reference group sums exactly but includes already-closed member(s): "
                      + ", ".join(sorted(m.id for m in closed_members))
                      + " (auto-rec-stranded; hand to backward engine).",
                      "P4_ref_group")
            elif competing:
                place(bsl, CANDIDATE, CONF_MEDIUM, _sorted(group),
                      ["AMBIGUOUS_GROUP"],
                      "Reference group sums exactly but a competing equal-sum group exists.",
                      "P4_ref_group")
            else:
                place(bsl, MATCH, CONF_HIGH, _sorted(group),
                      [], f"1:M reference group of {len(group)} deduped receipt(s) sums to BSL; reference tie holds.",
                      "P4_ref_group")

    # ---- P5 STATE / Edison bundle -----------------------------------
    state_ran = _p5_state(unplaced, place, pool, ledger, loaded, runlog)

    # ---- P6 Merchant / MID ------------------------------------------
    _p6_merchant(unplaced, place, pool, ledger, loaded, runlog)

    # ---- P7 Receivables SPN group -----------------------------------
    _p7_spn(unplaced, place, pool, ledger, loaded, runlog)

    # ---- P8 Named-payer rules (Amount + Payer) ----------------------
    _p8_named_payer(unplaced, place, pool, ledger, runlog)

    # ---- P9 Payables debit ------------------------------------------
    for bsl in unplaced():
        if bsl.amount_cents >= 0:
            continue
        cands = [e for e in _open_entries(pool, ledger)
                 if e.source == "AP" and _amount_matches(bsl, e)]
        tied = [e for e in cands if reference_equal(bsl.recon_reference, e.reference)]
        if len(tied) == 1:
            e = tied[0]
            place(bsl, MATCH, CONF_HIGH, [e],
                  [], "Negative BSL matched to single reference-tied open Payables ST.",
                  "P9_payables")
        elif cands:
            place(bsl, CANDIDATE, CONF_LOW, _sorted(cands),
                  ["MISSING_REFERENCE"],
                  "Payables amount match without a clean reference tie.",
                  "P9_payables")

    # ---- P10 Residual -> Review -------------------------------------
    for bsl in unplaced():
        codes, expl = _p10_review_cause(bsl, pool, feed_errors)
        gl = recommend_gl_string(bsl, loaded)
        if gl:
            expl += f" Recommended GL: {gl}."
        place(bsl, REVIEW, CONF_LOW, [], codes, expl, "P10_review")

    runlog["forward_pass_counts"] = pass_counts
    ordered = [placements[b.line_key] for b in bsls]

    # Consumption & conservation asserts (Section 9 trailer).
    _assert_conservation(bsls, ordered)
    runlog["consumed_st_ids"] = len(ledger.consumed())
    runlog["barred_st_ids"] = len(ledger.barred())
    return ordered


def _sorted(entries):
    """Total order for determinism (amount, then date, then id)."""
    return sorted(entries, key=lambda e: (e.amount_cents,
                                          e.date.toordinal() if e.date else 0,
                                          e.id))


def _p2_data_feed_errors(loaded):
    """Open ORT parked receipt ids (r:) absent from MET => DATA_FEED_ERROR."""
    met = loaded.get("MET")
    if not met:
        return []
    met_receipt_ids = set()
    rows, m, hi = met["rows"], met["map"], met["header_index"]
    for r in rows[hi + 1:]:
        desc = N(_cell(r, m.get("description")))
        _, rec_id, _ = parse_met_description(desc)
        if rec_id:
            met_receipt_ids.add(rec_id)
    errors = []
    for role in ("ORT_AR", "ORT_MISC"):
        src = loaded.get(role)
        if not src:
            continue
        rows, m, hi = src["rows"], src["map"], src["header_index"]
        for r in rows[hi + 1:]:
            # heuristic: any r:<id> token appearing in a description-like cell
            joined = " ".join(N(c) for c in r)
            for mm in re.finditer(r"r:\s*(\d+)", joined, re.IGNORECASE):
                rid = mm.group(1)
                if rid not in met_receipt_ids:
                    errors.append(rid)
    return sorted(set(errors))


def _p4_reference_group(bsl, pool, ledger):
    """Assemble the all-status deduped reference group for a BSL.  Returns
    (group_entries or None, closed_members, competing_bool)."""
    ref = bsl.recon_reference
    if not ref:
        return None, [], False
    members = []
    for e in pool:
        if e.source in JOURNAL_SOURCES:
            continue
        if reference_equal(ref, e.reference) or reference_equal(ref, e.id):
            members.append(e)
    if not members:
        return None, [], False
    # Dedup by id keeping largest (already deduped in pool, but ORT+RECEIPTS
    # may overlap on reference; keep unique ids).
    seen, uniq = set(), []
    for e in members:
        if e.id in seen:
            continue
        seen.add(e.id)
        uniq.append(e)
    open_members = [e for e in uniq if ledger.is_available(e)]
    closed_members = [e for e in uniq if not e.available]
    # A group needs at least one open member to be actionable.
    if not open_members:
        return None, [], False
    total_open = sum(e.amount_cents for e in open_members)
    if total_open == bsl.amount_cents:
        return open_members, [], False
    # Try open + closed sum (auto-rec-stranded case).
    total_all = sum(e.amount_cents for e in uniq)
    if total_all == bsl.amount_cents and closed_members:
        return uniq, closed_members, False
    return None, [], False


def _p5_state(unplaced, place, pool, ledger, loaded, runlog):
    """STATE / Edison bundle (Section 9 P5)."""
    edison_pay = loaded.get("EDISON_PAY")
    edison_inv = loaded.get("EDISON_INV")
    ran = edison_pay is not None and edison_inv is not None
    if not ran:
        runlog["p5_state"] = "not-run: EDISON files absent"
        return False

    # index invoice -> gross
    inv_gross = {}
    rows, m, hi = edison_inv["rows"], edison_inv["map"], edison_inv["header_index"]
    for r in rows[hi + 1:]:
        inv = znorm(_cell(r, m.get("invoice_number")))
        g = cents(_cell(r, m.get("gross_amount")))
        if inv and g is not None:
            inv_gross[inv] = g
    # index reference -> (invoice set, payment amount)  (amount repeated; one per ref)
    ref_invoices = {}
    ref_amount = {}
    rows, m, hi = edison_pay["rows"], edison_pay["map"], edison_pay["header_index"]
    for r in rows[hi + 1:]:
        ref = znorm(_cell(r, m.get("reference")))
        inv = znorm(_cell(r, m.get("invoice_number")))
        amt = cents(_cell(r, m.get("amount")))
        if not ref:
            continue
        ref_invoices.setdefault(ref, set())
        if inv:
            ref_invoices[ref].add(inv)
        if amt is not None:
            ref_amount.setdefault(ref, amt)  # one per reference, never sum

    for bsl in unplaced():
        if bsl.lane != LANE_STATE:
            continue
        eref = znorm(bsl.recon_reference)
        invoices = ref_invoices.get(eref)
        if not invoices:
            place(bsl, REVIEW, CONF_LOW, [], ["STATE_LANE_ISOLATION", "NO_MATCH_FOUND"],
                  "STATE line: no Edison bundle found for reference.", "P5_state")
            continue
        bundle_sum = sum(inv_gross.get(i, 0) for i in invoices if i in inv_gross)
        # Map invoices to Receivables receipts (reference contains invoice num).
        receipts = _state_receipts(invoices, pool, ledger)
        if bundle_sum == bsl.amount_cents and receipts:
            # Match when bundle sums and Edison reference ties a receipt.
            tie = [e for e in receipts if any(inv in znorm(e.reference) or inv in znorm(e.id)
                                              for inv in invoices)]
            chosen = tie or receipts
            # State date rule handled in _state_receipts filtering.
            place(bsl, MATCH if tie else CANDIDATE,
                  CONF_HIGH if tie else CONF_MEDIUM, _sorted(chosen),
                  [] if tie else ["INCOMPLETE_REFERENCE_SUPPORT"],
                  f"STATE bundle of {len(invoices)} Edison invoice(s) sums to BSL; "
                  + ("Edison reference ties receipt." if tie else "no receipt reference tie."),
                  "P5_state")
        elif bundle_sum == bsl.amount_cents:
            place(bsl, CANDIDATE, CONF_MEDIUM, [],
                  ["INCOMPLETE_REFERENCE_SUPPORT"],
                  "STATE Edison bundle sums to BSL but no Receivables receipt found.",
                  "P5_state")
        else:
            place(bsl, REVIEW, CONF_LOW, [], ["STATE_LANE_ISOLATION", "PARTIAL_CHAIN"],
                  "STATE line: Edison invoice grosses do not sum to BSL.", "P5_state")
    runlog["p5_state"] = "ran"
    return True


def _state_receipts(invoices, pool, ledger):
    """Receipts whose reference/id contains one of the bundle invoice numbers,
    open and available.  Never a card-batch (MID) receipt."""
    out = []
    for e in pool:
        if not ledger.is_available(e) or e.is_mid or e.source == "GL":
            continue
        if any(inv and (inv in znorm(e.reference) or inv in znorm(e.id)) for inv in invoices):
            out.append(e)
    return out


def _p6_merchant(unplaced, place, pool, ledger, loaded, runlog):
    """Merchant / MID (Section 9 P6)."""
    mid_master = loaded.get("MID_MASTER")
    ran = True
    for bsl in unplaced():
        if bsl.lane != LANE_MERCHANT:
            continue
        mid = bsl.mid or (znorm(bsl.recon_reference) if is_mid(bsl.recon_reference) else "")
        if not mid:
            # merchant text but no MID: defer to reference/chain lane, not straight to Review.
            continue
        group = [e for e in _open_entries(pool, ledger)
                 if e.is_mid and znorm(e.reference) == mid]
        # card window: ST precedes BSL by 1-4 days.
        in_window = [e for e in group if date_ok_merchant(signed_lag(bsl.date, e.date))]
        stale = [e for e in group if (signed_lag(bsl.date, e.date) or 0) > 30]
        total = sum(e.amount_cents for e in in_window)
        if in_window and total == bsl.amount_cents:
            place(bsl, MATCH, CONF_HIGH, _sorted(in_window),
                  [], f"Same-MID card group ({len(in_window)}) in 1-4d window sums to settlement.",
                  "P6_merchant")
        elif stale and not in_window:
            place(bsl, REVIEW, CONF_LOW, [], ["MID_GUARDRAIL"],
                  "Stale (>30d) same-MID group; likely auto-rec artifact.", "P6_merchant")
        elif group:
            place(bsl, CANDIDATE, CONF_MEDIUM, _sorted(group),
                  ["GROUPING_CONFLICT"],
                  "Same-MID receipts present but window/sum ambiguous.", "P6_merchant")
        # else: no card receipt -> defer to P7/P10.
    runlog["p6_merchant"] = "ran" if ran else "not-run"


def _p7_spn(unplaced, place, pool, ledger, loaded, runlog):
    """Receivables SPN group (Section 9 P7)."""
    for bsl in unplaced():
        if bsl.lane == LANE_MERCHANT:
            continue
        # Group open receipts by SPN root or shared structured payment reference.
        by_spn = {}
        for e in _open_entries(pool, ledger):
            if not e.spn:
                continue
            root = _spn_root(e.spn)
            by_spn.setdefault(root, []).append(e)
        placed = False
        for root, members in sorted(by_spn.items()):
            # corroborator: shared SPN root AND (reference tie to BSL OR single member)
            corroborated = reference_tie(bsl.recon_reference, root) or \
                any(reference_tie(bsl.recon_reference, e.reference) or
                    reference_equal(bsl.recon_reference, e.reference) for e in members)
            total = sum(e.amount_cents for e in members)
            if total == bsl.amount_cents and (corroborated or len(members) == 1):
                kind = MATCH if corroborated else CANDIDATE
                place(bsl, kind, CONF_HIGH if corroborated else CONF_MEDIUM,
                      _sorted(members),
                      [] if corroborated else ["INCOMPLETE_REFERENCE_SUPPORT"],
                      f"SPN group {root} of {len(members)} receipt(s) sums to BSL"
                      + ("; corroborated." if corroborated else "; amount-only."),
                      "P7_spn")
                placed = True
                break
        if placed:
            continue
    runlog["p7_spn"] = "ran"


def _spn_root(spn):
    """Strip only a trailing -N sequence, never a space-separated reference."""
    return re.sub(r"-\d+$", "", N(spn))


# Named-payer rules (Section 9 P8).  keyword -> counterparty token set.
NAMED_PAYER_RULES = [
    ("TVA", {"TVA", "TENNESSEEVALLEYAUTHORITY"}),
    ("ASAP", {"ENERGY", "ASAP", "DEPTOFENERGY"}),
    ("NSF", {"NSF", "NATIONALSCIENCEFOUNDATION"}),
    ("UT-BATTELLE", {"BATTELLE", "UTBATTELLE"}),
    ("JEFFERSON SCIENCE", {"JEFFERSON"}),
    ("PRINCETON PLASMA", {"PRINCETON"}),
    ("FISH", {"FISH", "WILDLIFE"}),
    ("AMERICAN HEART", {"AMERICANHEART"}),
    ("STATE OF TENNESSEE", {"STATEOFTENNESSEE"}),
]


def _p8_named_payer(unplaced, place, pool, ledger, runlog):
    for bsl in unplaced():
        info_z = znorm(bsl.additional_info) + znorm(bsl.recon_reference)
        for keyword, cp_tokens in NAMED_PAYER_RULES:
            if znorm(keyword) not in info_z:
                continue
            cands = []
            for e in _open_entries(pool, ledger):
                if not _amount_matches(bsl, e):
                    continue
                lag = signed_lag(bsl.date, e.date)
                if not date_ok_general(lag):
                    continue
                cp_z = znorm(e.counterparty)
                if any(tok in cp_z for tok in cp_tokens):
                    cands.append(e)
            if len(cands) == 1:
                place(bsl, MATCH, CONF_HIGH, cands, [],
                      f"Named-payer rule '{keyword}': amount + counterparty tie.",
                      "P8_named_payer")
                break
            elif len(cands) > 1:
                place(bsl, CANDIDATE, CONF_MEDIUM, _sorted(cands),
                      ["MULTIPLE_EQUAL_CANDIDATES"],
                      f"Named-payer rule '{keyword}': multiple counterparty matches.",
                      "P8_named_payer")
                break
    runlog["p8_named_payer"] = "ran"


def _p10_review_cause(bsl, pool, feed_errors):
    """Name the dominant Review cause (Section 9 P10)."""
    exact = [e for e in pool if e.amount_cents == bsl.amount_cents]
    codes = []
    if not exact:
        codes.append("NO_MATCH_FOUND")
        expl = "No exact-amount counterpart in the pool."
    else:
        # exact amount exists but no corroboration survived.
        ref_tie = [e for e in exact if reference_tie(bsl.recon_reference, e.reference)]
        if bsl.lane == LANE_MERCHANT:
            codes.append("MID_GUARDRAIL")
            expl = "Merchant line with exact amount but no in-window card group."
        elif ref_tie:
            codes.append("DATE_CONFLICT")
            expl = "Exact amount + reference tie exists but date is out of band."
        else:
            codes.append("MISSING_REFERENCE")
            expl = "Exact amount exists but no reference/payer corroboration."
    if not bsl.recon_reference:
        codes.append("MISSING_REFERENCE")
    return list(dict.fromkeys(codes)), expl


def recommend_gl_string(bsl, loaded):
    """Best-effort GL account string for a manual ECT (Section 6 sources)."""
    # MID master first (merchant lane), then MISC receipts offset.
    mid_master = loaded.get("MID_MASTER")
    if bsl.mid and mid_master:
        gl = mid_master.get("mid_gl", {}).get(bsl.mid)
        if gl:
            return gl
    return ""


def _assert_conservation(bsls, placements):
    """Every BSL placed exactly once; no ST id consumed twice across Match+
    Candidate (Section 9 trailer)."""
    assert len(placements) == len(bsls), (
        f"conservation: {len(placements)} placements != {len(bsls)} BSLs")
    seen_lines = set()
    consumed = {}
    for p in placements:
        assert p.bsl.line_key not in seen_lines, f"BSL {p.bsl.line_key} placed twice"
        seen_lines.add(p.bsl.line_key)
        if p.kind in (MATCH, CANDIDATE):
            for e in p.st_entries:
                assert e.id not in consumed, (
                    f"ST {e.id} consumed twice ({consumed.get(e.id)} and {p.pass_name})")
                consumed[e.id] = p.pass_name


# ======================================================================
# Section 10 — BACKWARD un-reconciliation engine
# ======================================================================

@dataclass
class UnwindRec:
    recon_grp: str
    bsl_date: object
    bsl_line_info: str
    bsl_amount_cents: int
    reconciled_sts: list          # list of (id, amount_cents)
    defect_codes: list
    signed_lag: object
    rule_name: str
    recommended_action: str
    recommended_rereconciliation: str
    explanation: str


DEFECT_AMOUNT = "AMOUNT_INTEGRITY"
DEFECT_REFERENCE = "REFERENCE_INTEGRITY"
DEFECT_DATE = "DATE_PLAUSIBILITY"
DEFECT_ONE_TO_MANY = "ONE_TO_MANY_INTEGRITY"
DEFECT_SOURCE = "SOURCE_LEGALITY"
DEFECT_DUAL_FIRE = "DUAL_FIRE"
DEFECT_STATUS = "STATUS_COHERENCE"
DEFECT_CROSS_ACCOUNT = "CROSS_ACCOUNT_FLAG"

DEFECT_SEVERITY = {
    DEFECT_AMOUNT: (0, "High"),
    DEFECT_SOURCE: (0, "High"),
    DEFECT_ONE_TO_MANY: (1, "Medium"),
    DEFECT_REFERENCE: (1, "Medium"),
    DEFECT_DUAL_FIRE: (1, "Medium"),
    DEFECT_STATUS: (1, "Medium"),
    DEFECT_DATE: (2, "Low"),
    DEFECT_CROSS_ACCOUNT: (3, "Advisory"),
}


def backward_reconcile(loaded, account, pool, runlog):
    """Re-audit reconciled groups from ALL_DATA + Recon History (Section 10)."""
    all_data = loaded.get("ALL_DATA")
    if not all_data:
        runlog["backward"] = "not-run: ALL_DATA absent"
        return []
    groups = _assemble_reconciled_groups(all_data, account)
    runlog["backward_group_count"] = len(groups)
    recs = []
    defect_counts = {}
    for grp in groups:
        defects, lag = _reverify_group(grp)
        if not defects:
            continue
        for d in defects:
            defect_counts[d] = defect_counts.get(d, 0) + 1
        action, rerec = _recommend_fix(grp, defects, pool)
        recs.append(UnwindRec(
            recon_grp=grp["recon_grp"],
            bsl_date=grp["bsl_date"],
            bsl_line_info=grp["bsl_line_info"],
            bsl_amount_cents=grp["bsl_amount_cents"],
            reconciled_sts=[(e["id"], e["amount_cents"]) for e in grp["members"]],
            defect_codes=defects,
            signed_lag=lag,
            rule_name=grp["rule_name"],
            recommended_action=action,
            recommended_rereconciliation=rerec,
            explanation=_defect_explanation(grp, defects, lag),
        ))
    # Priority ordering (Section 10.4): severity, then auto-rec first.
    recs.sort(key=lambda r: (min(DEFECT_SEVERITY[d][0] for d in r.defect_codes),
                             0 if grp else 1, r.recon_grp))
    runlog["backward"] = "ran"
    runlog["backward_defect_counts"] = defect_counts
    runlog["backward_recommendations"] = len(recs)
    return recs


def _assemble_reconciled_groups(all_data, account):
    """Gather bank line(s), member receipts/STs, and the history row per
    Recon Grp (Section 10.1).  Dedup members keeping the total."""
    bsl_rows = all_data.get("bank_statement_lines", [])
    misc_rows = all_data.get("misc_receipts", [])
    ar_rows = all_data.get("ar_matched", [])
    history = all_data.get("recon_history", {})

    # index members by rec_grp
    members_by_grp = {}
    for src in (misc_rows, ar_rows):
        for row in src:
            grp = row.get("rec_grp")
            if not grp:
                continue
            members_by_grp.setdefault(grp, []).append(row)

    groups = []
    bsl_by_grp = {}
    for row in bsl_rows:
        if _norm_header(row.get("rec_status")) != "REC":
            continue
        grp = row.get("rec_grp")
        if not grp:
            continue
        bsl_by_grp.setdefault(grp, []).append(row)

    for grp, bsl_lines in sorted(bsl_by_grp.items()):
        raw_members = members_by_grp.get(grp, [])
        members = _dedup_members_keep_total(raw_members)
        # Represent the group by its first bank line (usually 1).
        b0 = bsl_lines[0]
        hist = history.get(grp, {})
        groups.append({
            "recon_grp": grp,
            "bsl_date": b0.get("date"),
            "bsl_line_info": b0.get("line_info", ""),
            "bsl_amount_cents": b0.get("amount_cents"),
            "bsl_reference": b0.get("reference", ""),
            "bsl_lane": b0.get("lane", LANE_GENERAL),
            "members": members,
            "rule_name": hist.get("rule_name", ""),
            "auto_flag": hist.get("auto_flag", ""),
            "match_type": hist.get("match_type", ""),
            "ref_match_claim": hist.get("ref_match", ""),
            "amount_match_claim": hist.get("amount_match", ""),
            "date_match_claim": hist.get("date_match", ""),
            "type_match_claim": hist.get("type_match", ""),
            "account": account,
        })
    return groups


def _dedup_members_keep_total(rows):
    """Dedup by receipt/transaction number keeping the total (largest magnitude)."""
    groups = {}
    for r in rows:
        key = r.get("id")
        groups.setdefault(key, []).append(r)
    out = []
    for key, members in groups.items():
        keep = max(members, key=lambda m: abs(m.get("amount_cents") or 0))
        out.append(keep)
    return out


def _reverify_group(grp):
    """Run all Section 10.2 checks; return (defect_codes, worst_signed_lag)."""
    defects = []
    members = grp["members"]
    bsl_amt = grp["bsl_amount_cents"]
    bsl_date = grp["bsl_date"]
    lane = grp.get("bsl_lane", LANE_GENERAL)

    # 1. AMOUNT_INTEGRITY
    total = sum((m.get("amount_cents") or 0) for m in members)
    if bsl_amt is not None and total != bsl_amt:
        defects.append(DEFECT_AMOUNT)

    # 2. REFERENCE_INTEGRITY (only when history claims Ref Match = Y)
    if _is_yes(grp.get("ref_match_claim")):
        ref = grp.get("bsl_reference", "")
        if members and not any(reference_equal(ref, m.get("reference", "")) or
                               reference_equal(ref, m.get("id", "")) for m in members):
            defects.append(DEFECT_REFERENCE)

    # 3. DATE_PLAUSIBILITY
    worst_lag = None
    for m in members:
        lag = signed_lag(bsl_date, m.get("date"))
        if lag is None:
            continue
        if worst_lag is None or abs(lag) > abs(worst_lag):
            worst_lag = lag
        if lane == LANE_STATE and lag <= 0:
            continue  # State entry-lag exempt when BSL precedes ST
        if lag >= 31 or (lane != LANE_STATE and lag <= -31):
            if DEFECT_DATE not in defects:
                defects.append(DEFECT_DATE)

    # 4. ONE_TO_MANY_INTEGRITY
    if len(members) > 1:
        keys = [(_group_key(m)) for m in members]
        # every member must share the grouping key (reference / deposit / SPN root)
        ref = grp.get("bsl_reference", "")
        belongs = []
        for m in members:
            ok = reference_equal(ref, m.get("reference", "")) or \
                reference_equal(ref, m.get("id", "")) or \
                bool(spn_of(m.get("id", "")))
            belongs.append(ok)
        if not all(belongs) and any(belongs):
            defects.append(DEFECT_ONE_TO_MANY)

    # 5. SOURCE_LEGALITY
    for m in members:
        src = _norm_header(m.get("source", ""))
        if src in JOURNAL_SOURCES:
            defects.append(DEFECT_SOURCE)
            break
        if m.get("is_mid") and lane != LANE_MERCHANT:
            defects.append(DEFECT_SOURCE)
            break

    # 6. DUAL_FIRE — one receipt id feeding two open external STs that cleared.
    id_counts = {}
    for m in members:
        rid = m.get("receipt_id") or m.get("id")
        id_counts[rid] = id_counts.get(rid, 0) + 1
    if any(c >= 2 for c in id_counts.values()):
        defects.append(DEFECT_DUAL_FIRE)

    # 7. STATUS_COHERENCE — VOID/Reversed inside a live group.
    for m in members:
        st = _norm_header(m.get("status", ""))
        if st in {"VOID", "REVERSED"}:
            defects.append(DEFECT_STATUS)
            break

    # 8. CROSS_ACCOUNT_FLAG (advisory only)
    for m in members:
        macc = m.get("account")
        if macc and macc != grp.get("account"):
            defects.append(DEFECT_CROSS_ACCOUNT)
            break

    return list(dict.fromkeys(defects)), worst_lag


def _group_key(m):
    return znorm(m.get("reference", "")) or spn_of(m.get("id", "")) or N(m.get("deposit_id"))


def _is_yes(v):
    return _norm_header(v) in {"Y", "YES", "TRUE", "1"}


def _recommend_fix(grp, defects, pool):
    """Section 10.3: run forward over the group's bank line to find the correct
    group; attach as recommended re-reconciliation, else recommend unwind."""
    # Build a synthetic BSL for the group's bank line.
    bsl = make_bsl(
        line_key=grp["recon_grp"],
        dt=grp["bsl_date"],
        amount_cents=grp["bsl_amount_cents"],
        reference=grp.get("bsl_reference", ""),
        account_servicer_reference=grp.get("bsl_reference", ""),
        additional_info=grp.get("bsl_line_info", ""),
        transaction_type="",
        transaction_code="",
    )
    ledger = Ledger()
    grp_ref = grp.get("bsl_reference", "")
    cands = []
    for e in pool:
        if not e.available or e.source in JOURNAL_SOURCES:
            continue
        if reference_equal(grp_ref, e.reference) or reference_equal(grp_ref, e.id):
            cands.append(e)
    total = sum(e.amount_cents for e in cands)
    if cands and total == grp["bsl_amount_cents"]:
        rerec = "Re-reconcile to: " + ", ".join(sorted(e.id for e in cands))
        return "RE-RECONCILE", rerec
    return "UNWIND-TO-OPEN", "Return bank line to the forward queue."


def _defect_explanation(grp, defects, lag):
    parts = []
    total = sum((m.get("amount_cents") or 0) for m in grp["members"])
    if DEFECT_AMOUNT in defects:
        parts.append(f"Members sum {_usd(total)} != BSL {_usd(grp['bsl_amount_cents'])}.")
    if DEFECT_REFERENCE in defects:
        parts.append("History claims Ref Match=Y but no member ties the BSL reference.")
    if DEFECT_ONE_TO_MANY in defects:
        parts.append("A member does not share the grouping key (mis-pulled receipt).")
    if DEFECT_SOURCE in defects:
        parts.append("Illegal source: journal or card-batch reconciled to this line.")
    if DEFECT_DATE in defects:
        parts.append(f"Implausible date lag ({lag}d).")
    if DEFECT_DUAL_FIRE in defects:
        parts.append("One receipt id feeds two cleared STs.")
    if DEFECT_STATUS in defects:
        parts.append("VOID/Reversed member inside a live group.")
    if DEFECT_CROSS_ACCOUNT in defects:
        parts.append("Advisory: a member's native account differs from the group.")
    return " ".join(parts)


# ======================================================================
# Section 13 / 10.6 — Workbook writers
# ======================================================================

FONT_NAME = "Carlito"
FONT_SIZE = 11
NAVY = "FF1F4E78"
DARK_RED = "FF7A1F1F"
WHITE = "FFFFFFFF"


def _usd(c):
    if c is None:
        return ""
    neg = c < 0
    dollars = abs(c) / 100
    s = f"{dollars:,.2f}"
    return f"({s})" if neg else s


def _fmt_date(d):
    if d is None:
        return ""
    return d.strftime("%Y-%m-%d")


def _safe_cell(v):
    """Prepend a space to any value starting with '=' (Section 3/13)."""
    s = "" if v is None else str(v)
    if s.startswith("="):
        return " " + s
    return s


def _join_multi(values):
    """Multi-value cell: comma-space, or semicolon-space if any value embeds a
    thousands-separator comma (Section 13)."""
    strs = [str(v) for v in values]
    sep = "; " if any("," in s for s in strs) else ", "
    return sep.join(strs)


def _write_workbook(path, title, tabs, header_fill_hex):
    """Generic writer applying the Section 13 formatting standard."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    wb.remove(wb.active)
    base_font = Font(name=FONT_NAME, size=FONT_SIZE)
    header_font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True, color=WHITE)
    fill = PatternFill(start_color=header_fill_hex, end_color=header_fill_hex, fill_type="solid")

    for tab_title, columns, rows in tabs:
        ws = wb.create_sheet(title=tab_title[:31])
        # Row 1 title, row 2 blank, row 3 headers, row 4+ data (Section 13).
        ws.cell(row=1, column=1, value=title).font = header_font
        for ci, col in enumerate(columns, start=1):
            c = ws.cell(row=3, column=ci, value=col)
            c.font = header_font
            c.fill = fill
            c.alignment = Alignment(vertical="center")
        for ri, row in enumerate(rows, start=4):
            for ci, val in enumerate(row, start=1):
                c = ws.cell(row=ri, column=ci, value=_safe_cell(val))
                c.font = base_font
        ws.freeze_panes = "A4"
        # Column widths (cosmetic, deterministic).
        for ci, col in enumerate(columns, start=1):
            width = max(12, min(60, len(col) + 4))
            ws.column_dimensions[get_column_letter(ci)].width = width
    wb.save(path)


RECON_COLUMNS = ["BSL Date", "BSL Line Info", "BSL Amount", "ST Date(s)",
                 "ST Number(s)", "Confidence", "ORT d:", "ORT r:", "Explanation"]


def _placement_row(p: Placement):
    st_dates = _join_multi([_fmt_date(e.date) for e in p.st_entries]) if p.st_entries else ""
    st_numbers = _join_multi([e.id for e in p.st_entries]) if p.st_entries else ""
    dep = _join_multi(sorted(set(p.deposit_ids))) if p.deposit_ids else ""
    rec = _join_multi(sorted(set(p.receipt_ids))) if p.receipt_ids else ""
    codes = (" [" + ", ".join(p.codes) + "]") if p.codes else ""
    return [
        _fmt_date(p.bsl.date),
        p.bsl.line_info,
        _usd(p.bsl.amount_cents),
        st_dates,
        st_numbers,
        p.confidence,
        dep,
        rec,
        (p.explanation + codes).strip(),
    ]


def write_reconciliation_workbook(path, account, placements):
    """Matches / Candidate Matches / Review Notes (Section 13)."""
    matches = [p for p in placements if p.kind == MATCH]
    candidates = [p for p in placements if p.kind == CANDIDATE]
    reviews = [p for p in placements if p.kind == REVIEW]
    title = f"UT Reconciliation — {account}"
    tabs = [
        ("Matches", RECON_COLUMNS, [_placement_row(p) for p in matches]),
        ("Candidate Matches", RECON_COLUMNS, [_placement_row(p) for p in candidates]),
        ("Review Notes", RECON_COLUMNS, [_placement_row(p) for p in reviews]),
    ]
    _write_workbook(path, title, tabs, NAVY)
    return {"matches": len(matches), "candidates": len(candidates), "reviews": len(reviews)}


UNWIND_COLUMNS = ["Recon Grp", "BSL Date", "BSL Line Info", "BSL Amount",
                  "Reconciled ST(s)", "Defect Code(s)", "Signed Lag",
                  "Rule Name (as reconciled)", "Recommended Action",
                  "Recommended Re-Reconciliation", "Explanation"]


def write_unwind_workbook(path, account, recs):
    """Unwind Recommendations (Section 10.6)."""
    title = f"UT Un-Reconciliation (Forensic) — {account}"
    rows = []
    for r in recs:
        sts = _join_multi([f"{i} ({_usd(a)})" for i, a in r.reconciled_sts]) if r.reconciled_sts else ""
        rows.append([
            r.recon_grp,
            _fmt_date(r.bsl_date),
            r.bsl_line_info,
            _usd(r.bsl_amount_cents),
            sts,
            ", ".join(r.defect_codes),
            "" if r.signed_lag is None else str(r.signed_lag),
            r.rule_name,
            r.recommended_action,
            r.recommended_rereconciliation,
            r.explanation,
        ])
    _write_workbook(path, title, [("Unwind Recommendations", UNWIND_COLUMNS, rows)], DARK_RED)
    return {"recommendations": len(recs)}


# ======================================================================
# ALL_DATA loader (multi-sheet, Section 4.2)
# ======================================================================

def load_all_data(routed_file: RoutedFile, account):
    """Load the ALL_DATA lifecycle workbook into a dict of normalized rows."""
    path = routed_file.path
    out = {"bank_statement_lines": [], "misc_receipts": [], "ar_matched": [],
           "recon_history": {}}

    rows, m, hi = read_named_sheet(path, "Bank Statement Lines", BANK_STATEMENT_LINES_ROLES)
    if rows is not None:
        for r in rows[hi + 1:]:
            amt = cents(_cell(r, m.get("amount")))
            if amt is None:
                continue
            ref = build_recon_reference(_cell(r, m.get("reference")),
                                        _cell(r, m.get("reference")), "")
            info = N(_cell(r, m.get("additional_info")))
            lane = classify_lane(BSL(
                line_key="", date=None, amount_cents=amt, recon_reference=ref,
                reference_raw=ref, additional_info=info, transaction_type="",
                transaction_code="", mid=_mid_from_text(info)))
            out["bank_statement_lines"].append({
                "date": parse_date(_cell(r, m.get("date"))),
                "amount_cents": amt,
                "line_key": N(_cell(r, m.get("line_key"))),
                "rec_status": N(_cell(r, m.get("rec_status"))),
                "rec_grp": N(_cell(r, m.get("rec_grp"))),
                "reference": ref,
                "line_info": (N(_cell(r, m.get("line_key"))) + " " + info).strip()[:200],
                "lane": lane,
            })

    rows, m, hi = read_named_sheet(path, "MISC Receipts", MISC_RECEIPTS_ROLES)
    if rows is not None:
        for r in rows[hi + 1:]:
            amt = cents(_cell(r, m.get("rec_amnt")))
            rid = N(_cell(r, m.get("rec_num")))
            if amt is None or not rid:
                continue
            out["misc_receipts"].append({
                "id": rid,
                "amount_cents": amt,
                "reference": N(_cell(r, m.get("rec_ref"))),
                "rec_grp": N(_cell(r, m.get("rec_grp"))),
                "deposit_id": N(_cell(r, m.get("dep_num"))),
                "status": N(_cell(r, m.get("rec_status"))),
                "date": None,
                "source": "EXT",
                "is_mid": is_mid(_cell(r, m.get("rec_ref"))),
                "account": account,
            })

    # Recon History — the ground truth (Section 4.2 / 5.3).
    rows, m, hi = read_named_sheet(path, "Recon History", RECON_HISTORY_ROLES)
    if rows is not None:
        for r in rows[hi + 1:]:
            grp = N(_cell(r, m.get("recon_grp")))
            if not grp:
                continue
            out["recon_history"][grp] = {
                "amount_cents": cents(_cell(r, m.get("amount"))),
                "auto_flag": N(_cell(r, m.get("auto_flag"))),
                "rule_name": N(_cell(r, m.get("rule_name"))),
                "match_type": N(_cell(r, m.get("match_type"))),
                "type_match": N(_cell(r, m.get("type_match"))),
                "amount_match": N(_cell(r, m.get("amount_match"))),
                "date_match": N(_cell(r, m.get("date_match"))),
                "ref_match": N(_cell(r, m.get("ref_match"))),
            }
    return out


def load_mid_master(routed_file: RoutedFile):
    """Scan every sheet for MID tokens; build MID -> account-string map."""
    from openpyxl import load_workbook
    wb = load_workbook(routed_file.path, read_only=False, data_only=True)
    mid_gl = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            mid = None
            gl = None
            for c in row:
                s = N(c)
                if is_mid(s):
                    mid = znorm(s)
                if pred_gl_string(s):
                    gl = s
            if mid and gl:
                mid_gl.setdefault(mid, gl)
    wb.close()
    return {"mid_gl": mid_gl}


# ======================================================================
# Section 9 P0 — Load & validate; build BSL list
# ======================================================================

def load_and_bind(by_role, runlog):
    """Read + bind every routed file to its required roles.  Returns `loaded`
    dict.  Raises on any required role that cannot be bound."""
    loaded = {}
    role_specs = {
        "BSL": BSL_ROLES,
        "ST": ST_ROLES,
        "RECEIPTS": RECEIPTS_ROLES,
        "MET": MET_ROLES,
        "EDISON_PAY": EDISON_PAY_ROLES,
        "EDISON_INV": EDISON_INV_ROLES,
        "APPLIED_UNAPPLIED": APPLIED_UNAPPLIED_ROLES,
        "ORT_AR": {"trx_id": _rs(False, ["Transaction Number"], pred_reference)},
        "ORT_MISC": {"trx_id": _rs(False, ["Transaction Number"], pred_reference)},
    }
    bound_report = {}
    for role, files in by_role.items():
        rf = files[0]  # newest wins; union+dedup handled in pool by id
        if role == "ALL_DATA":
            continue  # loaded separately (multi-sheet)
        if role == "MID_MASTER":
            loaded["MID_MASTER"] = load_mid_master(rf)
            continue
        specs = role_specs.get(role)
        if specs is None:
            continue  # optional file with no binding required at this layer
        try:
            rows, _title = read_rows(rf)
        except InvalidSourceData:
            raise
        mapping, hi = bind_columns(rows, specs, filename=rf.filename)
        loaded[role] = {"rows": rows, "map": mapping, "header_index": hi, "file": rf.filename}
        bound_report[role] = {"file": rf.filename, "columns": mapping, "header_index": hi}
    runlog["roles_bound"] = bound_report
    return loaded


def build_bsls(loaded, by_role, account):
    """Build the open BSL list from the BSL file (or ALL_DATA UNR lines)."""
    bsls = []
    if "BSL" in loaded:
        b = loaded["BSL"]
        rows, m, hi = b["rows"], b["map"], b["header_index"]
        for r in rows[hi + 1:]:
            amt = cents(_cell(r, m.get("amount")))
            dt = parse_date(_cell(r, m.get("date")))
            if amt is None:
                continue
            bsl = make_bsl(
                line_key=_cell(r, m.get("line_key")) or f"L{len(bsls)+1}",
                dt=dt,
                amount_cents=amt,
                reference=_cell(r, m.get("reference")),
                account_servicer_reference=_cell(r, m.get("account_servicer_reference")),
                additional_info=_cell(r, m.get("additional_info")),
                transaction_type=_cell(r, m.get("transaction_type")),
                transaction_code=_cell(r, m.get("transaction_code")),
            )
            bsls.append(bsl)
    # Deterministic order; also disambiguate any duplicate line keys.
    bsls.sort(key=lambda b: (b.date.toordinal() if b.date else 0, b.amount_cents, b.line_key))
    seen = {}
    for b in bsls:
        if b.line_key in seen:
            seen[b.line_key] += 1
            b.line_key = f"{b.line_key}#{seen[b.line_key]}"
        else:
            seen[b.line_key] = 0
    return bsls


# ======================================================================
# Section 15 — run() orchestrator + JSON run log
# ======================================================================

def run(account_input_dir, output_dir="./outputs", present=True):
    """Single entry point (Section 15.1): route -> bind -> validate -> build
    pool -> forward P0-P10 -> backward -> write workbooks -> audit -> present."""
    runlog = {"input_dir": account_input_dir, "stages": []}

    # P0 route + bind + validate
    by_role = route_folder(account_input_dir)
    runlog["files_routed"] = {role: [f.filename for f in files]
                              for role, files in by_role.items()}
    account = _resolve_account(by_role)
    runlog["account"] = account
    runlog["stages"].append("route")

    loaded = load_and_bind(by_role, runlog)
    runlog["stages"].append("bind")

    # ALL_DATA (multi-sheet) loaded separately for the backward engine.
    all_data = None
    if "ALL_DATA" in by_role:
        all_data = load_all_data(by_role["ALL_DATA"][0], account)
        loaded["ALL_DATA"] = all_data

    pool = build_pool(loaded, account, runlog)
    runlog["stages"].append("pool")

    bsls = build_bsls(loaded, by_role, account)
    runlog["bsl_count"] = len(bsls)
    runlog["bsl_by_lane"] = _count_lanes(bsls)

    # Forward P0-P10
    placements = forward_reconcile(bsls, pool, loaded, account, runlog)
    runlog["stages"].append("forward")

    # Backward 10.x
    recs = backward_reconcile(loaded, account, pool, runlog)
    runlog["stages"].append("backward")

    # Writers
    os.makedirs(output_dir, exist_ok=True)
    recon_path = os.path.join(output_dir, f"{account}_reconciliation.xlsx")
    unwind_path = os.path.join(output_dir, f"{account}_unwind.xlsx")
    recon_summary = write_reconciliation_workbook(recon_path, account, placements)
    unwind_summary = write_unwind_workbook(unwind_path, account, recs)
    runlog["recon_workbook"] = recon_path
    runlog["unwind_workbook"] = unwind_path
    runlog["recon_summary"] = recon_summary
    runlog["unwind_summary"] = unwind_summary
    runlog["stages"].append("write")

    # Conservation ledger (Section 15.3): input == Matches + Candidates + Review.
    total = recon_summary["matches"] + recon_summary["candidates"] + recon_summary["reviews"]
    assert total == len(bsls), f"conservation: {total} placed != {len(bsls)} BSLs"

    # Audit (Section 14) — imported here so the engine stays independently
    # importable; the audit module imports nothing from the engine.
    audit_result = None
    try:
        import recon_audit
        audit_result = recon_audit.audit(account_input_dir, recon_path, account)
    except ImportError:
        audit_result = {"status": "SKIPPED", "reason": "recon_audit not importable"}
    runlog["audit"] = audit_result
    runlog["stages"].append("audit")

    # Write JSON run log (Section 15.7).
    log_path = os.path.join(output_dir, f"{account}_runlog.json")
    with open(log_path, "w", encoding="utf-8") as fh:
        json.dump(runlog, fh, indent=2, default=_json_default)
    runlog["runlog_path"] = log_path

    if present and audit_result and audit_result.get("status") not in ("PASS", "SKIPPED"):
        raise ReconError(f"Audit failed; workbooks withheld: {audit_result}")
    return runlog


def _resolve_account(by_role):
    for role in ("BSL", "ALL_DATA"):
        for f in by_role.get(role, []):
            acc = infer_account(f.filename)
            if acc:
                return acc
    return "UNKNOWN"


def _count_lanes(bsls):
    out = {}
    for b in bsls:
        out[b.lane] = out.get(b.lane, 0) + 1
    return out


def _json_default(o):
    if isinstance(o, (date, datetime)):
        return o.isoformat()
    if isinstance(o, set):
        return sorted(o)
    return str(o)


# ======================================================================
# CLI
# ======================================================================

def main(argv=None):
    ap = argparse.ArgumentParser(description="UT Cash Management Reconciliation Engine")
    ap.add_argument("input_dir", help="Folder of source files for one account")
    ap.add_argument("-o", "--output", default="./outputs", help="Output folder")
    ap.add_argument("--no-present-gate", action="store_true",
                    help="Do not raise if the audit fails (write anyway)")
    args = ap.parse_args(argv)
    runlog = run(args.input_dir, args.output, present=not args.no_present_gate)
    print(json.dumps({
        "account": runlog.get("account"),
        "bsl_count": runlog.get("bsl_count"),
        "recon_summary": runlog.get("recon_summary"),
        "unwind_summary": runlog.get("unwind_summary"),
        "audit": (runlog.get("audit") or {}).get("status"),
        "runlog": runlog.get("runlog_path"),
    }, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
