#!/usr/bin/env python3
"""
UT Cash Management Reconciliation Engine
========================================

A production-grade, deterministic reconciliation engine for University of
Tennessee bank accounts in Oracle Cash Management (DASH).

Forward matching engine: match open bank statement lines (BSL) to open
system transactions (ST), classifying each line Match / Candidate / Review.
Every available source feeds the candidate pool — ST exports, Receivables
receipts, and the MET/ORT chain (d:/r: deposit-receipt bridge).

The BACKWARD engine (re-audit reconciled groups -> recommend unwinds) was
split into the separate Unreconcile2 project so this engine stays fast:
an ALL_DATA workbook is recognized but never loaded here.

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
import functools
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


@functools.lru_cache(maxsize=1 << 18)
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

_NULL_REF_TOKENS = {"NA", "NONE", "NULL", "UNKNOWN"}


def clean_ref(v):
    """Null out placeholder reference tokens: a literal 'NA' must never tie
    another literal 'NA' (37 of 283 real FHB Master BSL references are 'NA')."""
    s = N(v)
    if znorm(s) in _NULL_REF_TOKENS or not re.search(r"[A-Za-z0-9]", s):
        return ""
    return s


def _blankish(v):
    """A placeholder cell ('NA', '.', '-') is blank for content sampling:
    real Oracle exports pad id columns with them, and counting them dilutes
    the true column's content score below junk text columns."""
    s = N(v)
    return not s or znorm(s) in _NULL_REF_TOKENS or not re.search(r"[A-Za-z0-9]", s)


def reference_equal(a, b) -> bool:
    """znorm equality OR full containment of one znorm token inside the other
    for length >= 6.  A sibling pair is a conflict, never equal."""
    za, zb = znorm(a), znorm(b)
    if not za or not zb:
        return False
    hit = (za == zb) or (len(za) >= 6 and za in zb) or (len(zb) >= 6 and zb in za)
    if not hit:
        return False
    return not sibling(a, b)


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
    # 'enriched' excluded: an Enriched_..._BSL_... workbook is the parsed
    # enrichment source (ENRICHED below), not a competing BSL export.
    RouterRule("BSL", ["bsl"], ["all_data", "enriched"], [], "first", True),
    RouterRule("ALL_DATA", ["all_data"], [], [], "multi", False),
    RouterRule("MET", ["met"], [], [], "first", False),
    # "_st" alone is too greedy: it matches All_Status / Rosetta_Stone /
    # _Statement.  Require a separator (or end) after the token.
    RouterRule("ST", [], [], ["_st_", "_st.", "account_st"], "first", False),
    # 'ar_matched' excluded: "AR_Matched_Invoice_Receipts_AR_Deposit_Receipts_
    # All_NonMisc" embeds 'receipts_all' but is the receipt-APPLICATION feed
    # (ACRA/ABA), not a receipts export — it routes to AR_MATCHED below.
    RouterRule("RECEIPTS", [], ["ar_matched"], ["receivables_receipts", "receipts_all", "oracle_receipts"], "Export to Excel", False),
    RouterRule("DEPT_INFO", [], [], ["ort_department", "department_info"], "Report", False),
    RouterRule("CHART_OF_ACCOUNTS", ["chart_of_accounts"], [], [], "Report", False),
    RouterRule("ORT_AR", ["ort", "_ar"], [], [], "Report", False),
    RouterRule("ORT_MISC", ["ort", "misc"], [], [], "Report", False),
    RouterRule("BAI2", ["bai"], [], [], "first", False),
    RouterRule("EDISON_PAY", ["edison_payments"], [], [], "first", False),
    RouterRule("EDISON_INV", ["edison_invoices"], [], [], "first", False),
    RouterRule("MID_MASTER", ["mid_master"], [], [], "all", False),
    RouterRule("ENRICHED", [], [], ["enriched", "crossref"], "first", False),
    RouterRule("APPLIED_UNAPPLIED", ["applied", "unapplied"], [], [], "first", False),
    RouterRule("CONTRACTS_INV", ["contracts_to_receivable_invoices"], [], [], "first", False),
    RouterRule("GMS_AGING", [], [], ["gms_001", "sponsored_aging"], "first", False),
    RouterRule("AR_INVOICES", ["ar_invoices"], [], [], "first", False),
    RouterRule("AR_MATCHED", [], [], ["ar_matched", "deposit_receipts"], "first", False),
    RouterRule("AR_UNAPPLIED_SUMMARY", [], [], ["ar_063", "unapplied_receipts_summary"], "first", False),
    RouterRule("GMS_SPONSOR_MAP", ["rpt_gms_0"], [], [], "first", False),
    RouterRule("CFG_MATCHING", ["matching_rules"], [], [], "first", False),
    RouterRule("CFG_PARSE", ["parse_rules"], [], [], "first", False),
    RouterRule("CFG_TOLERANCE", ["tolerance_rules"], [], [], "first", False),
    RouterRule("CFG_RULESETS", ["recon_rulesets"], [], [], "first", False),
    RouterRule("CFG_TCR", ["transaction_creation_rules"], [], [], "first", False),
    RouterRule("RELATIONSHIP_MAP", [], [], ["relationship_map", "rosetta"], "reference", False),
]

# Roles the pipeline treats as hard requirements at the top level (Section 4.1
# "Required?" column reads "yes"/conditional).  BSL is the sole unconditional
# requirement.  ALL_DATA is recognized but not loaded — it fuels the backward
# engine, which lives in the separate Unreconcile2 project.
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


def _name_segments(low):
    return [s for s in re.split(r"[^a-z0-9]+", low) if s]


def _token_hit(low, segments, tok):
    """Filename-token semantics: a token carrying a separator ('_st_',
    'all_data') or 5+ chars matches as a substring; a short bare token
    ('ar', 'st', 'met', 'bai') matches only a WHOLE name segment, with a
    numeric suffix allowed ('bai' matches 'bai2') — so 'ar' can never fire
    inside 'chart' or 'departments'."""
    if len(tok) >= 5 or any(ch in tok for ch in "_.-"):
        return tok in low
    for seg in segments:
        if seg == tok or (seg.startswith(tok) and seg[len(tok):].isdigit()):
            return True
    return False


def _tokens_present(fname_lower, tokens):
    segments = _name_segments(fname_lower)
    return all(_token_hit(fname_lower, segments, t) for t in tokens)


def classify_file(filename) -> str | None:
    """Return the internal role key for a filename, or None if unmatched.
    First matching router rule wins (Section 4.1)."""
    low = filename.lower()
    segments = _name_segments(low)
    for rule in ROUTER_TABLE:
        if rule.contains and not all(_token_hit(low, segments, t) for t in rule.contains):
            continue
        if rule.excludes and any(_token_hit(low, segments, t) for t in rule.excludes):
            continue
        if rule.any_of and not any(_token_hit(low, segments, t) for t in rule.any_of):
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


def account_of_bank_name(name):
    """Map a long bank-account name ("FHB - Master Account") to the short
    engine account (FHB_MASTER) — the scope join of ORT doc section 4.6.
    Returns None for names belonging to no configured account."""
    low = N(name).lower()
    for short, tokens in _ACCOUNT_TOKENS:
        if _tokens_present(low, tokens):
            return short
    return None


def _leading_date_key(filename):
    """Newest plausible YYYYMMDD stamp in the filename ('00000000' if none).
    An 8-digit run that does not parse as a calendar date (an account or
    merchant number like 99999999) is ignored, never mistaken for a date."""
    best = "00000000"
    for m in _YYYYMMDD_RE.finditer(filename):
        s = m.group(1)
        y, mo, d = int(s[:4]), int(s[4:6]), int(s[6:8])
        if 1990 <= y <= 2099 and 1 <= mo <= 12 and 1 <= d <= 31 and s > best:
            best = s
    return best


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


@functools.lru_cache(maxsize=1 << 17)
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
    # Reject only genuinely date-shaped values.  A bare integer is a
    # plausible reference: real Oracle exports deliver references and
    # transaction numbers as ints (1390, 852256), which parse_date would
    # otherwise swallow as Excel serials and poison the content score.
    if isinstance(v, (datetime, date)):
        return False
    if isinstance(v, str) and not s.isdigit() and parse_date(s) is not None:
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


_MET_DESC_RE = re.compile(r"(?i)\b[dr]\s*:\s*\d+")


def pred_met_description(v):
    """MET CET_DESCRIPTION carries 'd:N | r:N | payer/purpose...' — at least
    two pipe segments.  The composite ID column ('d:N | r:N', one pipe) and
    generic text columns must not satisfy this, so content scoring cannot
    prefer a stub column over the real description (real FHB Master lesson:
    the ID column scored 1.0 and hid every payer string)."""
    return N(v).count("|") >= 2


def pred_any(v):  # non-empty
    return N(v) != ""


def _cols_identical(data_rows, cols):
    """True when every listed column carries value-for-value identical content
    over the sampled data rows (normalized text).  A bank export that repeats
    a column verbatim (e.g. BAI2 with two 'Amount' columns) is not a real
    ambiguity: either binding yields identical output, so the leftmost is
    chosen deterministically instead of failing loud."""
    ref = None
    for c in cols:
        vals = tuple(N(r[c]) if c < len(r) else "" for r in data_rows)
        if ref is None:
            ref = vals
        elif vals != ref:
            return False
    return True


def _signed_twin(data_rows, cols):
    """Disambiguate a signed/unsigned amount twin: some bank exports carry the
    same amount twice — once signed, once as magnitude with the sign in a
    separate Debit/Credit column.  When every sampled row agrees in absolute
    cents across the tied columns and exactly one column carries a negative
    value, that signed column is the real amount role.  Returns its index or
    None when the pattern doesn't hold."""
    has_negative = {c: False for c in cols}
    for r in data_rows:
        mags = set()
        for c in cols:
            v = r[c] if c < len(r) else None
            if _blankish(v):
                return None
            try:
                cts = cents(v)
            except Exception:
                return None
            mags.add(abs(cts))
            if cts < 0:
                has_negative[c] = True
        if len(mags) != 1:
            return None
    signed = [c for c in cols if has_negative[c]]
    return signed[0] if len(signed) == 1 else None


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

    # 1. Locate the header row.  Never assume row 0: a deep report preamble
    # (banner + total rows) bound at row 0 silently emits phantom data rows.
    header_index = None
    best_hits = -1
    scan = min(header_scan, len(rows))
    need_hits = min(2, len({a for a in all_aliases if a})) or 1
    for i in range(scan):
        row = rows[i]
        nonempty = sum(1 for c in row if N(c) != "")
        hits = 0
        for c in row:
            hc = _norm_header(c)
            if hc and any(hc == a or (a and a in hc) for a in all_aliases):
                hits += 1
        # Prefer a row with enough columns AND enough alias hits.
        if nonempty >= min(max_arity, 2) and hits >= need_hits and hits > best_hits:
            best_hits = hits
            header_index = i
    if header_index is None:
        if any(s.required for s in role_specs.values()):
            raise InvalidSourceData(
                filename, "HEADER",
                f"no header row found in the first {scan} rows "
                f"(need >= {need_hits} recognized column headers); "
                "refusing to assume row 0")
        header_index = 0
    header = rows[header_index]
    ncols = max(len(r) for r in rows)
    data_rows = rows[header_index + 1: header_index + 1 + sample]

    # Content sampling spans the WHOLE sheet, not the first `sample` rows:
    # a sparsely-populated column (e.g. Structured Payment Reference blank on
    # every recent Journal row) must be scored on the values it actually
    # carries, or reference-shaped neighbor columns blind-tie it out of
    # binding (real FHB Master lesson, 2026-07-11: an unbound SPR column
    # silently gutted the 4x4 cross-reference screen).  Up to `sample`
    # non-blank values are collected per column in one pass.
    col_samples = {c: [] for c in range(ncols)}
    unfilled = set(col_samples)
    for r in rows[header_index + 1:]:
        if not unfilled:
            break
        for c in list(unfilled):
            if c < len(r) and not _blankish(r[c]):
                bucket = col_samples[c]
                bucket.append(r[c])
                if len(bucket) >= sample:
                    unfilled.discard(c)

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
            sampled = col_samples[col]
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
                tied = [t[2] for t in scored if (t[0], t[1]) == (best[0], best[1])]
                if _cols_identical(data_rows, tied):
                    best = (best[0], best[1], min(tied))
                else:
                    signed = _signed_twin(data_rows, tied)
                    if signed is not None:
                        best = (best[0], best[1], signed)
                    else:
                        raise AmbiguousColumn(filename, role, tied)
        if spec.required and best[0] == 0:
            raise InvalidSourceData(
                filename, role,
                f"no column satisfied content predicate (header={_norm_header(header[best[2]]) if best[2] < len(header) else ''!r})")
        if not spec.required:
            # An optional role never binds by position alone: leave it unbound
            # on zero evidence or a blind tie (same content AND header score
            # on two columns) rather than guess the leftmost column.
            if best[0] == 0 and best[1] == 0:
                continue
            if len(scored) > 1 and (best[0], best[1]) == (scored[1][0], scored[1][1]):
                tied = [t[2] for t in scored if (t[0], t[1]) == (best[0], best[1])]
                if _cols_identical(data_rows, tied):
                    best = (best[0], best[1], min(tied))
                else:
                    continue
        mapping[role] = best[2]
    return mapping, header_index


# ---- Role specs per file (Section 5.3) -------------------------------

def _rs(required, aliases, pred):
    return RoleSpec(required, aliases, pred)


BSL_ROLES = {
    "date": _rs(True, ["Transaction Date", "Booking Date", "Value Date", "Date", "Post Date"], pred_date),
    "amount": _rs(True, ["Amount", "Transaction Amount", "Signed Amount"], pred_signed_amount),
    "line_key": _rs(True, ["Line Number", "Statement Line", "Statement", "Line", "Bank Statement Line", "Sequence"], pred_number),
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
    "transaction_number": _rs(True, ["Transaction Number", "Transaction", "Trx Number", "Transaction Num", "Number"], pred_reference),
    "source": _rs(True, ["Source", "Transaction Source", "Origin"], pred_any),
    "reference": _rs(False, ["Reference", "Recon Match Reference", "Match Reference"], pred_reference),
    "structured_payment_reference": _rs(False, ["Structured Payment Reference", "Strc Pay Ref"], pred_reference),
    "counterparty": _rs(False, ["Counterparty", "Customer", "Payer", "Customer Name"], pred_customer),
    "transaction_type": _rs(False, ["Transaction Type", "Type"], pred_any),
}

RECEIPTS_ROLES = {
    "receipt_number": _rs(True, ["Receipt Number", "Receipt Num"], pred_reference),
    # 'Status' is the lifecycle (Cleared/Remitted/Reversed); the export's
    # separate 'State' column (Applied/Unapplied) must not tie it.
    "status": _rs(True, ["Status", "Receipt Status"], pred_status),
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
    "status": _rs(True, ["CET Status", "Status", "State"], pred_any),
    # Optional: the real OTBI export carries native DEPOSIT_ID / RECEIPT_ID
    # columns (authoritative; the d:/r: description parse is the fallback).
    "description": _rs(False, ["CET Description", "Description", "Desc"], pred_met_description),
    "deposit_id": _rs(False, ["Deposit ID"], pred_number),
    "receipt_id": _rs(False, ["Receipt ID"], pred_number),
    "reference_text": _rs(False, ["CET Reference Text", "Reference Text"], pred_reference),
    "bank_account_name": _rs(False, ["CBE Bank Account Name", "Bank Account Name", "Bank Account"], pred_any),
    "transaction_type": _rs(False, ["CET Transaction Type", "Transaction Type"], pred_any),
    "cleared_date": _rs(False, ["Cleared Date"], pred_date),
    "offset": _rs(False, ["Offset Concatenated Segments", "Offset", "GL"], pred_any),
}




BAI2_ROLES = {
    "post_date": _rs(True, ["Post Date", "Date"], pred_date),
    "amount": _rs(True, ["Amount"], pred_signed_amount),
    "description": _rs(False, ["Transaction Description"], pred_any),
    "bank_reference": _rs(False, ["Bank Reference"], pred_reference),
    "customer_reference": _rs(False, ["Customer Reference"], pred_reference),
    "bai_code": _rs(False, ["BAI Code"], pred_number),
}

DEPT_INFO_ROLES = {
    "department": _rs(True, ["Department Name"], pred_customer),
    "campus": _rs(False, ["Campus Name"], pred_customer),
    "dept_bank_name": _rs(False, ["Dept Bank Name"], pred_any),
    "campus_bank_name": _rs(False, ["Campus Bank Name"], pred_any),
    "mid": _rs(False, ["Credit Card Mid"], pred_mid),
}



ENRICHED_ROLES = {
    # Pre-parsed enriched BSL workbook (all rows, full year): parsed trace
    # ids / MIDs / reconciliation keys the Oracle feed does not carry.
    # Joined back onto the open BSLs by Statement line, then (date, cents).
    "date": _rs(True, ["Date"], pred_date),
    "amount": _rs(True, ["Amount (USD)", "Amount"], pred_signed_amount),
    "line_key": _rs(False, ["Statement"], pred_any),
    "additional_info": _rs(False, ["Additional Information"], pred_any),
    "parsed_trace": _rs(False, ["Parsed_Trace_ID", "Parsed Trace ID"], pred_reference),
    "parsed_mid": _rs(False, ["Parsed_MID", "Parsed MID"], pred_mid),
    "recon_key": _rs(False, ["Reconciliation_Key", "Reconciliation Key"], pred_any),
    "category": _rs(False, ["Transaction_Category", "Transaction Category"], pred_any),
}

AR_MATCHED_ROLES = {
    # ACRA/ABA receipt-application feed (AR_Matched_Invoice_Receipts...):
    # per-receipt bank-deposit dates and application context.
    "receipt_number": _rs(True, ["ACRA Receipt Number", "Receipt Number"], pred_reference),
    "amount": _rs(True, ["ACRA Amount", "Amount"], pred_signed_amount),
    "receipt_date": _rs(False, ["ACRA Receipt Date", "Receipt Date"], pred_date),
    "deposit_date": _rs(False, ["ACRA Deposit Date", "Deposit Date"], pred_date),
    "status": _rs(False, ["ACRA Status", "Status"], pred_any),
    "bank_account_name": _rs(False, ["CBA Bank Account Name", "Bank Account Name"], pred_any),
    "comments": _rs(False, ["ACRA Comments", "Comments"], pred_any),
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
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = _pick_sheet(wb, sheet_hint, os.path.basename(path))
        ws.reset_dimensions()  # never trust stated dimensions (Oracle BI)
        rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
        title = ws.title
        wb.close()
    except Exception as e:
        raise InvalidSourceData(os.path.basename(path), "WORKBOOK", f"load failed: {e}")
    if len(rows) <= 1:
        # Pathological export the streaming reader cannot see — take the
        # slow full-parse path before concluding the sheet is empty.
        try:
            wb = load_workbook(path, read_only=False, data_only=True)
        except Exception as e:
            raise InvalidSourceData(os.path.basename(path), "WORKBOOK", f"load failed: {e}")
        ws = _pick_sheet(wb, sheet_hint, os.path.basename(path))
        rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
        title = ws.title
        wb.close()
    return rows, title


def _match_sheet(wb, title_substr, filename):
    """Resolve a sheet by title substring: an exact title wins; several
    non-exact matches are ambiguous (fail loud, never take the first tab);
    zero matches returns None (caller decides the spec-mandated fallback)."""
    low = title_substr.lower().strip()
    matches = [ws for ws in wb.worksheets if low in ws.title.lower()]
    for ws in matches:
        if ws.title.lower().strip() == low:
            return ws
    if len(matches) > 1:
        raise InvalidSourceData(
            filename, title_substr,
            f"multiple sheets match {title_substr!r}: "
            f"{[ws.title for ws in matches]} — rename or pass the exact title")
    return matches[0] if matches else None


def _pick_sheet(wb, sheet_hint, filename="<workbook>"):
    if sheet_hint in ("first", "multi", "all", "reference", "stream", None):
        return wb.worksheets[0]
    ws = _match_sheet(wb, sheet_hint, filename)
    # No match: spec §4.1 mandates the first-sheet fallback for named hints.
    return ws if ws is not None else wb.worksheets[0]


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


def _read_xlsb_rows(path, sheet_hint):
    """Binary workbook reader (pyxlsb).  Dates arrive as Excel serials, which
    parse_date already accepts."""
    try:
        from pyxlsb import open_workbook
    except ImportError:
        raise InvalidSourceData(
            os.path.basename(path), "WORKBOOK",
            ".xlsb requires the pyxlsb package (pip install pyxlsb) — or "
            "re-export the file as .xlsx")
    with open_workbook(path) as wb:
        names = wb.sheets
        target = names[0]
        if sheet_hint not in ("first", "multi", "all", "reference", "stream", None):
            low = sheet_hint.lower().strip()
            matches = [n for n in names if low in n.lower()]
            exact = [n for n in matches if n.lower().strip() == low]
            if exact:
                target = exact[0]
            elif len(matches) > 1:
                raise InvalidSourceData(
                    os.path.basename(path), sheet_hint,
                    f"multiple sheets match {sheet_hint!r}: {matches}")
            elif matches:
                target = matches[0]
        with wb.get_sheet(target) as ws:
            rows = [tuple(c.v for c in row) for row in ws.rows()]
    return rows, target


def read_rows(routed_file: RoutedFile):
    """Dispatch to xlsx/xlsb/csv loader based on extension.  Wraps failures
    in InvalidSourceData (Section 15.5)."""
    ext = os.path.splitext(routed_file.path)[1].lower()
    try:
        if ext == ".csv":
            return _read_csv_rows(routed_file.path), "csv"
        if ext == ".xlsb":
            return _read_xlsb_rows(routed_file.path, routed_file.sheet_hint)
        return _read_sheet_rows(routed_file.path, routed_file.sheet_hint)
    except InvalidSourceData:
        raise
    except Exception as e:
        raise InvalidSourceData(routed_file.filename, routed_file.role, f"read failed: {e}")

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
    transaction_type: str = ""  # Credit Card / Check / EFT (type gate)
    spr: str = ""           # Structured Payment Reference (screened separately)
    base_id: str = ""       # undisambiguated id when distinct rows share one label


def _mk_entry(id, amount_cents, dt, reference, counterparty, source, status,
              available, origin, deposit_id="", receipt_id="", transaction_type="",
              spr=""):
    ref = clean_ref(reference)
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
        transaction_type=N(transaction_type),
        spr=clean_ref(spr),
    )


def _dedup_keep_largest(entries, keyfn):
    """Dedup by keyfn keeping the largest-magnitude amount (the total, not a
    split).  When the kept total lacks a counterparty, borrow it from a dropped
    split (Section 8.1).

    The keep-largest doctrine targets total-plus-invoice-split repeats, whose
    signature is largest == sum(rest) in signed cents (a lone verbatim repeat
    also satisfies it).  Same-key rows that do NOT sum that way are distinct
    receipts sharing a transaction-number label (real FHB Master case: two
    'SPN070326 ACH HRSA' receipts whose SUM is the bank line) — all are kept,
    with the display id disambiguated by amount so the ledger, conservation
    assert, and audit C3 treat them as separate consumables."""
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
        rest = members_sorted[1:]
        if rest and keep.amount_cents != sum(m.amount_cents for m in rest):
            # Distinct receipts sharing a label — keep every one.
            for m in members_sorted:
                m.base_id = m.id
                m.id = f"{m.id} [{m.amount_cents / 100:.2f}]"
                out.append(m)
            continue
        if not keep.counterparty:
            for m in rest:
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
    #    keep in the pool so P10 can explain journal-only lines; dedup
    #    keep-largest by transaction number).
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
                reference=_cell(r, m.get("reference")),
                spr=_cell(r, m.get("structured_payment_reference")),
                counterparty=_cell(r, m.get("counterparty")),
                source=src_norm,
                status="UNR",
                available=(src_norm not in JOURNAL_SOURCES),  # journals kept but ineligible
                origin="ST",
                transaction_type=_cell(r, m.get("transaction_type")),
            )
            # ST AR rows always enter the pool; when a RECEIPTS export also
            # exists it MERGES onto them below (authoritative for lifecycle
            # status) instead of replacing them — the ST export often carries
            # reference/SPR/counterparty detail the receipts export lacks
            # (real FHB Master case: a 7-receipt reference group whose ties
            # lived only in the ST export).
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
        by_id = {}
        for p in pool:
            if p.source == "AR":
                by_id.setdefault(p.base_id or p.id, []).append(p)
        merged_rc = appended = 0
        for e in deduped:
            group = by_id.get(e.base_id or e.id) or []
            # Merge by equal signed cents only — same label, different amount
            # is a distinct receipt (mirrors the MET bridge rule).
            equal = [p for p in group if p.amount_cents == e.amount_cents]
            prev = equal[0] if len(equal) == 1 else None
            if prev is not None:
                prev.status = e.status or prev.status
                prev.available = e.available
                if not prev.counterparty and e.counterparty:
                    prev.counterparty = e.counterparty
                    prev.payer_tokens = prev.payer_tokens | payer_tokens(e.counterparty)
                if (not prev.reference or prev.reference == prev.id) and \
                        e.reference and e.reference != e.id:
                    prev.reference = e.reference
                    prev.znref = znorm(e.reference)
                    prev.digits = digit_runs(e.reference)
                if prev.date is None:
                    prev.date = e.date
                merged_rc += 1
            else:
                pool.append(e)
                appended += 1
        counts["RECEIPTS"] = appended
        runlog["receipts_merged_to_st"] = merged_rc

    # 3. ORT receipts from MET for the ECT chain (index by d: and reference).
    met = loaded.get("MET")
    if met:
        rows, m, hi = met["rows"], met["map"], met["header_index"]
        raw = []
        met_total = 0
        scoped = m.get("bank_account_name") is not None and account and account != "UNKNOWN"
        for r in rows[hi + 1:]:
            met_total += 1
            # Scope join (ORT doc section 4.6): the MET export spans EVERY
            # account; keep only rows whose long bank-account name maps to
            # the account being reconciled.  Cross-account rows never enter
            # the pool — a cross-account amount hit must not become a Match.
            if scoped and account_of_bank_name(_cell(r, m.get("bank_account_name"))) != account:
                continue
            amt = cents(_cell(r, m.get("amount")))
            trx = N(_cell(r, m.get("trx_id")))
            if amt is None or not trx:
                continue
            desc = N(_cell(r, m.get("description")))
            dep_desc, rec_desc, _payer = parse_met_description(desc)
            # Native DEPOSIT_ID / RECEIPT_ID columns are authoritative; the
            # d:/r: description parse is the fallback (they agree on 160,692
            # of 160,692 overlapping real rows).
            dep_id = N(_cell(r, m.get("deposit_id"))) or dep_desc
            rec_id = N(_cell(r, m.get("receipt_id"))) or rec_desc
            ref_text = clean_ref(_cell(r, m.get("reference_text")))
            cleared = parse_date(_cell(r, m.get("cleared_date")))
            status_raw = _norm_header(_cell(r, m.get("status")))
            # Status is authoritative when present; the cleared-date=>closed
            # inference (Section 8.3) applies only when status is absent —
            # the real OTBI export stamps CET_CLEARED_DATE on UNR rows too.
            if status_raw:
                closed = status_raw in CLOSED_RECEIPT_STATUSES
            else:
                closed = cleared is not None
            e = _mk_entry(
                id=trx,
                amount_cents=amt,
                dt=parse_date(_cell(r, m.get("transaction_date"))),
                reference=ref_text or _met_reference(desc) or trx,
                counterparty=_payer,
                source="EXT",
                status=status_raw or ("REC" if closed else "UNR"),
                available=not closed,
                origin="MET",
                deposit_id=dep_id,
                receipt_id=rec_id,
                transaction_type=_cell(r, m.get("transaction_type")),
            )
            raw.append(e)
        deduped = _dedup_keep_largest(raw, lambda e: e.id)
        counts["MET"] = len(deduped)
        if scoped:
            runlog["met_scope"] = {"rows_total": met_total, "rows_in_account": len(raw)}
        # MET <-> ST bridge (CET_TRANSACTION_ID == ST Transaction Number, a
        # 1:1 bijection): the same transaction must never sit in the pool
        # twice, or reference groups double-count.  The ST export row is
        # canonical; borrow the bridge fields the MET copy carries.  MET-only
        # transactions join the pool as their own entries.
        by_id = {}
        for p in pool:
            by_id.setdefault(p.base_id or p.id, []).append(p)
        merged = 0
        dropped_mismatch = 0
        for e in deduped:
            group = by_id.get(e.base_id or e.id)
            if not group:
                pool.append(e)
                continue
            prev = None
            if len(group) == 1:
                prev = group[0]
            else:
                # Disambiguated same-label STs: the MET copy of the same
                # transaction carries the same signed cents — merge into that
                # member; never guess among unequal ones.
                equal = [p for p in group if p.amount_cents == e.amount_cents]
                if len(equal) == 1:
                    prev = equal[0]
                else:
                    dropped_mismatch += 1
                    continue
            merged += 1
            if not prev.deposit_id:
                prev.deposit_id = e.deposit_id
            if not prev.receipt_id:
                prev.receipt_id = e.receipt_id
            if not prev.counterparty and e.counterparty:
                prev.counterparty = e.counterparty
                prev.payer_tokens = prev.payer_tokens | payer_tokens(e.counterparty)
            if (not prev.reference or prev.reference == prev.id) and \
                    e.reference and e.reference != e.id:
                prev.reference = e.reference
                prev.znref = znorm(e.reference)
                prev.digits = digit_runs(e.reference)
                prev.is_mid = prev.is_mid or e.is_mid
        runlog["met_bridged_to_st"] = merged
        if dropped_mismatch:
            runlog["met_bridge_amount_mismatch_dropped"] = dropped_mismatch

    # AR receipt-application enrichment (AR_MATCHED): join by receipt number
    # (scoped to this account's bank name) to backfill missing dates with the
    # BANK deposit date and missing counterparties with the receipt comments.
    am = loaded.get("AR_MATCHED")
    if am:
        rows, m, hi = am["rows"], am["map"], am["header_index"]
        by_recno = {}
        for r in rows[hi + 1:]:
            bank = account_of_bank_name(_cell(r, m.get("bank_account_name")))
            if bank and account and account != "UNKNOWN" and bank != account:
                continue
            recno = clean_ref(_cell(r, m.get("receipt_number")))
            if not recno:
                continue
            by_recno.setdefault(recno, r)
        enriched = 0
        for e in pool:
            if e.source != "AR":
                continue
            r = by_recno.get(clean_ref(e.base_id or e.id)) or by_recno.get(clean_ref(e.reference))
            if r is None:
                continue
            touched = False
            if e.date is None:
                d = parse_date(_cell(r, m.get("deposit_date"))) or                     parse_date(_cell(r, m.get("receipt_date")))
                if d is not None:
                    e.date = d
                    touched = True
            if not e.counterparty:
                c = N(_cell(r, m.get("comments")))
                if c:
                    e.counterparty = c
                    e.payer_tokens = e.payer_tokens | payer_tokens(c)
                    touched = True
            if touched:
                enriched += 1
        runlog["ar_matched_enriched"] = enriched
        runlog["ar_matched_rows"] = len(by_recno)

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
    customer_reference: str = ""
    account_servicer_reference: str = ""
    lane: str = LANE_GENERAL
    ref_digits: set = field(default_factory=set)
    payer_tokens: set = field(default_factory=set)
    mid: str = ""
    line_info: str = ""


def build_recon_reference(reference, account_servicer_reference, transaction_code):
    """RECON_REFERENCE is the whole Account Servicer Reference for the FHB
    whole-ref transaction codes (Parse Rules (X~)); fall back to reference."""
    asr = clean_ref(account_servicer_reference)
    code = N(transaction_code)
    reference = clean_ref(reference)
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
             additional_info, transaction_type, transaction_code,
             customer_reference=""):
    recon_ref = build_recon_reference(reference, account_servicer_reference, transaction_code)
    ach_payer = extract_ach_payer(additional_info)
    mid = _mid_from_text(additional_info) or (znorm(recon_ref) if is_mid(recon_ref) else "")
    bsl = BSL(
        line_key=N(line_key),
        date=dt,
        amount_cents=amount_cents,
        recon_reference=recon_ref,
        reference_raw=clean_ref(reference),
        additional_info=N(additional_info),
        customer_reference=clean_ref(customer_reference),
        account_servicer_reference=clean_ref(account_servicer_reference),
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


def date_ok_directional(lag):
    """Owner doctrine (2026-07-11, final): the date gate applies ONLY when
    the ST was entered 8 or more days before the BSL (lag = BSL - ST >= 8).
    An ST entered after the BSL — by any amount, even months — is valid;
    unknown dates are not gated."""
    return lag is None or lag < 8


def date_ok_general(lag):
    return lag is not None and abs(lag) <= DATE_CEILING




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

def cross_reference_tie(bsl, e):
    """Owner-mandated multi-cross-reference screen (2026-07-11): EVERY BSL
    identifier field — Reference, Additional Information (full text, BAI2-
    enriched), Customer Reference, Account Servicer Reference — is compared
    against EVERY ST identifier field — Reference, Transaction Number,
    Structured Payment Reference, Counterparty.  Full cell contents; no
    field skipped, no truncation."""
    bsl_vals = (bsl.reference_raw, bsl.additional_info, bsl.customer_reference,
                bsl.account_servicer_reference, bsl.recon_reference)
    st_vals = (e.reference, e.id, e.spr, e.counterparty)
    for b in bsl_vals:
        if not b:
            continue
        for s in st_vals:
            if s and reference_equal(b, s):
                return True
    return False


def cross_digit_tie(bsl, e):
    """Weak digit-run corroboration over the same 4x4 field grid as
    cross_reference_tie: a shared >=5-digit run anywhere between the BSL's
    identifier fields and the ST's.  Supports a DATE_GATE Candidate, never a
    Match."""
    bsl_vals = (bsl.reference_raw, bsl.additional_info, bsl.customer_reference,
                bsl.account_servicer_reference, bsl.recon_reference)
    st_vals = (e.reference, e.id, e.spr, e.counterparty)
    for b in bsl_vals:
        if not b:
            continue
        for s in st_vals:
            if s and reference_tie(b, s):
                return True
    return False


_PAYER_LABEL_RE = re.compile(
    r"(?:CO(?:MPANY)?\s+NAME|CUSTOMER\s+NAME|INDIVIDUAL\s+NAME)"
    r"\s*:?\s*([A-Za-z0-9 &.,'\-]{3,40})", re.I)
_CONTRA_STOP = {"THE", "AND", "FOR", "FROM", "WITH", "INC", "LLC", "CORP",
                "CO", "OF"}


def _contra_tokens(text):
    return {w.upper() for w in _ALPHA_TOKEN_RE.findall(N(text))
            if len(w) >= 3 and w.upper() not in _CONTRA_STOP}


def _bsl_payer_name_tokens(bsl):
    """Tokens of the payer NAMES the bank line actually discloses: the
    Customer Reference field plus 'CO NAME:' / 'CUSTOMER NAME:' labeled
    segments of the (BAI2-enriched) addenda.  Deliberately NOT the filtered
    payer_tokens set — common words like STATE/PAYMENTS are stopworded there
    for tie purposes, which would fabricate contradictions between identical
    payers ('STATE-TN PAYMNTS' vs 'State of TN Payments')."""
    toks = set()
    cr = N(bsl.customer_reference)
    if cr and not _blankish(cr):
        toks |= _contra_tokens(cr)
    for m in _PAYER_LABEL_RE.finditer(N(bsl.additional_info)):
        toks |= _contra_tokens(m.group(1))
    return toks


def _token_overlap(a, b):
    """Shared token, tolerating bank-side truncation (prefix, len >= 4)."""
    if a & b:
        return True
    for x in a:
        for y in b:
            if len(x) >= 4 and len(y) >= 4 and (x.startswith(y) or y.startswith(x)):
                return True
    return False


def payer_contradiction(bsl, entries):
    """Owner directive (2026-07-11, answer-key review): an amount-only pairing
    whose payer texts actively disagree is wrong — 'City of Chattanooga has
    nothing to do with Israel.'  When the bank line names a payer (labeled
    addenda name / Customer Reference) AND the ST side carries counterparty
    text (ST export or MET CET_DESCRIPTION) and the two share no token, the
    pairing is contradicted.  Consulted only by zero-corroboration
    placements — a reference tie always outranks payer text; silence on
    either side never contradicts.

    Credits only: on a debit the bank line names the payment CHANNEL
    (Convera, card processor) while the ST counterparty names the
    BENEFICIARY — different roles, not comparable, never a contradiction."""
    if bsl.amount_cents <= 0:
        return False
    bt = _bsl_payer_name_tokens(bsl)
    if not bt:
        return False
    st = set()
    for e in entries:
        st |= _contra_tokens(e.counterparty)
    if not st:
        return False
    return not _token_overlap(bt, st)


def _is_deposit_correction(bsl):
    """Deposit-correction bank lines are manual fixes (owner, 2026-07-11):
    they rarely have an ST and need a manual ECT.  They may surface as
    flagged Candidates — never as Matches from amount-sum passes."""
    txt = _norm_header(bsl.additional_info) + _norm_header(bsl.line_info)
    return "CORRECTION" in txt or "CORRECTED" in txt


def amount_distinctive(c):
    """Owner doctrine (2026-07-11): amount alone can suffice ONLY when the
    exact signed-cents value is statistically unlikely to collide — a rare
    combination of digits.  Deterministic proxy: non-zero cents AND at least
    $1,000 in magnitude (e.g. 2,455,469.69 qualifies; 500.00, 60.00, or any
    round-dollar amount never does)."""
    return c % 100 != 0 and abs(c) >= 100000


def _is_convera(bsl):
    return "CONVERA" in _norm_header(bsl.additional_info) + _norm_header(bsl.line_info)


def _type_gate_ok(bsl, e):
    """Deterministic type gate (ORT doc section 5.2): a Credit Card ST never
    pairs with a Check or Miscellaneous bank line.  EFT is never rejected on
    the label alone.  Convera lines (owner, 2026-07-11) are international
    wires and ALWAYS Payables — they never pair with a non-Payables ST."""
    if _is_convera(bsl) and e.source != "AP":
        return False
    et = _norm_header(e.transaction_type)
    if "CREDITCARD" in et:
        bt = _norm_header(bsl.transaction_type)
        if "CHECK" in bt or "MISC" in bt:
            return False
    return True


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
        # HARD GUARDRAIL (owner doctrine, 2026-07-11): a Match or Candidate
        # without STs, or whose cited STs do not sum exactly to the BSL in
        # signed cents, must not exist.  Engine invariant — fail loud.
        if kind in (MATCH, CANDIDATE):
            if not entries:
                raise ReconError(
                    f"engine bug: {kind} for BSL {bsl.line_key} cites no ST "
                    f"({pass_name})")
            cited = sum(e.amount_cents for e in entries)
            if cited != bsl.amount_cents:
                raise ReconError(
                    f"engine bug: {kind} for BSL {bsl.line_key} cites STs "
                    f"summing {cited} != BSL {bsl.amount_cents} ({pass_name})")
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

    bsl_amount_counts = {}
    for b in bsls:
        bsl_amount_counts[b.amount_cents] = bsl_amount_counts.get(b.amount_cents, 0) + 1

    # ---- P2 DATA_FEED_ERROR sweep -----------------------------------
    feed_errors = _p2_data_feed_errors(loaded, account)
    runlog.setdefault("p2_data_feed_errors", len(feed_errors))

    # ---- P3 Exact reference 1:1 (Amount + Reference) ----------------
    stale_1to1 = {}          # line_key -> out-of-band ties, placed after P4
    for bsl in unplaced():
        cands = []
        for e in _open_entries(pool, ledger):
            if not _amount_matches(bsl, e):
                continue
            if not cross_reference_tie(bsl, e):
                continue
            # Guardrails: journal never matches; MID-into-non-merchant
            # conflict; deterministic type gate (Credit Card ST never pairs
            # with a Check/Miscellaneous line).
            if e.source in JOURNAL_SOURCES:
                continue
            if e.is_mid and bsl.lane != LANE_MERCHANT:
                continue
            if not _type_gate_ok(bsl, e):
                continue
            cands.append(e)
        ties = cands
        cands = [e for e in ties
                 if date_ok_directional(signed_lag(bsl.date, e.date))]
        if len(cands) == 1:
            e = cands[0]
            lag = signed_lag(bsl.date, e.date)
            conf = CONF_HIGH if abs(lag) <= DATE_CEILING else CONF_MEDIUM
            note = "" if abs(lag) <= DATE_CEILING else \
                f" (receipt-entry lag {-lag}d — BSL precedes ST; no ceiling)"
            place(bsl, MATCH, conf, [e],
                  [], f"Exact amount + reference tie (lag {lag}d, band {date_band(lag)})"
                  f"{note}; source {e.source}.",
                  "P3_exact_1to1")
        elif len(cands) > 1:
            ordered = _sorted(cands)
            place(bsl, CANDIDATE, CONF_MEDIUM, [ordered[0]],
                  ["MULTIPLE_EQUAL_CANDIDATES"],
                  f"{len(cands)} open entries match amount+reference; citing "
                  f"{ordered[0].id}, alternates: "
                  + ", ".join(e.id for e in ordered[1:]) + ".",
                  "P3_exact_1to1")
        elif ties:
            # Directional doctrine: the only ties here TRAIL the ST beyond
            # the band (stale-ST) — Candidate, never Match.  DEFERRED until
            # after P4: a stale 1:1 coincidence must not pre-empt a live ORT
            # deposit-chain group for the same BSL (real FHB Master case:
            # a 77d-stale reference tie shadowed deposit d:128369's exact
            # 4-receipt sum).
            stale_1to1[bsl.line_key] = ties

    # ---- P4 1:M ECT / ORT reference group (the workhorse) -----------
    for bsl in unplaced():
        group, closed_members, competing = _p4_reference_group(bsl, pool, ledger)
        if group is None:
            continue
        total = sum(e.amount_cents for e in group)
        plausible = all(date_ok_directional(signed_lag(bsl.date, e.date))
                        for e in group)
        if total == bsl.amount_cents and not plausible and not closed_members \
                and not competing:
            place(bsl, CANDIDATE, CONF_MEDIUM, _sorted(group),
                  ["DATE_CONFLICT"],
                  f"1:M reference group of {len(group)} receipt(s) sums to BSL "
                  "but the BSL trails a member ST beyond the band (stale-ST).",
                  "P4_ref_group")
            continue
        if total == bsl.amount_cents and plausible:
            if closed_members:
                place(bsl, CANDIDATE, CONF_MEDIUM, _sorted(group),
                      ["POSSIBLE_AUTO_REC_SPLIT"],
                      "Reference group sums exactly but includes already-closed member(s): "
                      + ", ".join(sorted(m.id for m in closed_members))
                      + " (auto-rec-stranded; run Unreconcile2 to re-audit the group).",
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

    # ---- P4 phase 1b: Receivables counterparty group -----------------
    # Reverse-engineered from the reference reconciliation (2026-07-11):
    # combined federal remittances land as ONE bank line covering several
    # Receivables receipts that share a counterparty and each individually
    # cross-tie the BSL (e.g. two 'SPN070326 ACH HRSA' receipts under one
    # HHS TREAS deposit).  The whole tie-set can never sum — a truncated
    # shared reference prefix fans the 4x4 screen out to thousands of
    # entries — so the sum is taken per (AR, counterparty) group.
    for bsl in unplaced():
        groups = {}
        for e in pool:
            if e.source != "AR" or not e.counterparty:
                continue
            if not ledger.is_available(e):
                continue
            if not _type_gate_ok(bsl, e):
                continue
            if not cross_reference_tie(bsl, e):
                continue
            groups.setdefault(znorm(e.counterparty), []).append(e)
        exact = [(cp, g) for cp, g in groups.items()
                 if len(g) >= 2 and sum(x.amount_cents for x in g) == bsl.amount_cents]
        if len(exact) == 1:
            cp, group = exact[0]
            place(bsl, MATCH, CONF_HIGH, _sorted(group), [],
                  f"Receivables counterparty group '{group[0].counterparty}' — "
                  f"{len(group)} receipt(s) sum exactly to BSL; every member "
                  "cross-reference-tied.",
                  "P4_cp_group")
        elif len(exact) > 1:
            ordered = sorted(exact)
            cp, group = ordered[0]
            place(bsl, CANDIDATE, CONF_MEDIUM, _sorted(group),
                  ["MULTIPLE_EQUAL_CANDIDATES"],
                  f"{len(exact)} equal-sum tied counterparty groups; citing "
                  f"'{group[0].counterparty}'.",
                  "P4_cp_group")

    # ---- P4 phase 2: ORT deposit group (d: chain) --------------------
    # Primary ORT path (relationship doc §3): index MET rows by DEPOSIT_ID;
    # the deposit's deduped open receipts sum to one BSL.  An exact sum is
    # necessary, never sufficient — corroborate by (a) reference tie,
    # or (b) payer tie (deposit-type consistency confers NOTHING —
    # owner, 2026-07-11).  All-status context: closed members completing
    # the sum mean an auto-rec split -> Candidate naming them.
    deposit_index = {}
    for e in pool:
        if e.deposit_id and e.source not in JOURNAL_SOURCES:
            deposit_index.setdefault(e.deposit_id, []).append(e)
    # Dedup each deposit by receipt_id (dual-fire) and by id, deterministically.
    for dep, members in deposit_index.items():
        seen_id, seen_rid, uniq = set(), set(), []
        for e in _sorted(members):
            if e.id in seen_id:
                continue
            if e.receipt_id:
                if e.receipt_id in seen_rid:
                    continue
                seen_rid.add(e.receipt_id)
            seen_id.add(e.id)
            uniq.append(e)
        deposit_index[dep] = uniq

    def _deposit_corroboration(bsl, members):
        for e in members:
            if reference_equal(bsl.recon_reference, e.reference) or                reference_equal(bsl.recon_reference, e.id) or                reference_tie(bsl.recon_reference, e.reference):
                return "reference tie"
        # Merchant-lane (MID) bank lines corroborate ONLY through a reference/
        # MID tie (owner, 2026-07-11): a card deposit that doesn't carry the
        # line's MID belongs to some other department — payer text and
        # deposit-type consistency prove nothing in the merchant lane.
        if bsl.lane == LANE_MERCHANT:
            return ""
        for e in members:
            if e.payer_tokens & bsl.payer_tokens:
                return "payer tie"
        # Deposit-type consistency confers NOTHING (owner, 2026-07-11
        # false-match review): "if that's the only criteria, there is no
        # candidate."  A deposit-typed line with only an exact sum falls to
        # the amount-only path like any other line.
        return ""

    for bsl in unplaced():
        exact_open, split_groups = [], []
        for dep, members in deposit_index.items():
            gated = [e for e in members if _type_gate_ok(bsl, e)]
            if not gated:
                continue
            open_m = [e for e in gated if ledger.is_available(e)]
            if not open_m:
                continue
            if sum(e.amount_cents for e in open_m) == bsl.amount_cents:
                exact_open.append((dep, open_m))
            elif sum(e.amount_cents for e in gated) == bsl.amount_cents:
                closed = [e for e in gated if not e.available]
                if closed:
                    split_groups.append((dep, gated, closed))
        correction = _is_deposit_correction(bsl)
        # Directional date gate: every member may precede the BSL by at
        # most 8 days OR be entered after it by any amount (entry lag has
        # no ceiling); a deposit the BSL trails beyond the band is stale.
        dated = [(d, g) for d, g in exact_open
                 if all(date_ok_directional(signed_lag(bsl.date, e.date))
                        for e in g)]
        if not dated and exact_open:
            corro = [(d, g) for d, g in exact_open if _deposit_corroboration(bsl, g)]
            if len(corro) == 1:
                dep, group = corro[0]
                why = _deposit_corroboration(bsl, group)
                place(bsl, CANDIDATE, CONF_MEDIUM, _sorted(group),
                      ["DATE_CONFLICT"],
                      f"ORT deposit d:{dep} sums to BSL with {why}, but the "
                      "BSL trails the deposit beyond the band (stale-ST).",
                      "P4_deposit_group")
                continue
            if corro:
                dep, group = sorted(corro)[0]
                why = _deposit_corroboration(bsl, group)
                place(bsl, CANDIDATE, CONF_MEDIUM, _sorted(group),
                      ["DATE_CONFLICT", "MULTIPLE_EQUAL_CANDIDATES"],
                      f"{len(corro)} corroborated out-of-band deposits sum to "
                      f"BSL; citing d:{dep} ({why}).",
                      "P4_deposit_group")
                continue
            # Owner call (2026-07-11): stale amount-only deposit sums surface
            # as LOW-confidence Candidates (double-flagged) rather than
            # vanishing into Review — the reviewer decides.  A payer
            # contradiction bars even the Candidate.
            uncontra = [(d, g) for d, g in exact_open
                        if not payer_contradiction(bsl, g)]
            if not uncontra:
                continue
            dep, group = sorted(uncontra)[0]
            stale_codes = ["AMOUNT_ONLY_GROUP", "DATE_CONFLICT"]
            if correction:
                stale_codes.append("MANUAL_ECT")
            place(bsl, CANDIDATE, CONF_LOW, _sorted(group),
                  stale_codes,
                  f"ORT deposit d:{dep} sums to BSL but carries no reference/"
                  "payer corroboration and no member date in band"
                  + (f"; {len(uncontra)} equal-sum deposit(s)."
                     if len(uncontra) > 1 else "."),
                  "P4_deposit_group")
            continue
        exact_open = dated
        if len(exact_open) == 1:
            dep, group = exact_open[0]
            why = _deposit_corroboration(bsl, group)
            if why and correction:
                # Owner (2026-07-11): corrections are manual fixes — surface
                # the exact-sum group as a Candidate flagged for a manual ECT,
                # never a Match.
                place(bsl, CANDIDATE, CONF_MEDIUM, _sorted(group),
                      ["MANUAL_ECT"],
                      f"Deposit correction — manual ECT required. ORT deposit "
                      f"d:{dep} sums to BSL ({why}), cited for reference only.",
                      "P4_deposit_group")
            elif why:
                place(bsl, MATCH, CONF_HIGH, _sorted(group), [],
                      f"ORT deposit d:{dep} — {len(group)} open receipt(s) sum to BSL; {why}.",
                      "P4_deposit_group")
            else:
                if payer_contradiction(bsl, group):
                    continue
                if amount_distinctive(bsl.amount_cents) and \
                        bsl_amount_counts.get(bsl.amount_cents) == 1 and \
                        not correction:
                    place(bsl, MATCH, CONF_MEDIUM, _sorted(group),
                          ["DISTINCTIVE_AMOUNT"],
                          f"Distinctive-amount ORT deposit group: d:{dep} is "
                          "the ONLY deposit summing to this non-round amount, "
                          "and no other bank line carries it — a rare digit "
                          "combination is valid match evidence (owner "
                          "doctrine).",
                          "P4_deposit_group")
                    continue
                codes = ["AMOUNT_ONLY_GROUP"]
                note = ""
                if correction:
                    codes.append("MANUAL_ECT")
                    note = " Deposit correction — manual ECT required; cited for reference only."
                place(bsl, CANDIDATE, CONF_MEDIUM, _sorted(group),
                      codes,
                      f"ORT deposit d:{dep} sums to BSL but carries no reference/payer corroboration." + note,
                      "P4_deposit_group")
        elif len(exact_open) > 1:
            corroborated = [(d, g) for d, g in exact_open if _deposit_corroboration(bsl, g)]
            if len(corroborated) == 1:
                dep, group = corroborated[0]
                why = _deposit_corroboration(bsl, group)
                if correction:
                    place(bsl, CANDIDATE, CONF_MEDIUM, _sorted(group),
                          ["MANUAL_ECT"],
                          f"Deposit correction — manual ECT required. ORT "
                          f"deposit d:{dep} sums to BSL ({why}), cited for "
                          "reference only.",
                          "P4_deposit_group")
                else:
                    place(bsl, MATCH, CONF_HIGH, _sorted(group), [],
                          f"ORT deposit d:{dep} uniquely corroborated among {len(exact_open)} equal-sum deposits.",
                          "P4_deposit_group")
            else:
                uncontra = [(d, g) for d, g in exact_open
                            if not payer_contradiction(bsl, g)]
                if not uncontra:
                    continue
                dep, group = sorted(uncontra)[0]
                multi_codes = ["MULTIPLE_EQUAL_CANDIDATES"]
                if correction:
                    multi_codes.append("MANUAL_ECT")
                place(bsl, CANDIDATE, CONF_MEDIUM, _sorted(group),
                      multi_codes,
                      f"{len(uncontra)} deposits sum to BSL: "
                      + ", ".join(f"d:{d}" for d, _ in sorted(uncontra)) + ".",
                      "P4_deposit_group")
        elif split_groups:
            dep, group, closed = sorted(split_groups)[0]
            place(bsl, CANDIDATE, CONF_MEDIUM, _sorted(group),
                  ["POSSIBLE_AUTO_REC_SPLIT"],
                  f"ORT deposit d:{dep} sums to BSL only with already-closed member(s): "
                  + ", ".join(sorted(m.id for m in closed))
                  + " (auto-rec split; run Unreconcile2).",
                  "P4_deposit_group")

    runlog["p5_state"] = "retired (owner doctrine 2026-07-11: Edison pass eliminated)"

    # ---- P6 Merchant / MID ------------------------------------------
    _p6_merchant(unplaced, place, pool, ledger, loaded, runlog, account)

    # ---- P7 Receivables SPN group -----------------------------------
    _p7_spn(unplaced, place, pool, ledger, loaded, runlog)

    # ---- P8 Named-payer rules (Amount + Payer) ----------------------
    _p8_named_payer(unplaced, place, pool, ledger, runlog)

    # ---- P8b Deferred stale 1:1 (out-of-band amount+reference ties) --
    # Held back from P3 so a stale coincidence never shadows a live ORT
    # chain / merchant / SPN group; whatever survives the group passes is
    # placed here with the same wording P3 used.
    for bsl in unplaced():
        ties = stale_1to1.get(bsl.line_key)
        if not ties:
            continue
        live = [e for e in ties if ledger.is_available(e)]
        if len(live) == 1:
            e = live[0]
            lag = signed_lag(bsl.date, e.date)
            place(bsl, CANDIDATE, CONF_MEDIUM, [e],
                  ["DATE_CONFLICT"],
                  f"Exact amount + reference tie to {e.id}, but the BSL "
                  f"trails the ST by {lag}d — stale-ST pairing, never a Match.",
                  "P8b_stale_1to1")
        elif live:
            ordered = _sorted(live)
            place(bsl, CANDIDATE, CONF_MEDIUM, [ordered[0]],
                  ["MULTIPLE_EQUAL_CANDIDATES", "DATE_CONFLICT"],
                  f"{len(live)} out-of-band amount+reference ties; citing "
                  f"{ordered[0].id}, alternates: "
                  + ", ".join(e.id for e in ordered[1:]) + ".",
                  "P8b_stale_1to1")

    # ---- P9 Payables debit ------------------------------------------
    for bsl in unplaced():
        if bsl.amount_cents >= 0:
            continue
        cands = [e for e in _open_entries(pool, ledger)
                 if e.source == "AP" and _amount_matches(bsl, e)]
        tied = [e for e in cands if cross_reference_tie(bsl, e)]
        if len(tied) == 1:
            e = tied[0]
            place(bsl, MATCH, CONF_HIGH, [e],
                  [], "Negative BSL matched to single reference-tied open Payables ST.",
                  "P9_payables")
        elif cands:
            cands = [e for e in cands if not payer_contradiction(bsl, [e])]
            if not cands:
                continue
            ordered = _sorted(cands)
            place(bsl, CANDIDATE, CONF_LOW, [ordered[0]],
                  ["MISSING_REFERENCE"],
                  "Payables amount match without a clean reference tie; citing "
                  f"{ordered[0].id}"
                  + (", alternates: " + ", ".join(e.id for e in ordered[1:])
                     if len(ordered) > 1 else "") + ".",
                  "P9_payables")

    # ---- P9b Amount-only singles -> LOW candidates -------------------
    # Reverse-engineered from the reference reconciliation (2026-07-11): an
    # unplaced BSL with open exact-cents ST(s) surfaces as a LOW Candidate
    # naming them — the hard guardrail still bars a Match without
    # corroboration, and a payer contradiction bars even the Candidate.
    for bsl in unplaced():
        elig = []
        for e in _open_entries(pool, ledger):
            if not _amount_matches(bsl, e):
                continue
            if e.source in JOURNAL_SOURCES:
                continue
            if e.is_mid and bsl.lane != LANE_MERCHANT:
                continue
            if not _type_gate_ok(bsl, e):
                continue
            if payer_contradiction(bsl, [e]):
                continue
            elig.append(e)
        if not elig:
            continue
        inband = [e for e in elig
                  if date_ok_directional(signed_lag(bsl.date, e.date))]
        if inband:
            ordered = _sorted(inband)
            if len(inband) == 1 and amount_distinctive(bsl.amount_cents) and \
                    bsl_amount_counts.get(bsl.amount_cents) == 1 and \
                    not _is_deposit_correction(bsl):
                place(bsl, MATCH, CONF_MEDIUM, [ordered[0]],
                      ["DISTINCTIVE_AMOUNT"],
                      f"Distinctive-amount 1:1: {ordered[0].id} is the ONLY "
                      "open ST at this non-round amount and no other bank "
                      "line carries it — a rare digit combination is valid "
                      "match evidence (owner doctrine).",
                      "P9b_amount_only")
                continue
            codes = ["AMOUNT_ONLY"]
            note = ""
            if _is_deposit_correction(bsl):
                codes.append("MANUAL_ECT")
                note = " Deposit correction — manual ECT required."
            place(bsl, CANDIDATE, CONF_LOW, [ordered[0]],
                  codes,
                  f"AMOUNT_ONLY: {len(inband)} open ST(s) at exact cents but "
                  f"zero cross-reference corroboration; citing {ordered[0].id}"
                  + (", alternates: " + ", ".join(e.id for e in ordered[1:])
                     if len(ordered) > 1 else "")
                  + ". Hard guardrail bars Match without corroboration." + note,
                  "P9b_amount_only")
            continue
        # Only stale exact-cents STs remain: surface as a DATE_GATE Candidate
        # when at least a digit-run tie corroborates; a zero-evidence stale
        # coincidence stays Review.
        evid = [e for e in elig if cross_digit_tie(bsl, e)]
        if evid:
            ordered = _sorted(evid)
            lag = signed_lag(bsl.date, ordered[0].date)
            place(bsl, CANDIDATE, CONF_LOW, [ordered[0]],
                  ["DATE_GATE"],
                  f"DATE_GATE: only exact-cents ST(s) precede the BSL by 8+ "
                  f"days (lag {lag}d); digit-run evidence only; citing "
                  f"{ordered[0].id}. Revised gate bars Match.",
                  "P9b_amount_only")
    runlog["p9b_amount_only"] = "ran"

    # ---- P10 Residual -> Review -------------------------------------
    for bsl in unplaced():
        codes, expl = _p10_review_cause(bsl, pool, feed_errors)
        if _is_deposit_correction(bsl):
            codes = ["MANUAL_ECT"] + [c for c in codes if c != "MANUAL_ECT"]
            expl = ("Deposit correction — rarely has an ST; manual ECT "
                    "required. " + expl)
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


def _norm_id(v):
    """Normalize a numeric id cell: ints/floats from Excel/xlsb ('64017.0')
    become bare digit strings."""
    s = N(v)
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _p2_data_feed_errors(loaded, account=None):
    """Open ORT parked receipt ids absent from MET => DATA_FEED_ERROR.
    Prefers the ORT export's bound Parked Receipt ID column (scoped to the
    account's bank name); falls back to the r:-token text scan."""
    met = loaded.get("MET")
    if not met:
        return []
    met_receipt_ids = set()
    rows, m, hi = met["rows"], met["map"], met["header_index"]
    for r in rows[hi + 1:]:
        rid = _norm_id(_cell(r, m.get("receipt_id")))
        if rid:
            met_receipt_ids.add(rid)
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
        prid = m.get("parked_receipt_id")
        bank = m.get("bank_name")
        if prid is not None:
            for r in rows[hi + 1:]:
                if bank is not None and account and account != "UNKNOWN" and \
                        account_of_bank_name(_cell(r, bank)) != account:
                    continue
                rid = _norm_id(_cell(r, prid))
                if rid and rid not in met_receipt_ids:
                    errors.append(rid)
            continue
        for r in rows[hi + 1:]:
            # fallback heuristic: any r:<id> token in a description-like cell
            joined = " ".join(N(c) for c in r)
            for mm in re.finditer(r"r:\s*(\d+)", joined, re.IGNORECASE):
                rid = mm.group(1)
                if rid not in met_receipt_ids:
                    errors.append(rid)
    return sorted(set(errors))


def _p4_reference_group(bsl, pool, ledger):
    """Assemble the all-status deduped reference group for a BSL via the
    owner-mandated 4x4 cross-reference screen.  Returns
    (group_entries or None, closed_members, competing_bool)."""
    members = []
    for e in pool:
        if e.source in JOURNAL_SOURCES:
            continue
        if not _type_gate_ok(bsl, e):
            continue
        if cross_reference_tie(bsl, e):
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
    # Dual-fire guard (ORT doc section 6.3): one ORT receipt id spawning two
    # identical STs — dedup by receipt_id before summing, deterministic keep.
    seen_rid, deduped = set(), []
    for e in _sorted(uniq):
        if e.receipt_id:
            if e.receipt_id in seen_rid:
                continue
            seen_rid.add(e.receipt_id)
        deduped.append(e)
    uniq = deduped
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








def _p6_merchant(unplaced, place, pool, ledger, loaded, runlog, account=None):
    """Merchant / MID (Section 9 P6)."""
    mid_dir = _mid_directory(loaded, account)

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
                  "Stale (>30d) same-MID group; likely auto-rec artifact."
                  + _mid_note(bsl, mid_dir, account), "P6_merchant")
        elif group:
            whole = sum(e.amount_cents for e in group)
            singles = [e for e in group if e.amount_cents == bsl.amount_cents]
            if whole == bsl.amount_cents:
                place(bsl, CANDIDATE, CONF_MEDIUM, _sorted(group),
                      ["GROUPING_CONFLICT"],
                      "Same-MID receipts sum to settlement but not inside the 1-4d window.",
                      "P6_merchant")
            elif singles:
                best = _sorted(singles)[0]
                place(bsl, CANDIDATE, CONF_MEDIUM, [best],
                      ["GROUPING_CONFLICT"],
                      f"Same-MID receipt {best.id} equals the settlement out of window; "
                      f"{len(group)} same-MID receipt(s) present.", "P6_merchant")
            else:
                place(bsl, REVIEW, CONF_LOW, [], ["GROUPING_CONFLICT"],
                      f"Same-MID receipts present ({len(group)}) but no subset sums "
                      "to the settlement.", "P6_merchant")
        # else: no card receipt -> defer to P7/P10.
    runlog["p6_merchant"] = "ran"


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
                if not _type_gate_ok(bsl, e):
                    continue
                lag = signed_lag(bsl.date, e.date)
                if not date_ok_directional(lag):
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


def _mid_directory(loaded, account):
    """MID -> {department, campus, home_account} from DEPT_INFO (+ any MIDs
    the MID master names).  Used to annotate merchant lines; never a gate."""
    out = {}
    di = loaded.get("DEPT_INFO")
    if di:
        rows, m, hi = di["rows"], di["map"], di["header_index"]
        for r in rows[hi + 1:]:
            mid = znorm(clean_ref(_cell(r, m.get("mid"))))
            if not mid or not is_mid(mid):
                continue
            home = account_of_bank_name(_cell(r, m.get("dept_bank_name"))) or \
                   account_of_bank_name(_cell(r, m.get("campus_bank_name")))
            out.setdefault(mid, {
                "department": N(_cell(r, m.get("department"))),
                "campus": N(_cell(r, m.get("campus"))),
                "home_account": home,
            })
    return out


def _mid_note(bsl, mid_dir, account):
    info = mid_dir.get(bsl.mid or znorm(bsl.recon_reference) or "")
    if not info:
        return ""
    note = f" MID belongs to {info['department']} ({info['campus']})."
    if info["home_account"] and account and account != "UNKNOWN" and \
            info["home_account"] != account:
        note += f" ADVISORY: MID's home account is {info['home_account']}, not {account}."
    return note


def _p10_review_cause(bsl, pool, feed_errors):
    """Name the dominant Review cause (Section 9 P10), distinguishing open
    counterparts from already-reconciled ones and testing the ties it names."""
    exact = [e for e in pool if e.amount_cents == bsl.amount_cents]
    open_exact = [e for e in exact if e.available]
    closed_exact = [e for e in exact if not e.available]

    def _tied(entries):
        return [e for e in entries
                if cross_reference_tie(bsl, e)
                or reference_tie(bsl.recon_reference, e.reference)
                or (e.payer_tokens & bsl.payer_tokens)]

    codes = []
    if not exact:
        codes.append("NO_MATCH_FOUND")
        expl = "No exact-amount counterpart in the pool."
    elif bsl.lane == LANE_MERCHANT and open_exact:
        codes.append("MID_GUARDRAIL")
        expl = "Merchant line with exact amount but no in-window card group."
    elif open_exact:
        tied = _tied(open_exact)
        if tied:
            codes.append("DATE_CONFLICT")
            expl = "Exact amount + reference/payer tie exists but date is out of band."
        else:
            codes.append("MISSING_REFERENCE")
            expl = ("Exact-amount open entry exists but shares no reference "
                    "or payer tie.")
    else:
        tied = _tied(closed_exact)
        codes.append("ALREADY_RECONCILED_COUNTERPART")
        if tied:
            expl = ("Only already-reconciled counterpart(s) at this amount — "
                    + ("payer/reference tie " if tied else "")
                    + "suggests an auto-rec error: "
                    + ", ".join(sorted(e.id for e in tied))
                    + ". Run Unreconcile2.")
        else:
            expl = ("Only already-reconciled (closed) counterpart(s) at this "
                    "amount; no open entry. Run Unreconcile2 if misdirected.")
    if not bsl.recon_reference:
        codes.append("MISSING_REFERENCE")
    return list(dict.fromkeys(codes)), expl


def recommend_gl_string(bsl, loaded):
    """Best-effort GL account string for a manual ECT (Section 6 sources)."""
    # MID master only (merchant lane); MISC Receipts is an ALL_DATA sheet,
    # never loaded here — returns "" when there is no MID hit.
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
# Section 13 — Workbook writer (the §10.6 unwind writer is Unreconcile2's)
# ======================================================================

FONT_NAME = "Carlito"
FONT_SIZE = 11
NAVY = "FF1F4E78"
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


# HARD GUARDRAIL (owner, 2026-07-11): the output carries ALL BSL identifier
# fields and ALL ST detail fields — they are essential for reconciliation.
# The ST Number(s) cell lists EVERY cited ST, never truncated; the audit's
# C9 pins this exact 19-column layout and C10 rejects extra columns.
RECON_COLUMNS = [
    "BSL Date", "BSL Line Info", "BSL Amount",
    "BSL Reference", "BSL Additional Information", "BSL Customer Reference",
    "BSL Account Servicer Reference", "BSL Transaction Type",
    "ST Date(s)", "ST Number(s)", "ST Amount(s)", "ST Reference(s)",
    "ST Structured Payment Reference(s)", "ST Counterparty(ies)",
    "ST Source(s)",
    "Confidence", "ORT d:", "ORT r:", "Explanation",
]


def _placement_row(p: Placement):
    ents = p.st_entries
    st_dates = _join_multi([_fmt_date(e.date) for e in ents]) if ents else ""
    st_numbers = _join_multi([e.id for e in ents]) if ents else ""
    st_amounts = _join_multi([_usd(e.amount_cents) for e in ents]) if ents else ""
    st_refs = _join_multi([e.reference or "" for e in ents]) if ents else ""
    st_sprs = _join_multi([e.spr or "" for e in ents]) if ents else ""
    st_cps = _join_multi([e.counterparty or "" for e in ents]) if ents else ""
    st_srcs = _join_multi([e.source or "" for e in ents]) if ents else ""
    dep = _join_multi(sorted(set(p.deposit_ids))) if p.deposit_ids else ""
    rec = _join_multi(sorted(set(p.receipt_ids))) if p.receipt_ids else ""
    codes = (" [" + ", ".join(p.codes) + "]") if p.codes else ""
    return [
        _fmt_date(p.bsl.date),
        p.bsl.line_info,
        _usd(p.bsl.amount_cents),
        p.bsl.reference_raw,
        p.bsl.additional_info,
        p.bsl.customer_reference,
        p.bsl.account_servicer_reference,
        p.bsl.transaction_type,
        st_dates,
        st_numbers,
        st_amounts,
        st_refs,
        st_sprs,
        st_cps,
        st_srcs,
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


# ======================================================================
# MID master loader (Section 4.2)
# ======================================================================
def load_mid_master(routed_file: RoutedFile):
    """Scan every sheet for MID tokens; build MID -> account-string map."""
    from openpyxl import load_workbook
    wb = load_workbook(routed_file.path, read_only=False, data_only=True)
    mid_gl = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            mids, gls = [], []
            for c in row:
                s = N(c)
                if is_mid(s):
                    mids.append(znorm(s))
                if pred_gl_string(s):
                    gls.append(s)
            if not mids or not gls:
                continue
            # Never let column position pick silently: a row naming two
            # distinct GL strings, or remapping a MID already mapped to a
            # different GL, is genuinely ambiguous — fail loud.
            if len(set(gls)) > 1:
                raise InvalidSourceData(
                    routed_file.filename, "MID_MASTER",
                    f"sheet {ws.title!r}: row maps MID(s) {sorted(set(mids))} "
                    f"to multiple distinct GL strings {sorted(set(gls))}")
            for mid in dict.fromkeys(mids):
                prev = mid_gl.setdefault(mid, gls[0])
                if prev != gls[0]:
                    raise InvalidSourceData(
                        routed_file.filename, "MID_MASTER",
                        f"MID {mid} maps to conflicting GL strings "
                        f"{prev!r} and {gls[0]!r}")
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
        "BAI2": BAI2_ROLES,
        "DEPT_INFO": DEPT_INFO_ROLES,
        "APPLIED_UNAPPLIED": APPLIED_UNAPPLIED_ROLES,
        "AR_MATCHED": AR_MATCHED_ROLES,
        "ENRICHED": ENRICHED_ROLES,
        "ORT_AR": {"trx_id": _rs(False, ["Transaction Number"], pred_reference),
                   "parked_receipt_id": _rs(False, ["Parked Receipt ID"], pred_number),
                   "bank_name": _rs(False, ["BANK_NAME", "Bank Account Name"], pred_any)},
        "ORT_MISC": {"trx_id": _rs(False, ["Transaction Number"], pred_reference),
                     "parked_receipt_id": _rs(False, ["Parked Receipt ID"], pred_number),
                     "bank_name": _rs(False, ["BANK_NAME", "Bank Account Name"], pred_any)},
    }
    bound_report = {}
    for role, files in by_role.items():
        if role == "ALL_DATA":
            continue  # recognized but never loaded (Unreconcile2 owns it)
        specs = role_specs.get(role)
        if specs is None and role != "MID_MASTER":
            continue  # recognized; no binding required at this layer
        rf = _pick_newest(files, role)  # newest YYYYMMDD stamp wins
        if role == "MID_MASTER":
            loaded["MID_MASTER"] = load_mid_master(rf)
            continue
        try:
            rows, _title = read_rows(rf)
        except InvalidSourceData:
            raise
        mapping, hi = bind_columns(rows, specs, filename=rf.filename)
        loaded[role] = {"rows": rows, "map": mapping, "header_index": hi, "file": rf.filename}
        bound_report[role] = {"file": rf.filename, "columns": mapping, "header_index": hi}
    runlog["roles_bound"] = bound_report
    return loaded


def _pick_newest(files, role):
    """Newest-wins file selection for a role (spec §4.1).  Two files whose
    date keys tie would mean one silently ignored — fail loud instead and
    ask for a disambiguating YYYYMMDD stamp."""
    if len(files) > 1 and _leading_date_key(files[0].filename) == _leading_date_key(files[1].filename):
        raise InvalidSourceData(
            files[0].filename, role,
            f"multiple files tie for role {role}: "
            f"{[f.filename for f in files]} — add a YYYYMMDD date stamp so "
            "the newest can be chosen (the rest are never read silently)")
    return files[0]


def _bai2_index(loaded):
    """Index BAI2 rows by (date, signed cents).  DETAIL1..DETAIL10 columns are
    located by exact header name (deterministic; they carry the full addenda
    the Oracle BSL feed truncates at ~1000 chars)."""
    bai = loaded.get("BAI2")
    if not bai:
        return None
    rows, m, hi = bai["rows"], bai["map"], bai["header_index"]
    header = rows[hi]
    detail_cols = [i for i, h in enumerate(header)
                   if re.fullmatch(r"DETAIL\d+", _norm_header(h) or "")]
    idx = {}
    for r in rows[hi + 1:]:
        dt = parse_date(_cell(r, m.get("post_date")))
        amt = cents(_cell(r, m.get("amount")))
        if dt is None or amt is None:
            continue
        details = "".join(N(_cell(r, c)) for c in detail_cols)
        idx.setdefault((dt, amt), []).append({
            "details": details,
            "description": N(_cell(r, m.get("description"))),
            "bank_reference": clean_ref(_cell(r, m.get("bank_reference"))),
            "customer_reference": clean_ref(_cell(r, m.get("customer_reference"))),
        })
    return idx


def _bai2_enrich(info, dt, amt, bai2_idx, counters):
    """Resolve the BSL line's BAI2 row by (date, cents), tiebreaking on the
    space-normalized addenda prefix; never guess on a residual tie.  Returns
    the enriched additional-information text."""
    cands = bai2_idx.get((dt, amt), [])
    if not cands:
        counters["no_hit"] += 1
        return info
    chosen = None
    if len(cands) == 1:
        chosen = cands[0]
    else:
        key = N(info)[:40].replace(" ", "")
        pref = [c for c in cands
                if key and c["details"].replace(" ", "").startswith(key)]
        if len(pref) == 1:
            chosen = pref[0]
    if chosen is None:
        counters["ambiguous"] += 1
        return info
    counters["joined"] += 1
    extra = " ".join(x for x in (chosen["description"], chosen["details"],
                                 chosen["bank_reference"], chosen["customer_reference"]) if x)
    return (N(info) + " | BAI2: " + extra).strip(" |")


def _enriched_index(loaded):
    """Index the pre-parsed enriched-BSL workbook by Statement line key and by
    (date, signed cents)."""
    en = loaded.get("ENRICHED")
    if not en:
        return None
    rows, m, hi = en["rows"], en["map"], en["header_index"]
    by_line, by_da = {}, {}
    for r in rows[hi + 1:]:
        dt = parse_date(_cell(r, m.get("date")))
        amt = cents(_cell(r, m.get("amount")))
        if amt is None:
            continue
        rec = {
            "info": N(_cell(r, m.get("additional_info"))),
            "trace": clean_ref(_cell(r, m.get("parsed_trace"))),
            "mid": znorm(clean_ref(_cell(r, m.get("parsed_mid")))),
            "key": N(_cell(r, m.get("recon_key"))),
            "category": N(_cell(r, m.get("category"))),
        }
        lk = znorm(_cell(r, m.get("line_key")))
        if lk:
            by_line.setdefault(lk, []).append(rec)
        if dt is not None:
            by_da.setdefault((dt, amt), []).append(rec)
    return {"by_line": by_line, "by_da": by_da}


def _enriched_enrich(info, line_key, dt, amt, en_idx, counters):
    """Join one BSL to its enriched-workbook row: exact Statement line key
    first, unique (date, cents) as fallback; never guess a residual tie.
    Appends the parsed trace/MID/recon-key text to the addenda."""
    cands = en_idx["by_line"].get(znorm(line_key), [])
    if len(cands) != 1:
        da = en_idx["by_da"].get((dt, amt), [])
        if len(da) == 1:
            cands = da
        elif cands or da:
            counters["ambiguous"] += 1
            return info
        else:
            counters["no_hit"] += 1
            return info
    chosen = cands[0]
    counters["joined"] += 1
    extra = " ".join(x for x in (
        chosen["category"],
        ("trace:" + chosen["trace"]) if chosen["trace"] else "",
        ("mid:" + chosen["mid"]) if chosen["mid"] else "",
        chosen["key"],
    ) if x)
    base = N(info)
    # Prefer the LONGER additional-information text (the enriched workbook
    # carries the untruncated addenda when no BAI2 export is present).
    if len(chosen["info"]) > len(base):
        base = chosen["info"]
    if not extra:
        return base
    return (base + " | ENR: " + extra).strip(" |")


def build_bsls(loaded, by_role, account, runlog=None):
    """Build the open BSL list from the BSL file, enriched with the matching
    BAI2 row's full addenda + raw bank references when a BAI2 export is
    present (the chain: BSL -> BAI2 -> ORT/MET text -> amounts -> ST)."""
    bsls = []
    bai2_idx = _bai2_index(loaded)
    en_idx = _enriched_index(loaded)
    counters = {"joined": 0, "ambiguous": 0, "no_hit": 0}
    en_counters = {"joined": 0, "ambiguous": 0, "no_hit": 0}
    if "BSL" in loaded:
        b = loaded["BSL"]
        rows, m, hi = b["rows"], b["map"], b["header_index"]
        for r in rows[hi + 1:]:
            amt = cents(_cell(r, m.get("amount")))
            dt = parse_date(_cell(r, m.get("date")))
            if amt is None:
                continue
            info = N(_cell(r, m.get("additional_info")))
            if bai2_idx is not None:
                info = _bai2_enrich(info, dt, amt, bai2_idx, counters)
            if en_idx is not None:
                info = _enriched_enrich(
                    info, _cell(r, m.get("line_key")) or "", dt, amt,
                    en_idx, en_counters)
            bsl = make_bsl(
                line_key=_cell(r, m.get("line_key")) or f"L{len(bsls)+1}",
                dt=dt,
                amount_cents=amt,
                reference=_cell(r, m.get("reference")),
                account_servicer_reference=_cell(r, m.get("account_servicer_reference")),
                additional_info=info,
                transaction_type=_cell(r, m.get("transaction_type")),
                transaction_code=_cell(r, m.get("transaction_code")),
                customer_reference=_cell(r, m.get("customer_reference")),
            )
            bsls.append(bsl)
    if runlog is not None and bai2_idx is not None:
        runlog["bai2_enrichment"] = counters
    if runlog is not None and en_idx is not None:
        runlog["enriched_bsl_enrichment"] = en_counters
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
    pool -> forward P0-P10 -> write workbook -> audit -> present."""
    runlog = {"input_dir": account_input_dir, "stages": []}

    # P0 route + bind + validate
    by_role = route_folder(account_input_dir)
    runlog["files_routed"] = {role: [f.filename for f in files]
                              for role, files in by_role.items()}
    account = _resolve_account(by_role)
    runlog["account"] = account
    runlog["stages"].append("route")

    # ALL_DATA feeds nothing forward — it is the backward engine's fuel,
    # which now lives in the Unreconcile2 project.  Recognized, noted, and
    # deliberately not loaded (this is the forward engine's speed win).
    if "ALL_DATA" in by_role:
        runlog["all_data"] = ("present: not loaded (forward-only engine; "
                              "run Unreconcile2 for the backward re-audit)")

    loaded = load_and_bind(by_role, runlog)
    runlog["stages"].append("bind")

    pool = build_pool(loaded, account, runlog)
    runlog["stages"].append("pool")

    bsls = build_bsls(loaded, by_role, account, runlog)
    runlog["bsl_count"] = len(bsls)
    runlog["bsl_by_lane"] = _count_lanes(bsls)

    # Forward P0-P10
    placements = forward_reconcile(bsls, pool, loaded, account, runlog)
    runlog["stages"].append("forward")

    # Writer
    os.makedirs(output_dir, exist_ok=True)
    recon_path = os.path.join(output_dir, f"{account}_reconciliation.xlsx")
    recon_summary = write_reconciliation_workbook(recon_path, account, placements)
    runlog["recon_workbook"] = recon_path
    runlog["recon_summary"] = recon_summary
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
        "audit": (runlog.get("audit") or {}).get("status"),
        "runlog": runlog.get("runlog_path"),
    }, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
