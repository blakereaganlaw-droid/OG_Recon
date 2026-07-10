#!/usr/bin/env python3
"""
run_recon.py — per-run (per-upload) wrapper around the reconciliation engine.

One invocation = one upload = one account = one immutable run folder.

Given the files from a single upload (individual files, a folder, or a mix),
this script:

  1. COLLECTS every spreadsheet (.xlsx/.xlsm/.xlsb/.csv) from the given paths,
     stripping the hex upload prefix Claude web uploads add (``933782d6-Name``
     -> ``Name``; a prefix that parses as a plausible YYYYMMDD date, like
     ``20240101-``, is kept — the router's newest-wins ordering needs it).
  2. STAGES them into an isolated run folder ``<runs-root>/<run-id>/input/``.
     A filename collision between two source files fails loud — nothing is
     ever silently overwritten.
  3. PRE-FLIGHTS the upload before the engine fires: every file is classified
     with the engine's own router and printed (file -> role); unrouted
     spreadsheets are warned about, never silently dropped; a missing BSL or
     a mixed-account upload (two different account tokens) fails loud here,
     with the full routing table in the message.
  4. RECORDS provenance in ``manifest.json``: original path, staged name,
     size, sha256, and router role of every file, plus the sha256 of the
     engine and audit code that will process them and the git commit if
     available.
  5. RUNS the engine (`recon_engine.run`): route -> bind -> pool -> forward
     P0-P10 -> backward -> write workbooks -> independent audit. Outputs land
     in ``<runs-root>/<run-id>/outputs/``.
  6. REPORTS a human-readable summary (account, placements, unwind
     recommendations, audit status, output paths) and exits:
         0  engine ran and the independent audit PASSed
         2  engine ran but the audit FAILed — treat outputs/ as quarantined:
            the workbooks are on disk for forensics but are NOT approved
            for delivery; the run log names every failed check
         1  the upload itself was unusable (no spreadsheets, no BSL, mixed
            accounts, name collision, unreadable source data); the run
            folder is removed so the run-id stays free

Usage:
    python3 run_recon.py <path> [<path> ...]
    python3 run_recon.py /root/.claude/uploads/<session>/        # a web upload
    python3 run_recon.py ./drop/*.xlsx --runs-root ./runs

Options:
    --runs-root DIR      Where run folders are created (default ./runs)
    --run-id ID          Run folder name (default: UTC timestamp
                         run_YYYYMMDD_HHMMSSZ). Pass explicitly for
                         reproducible folder names.
    --no-strip-upload-prefix
                         Keep leading 8-hex-char upload prefixes on filenames.
    --no-present-gate    Do not exit 2 on audit failure (debugging only).

The engine itself stays deterministic — identical input files produce
identical workbook and run-log contents; only the run-folder name carries the
clock, and --run-id removes even that.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime, timezone

import recon_engine as E

SPREADSHEET_EXTS = (".xlsx", ".xlsm", ".xlsb", ".csv")

# Claude web uploads are prefixed "933782d6-Name.xlsx" (8 hex chars + dash).
# A prefix that parses as a plausible YYYYMMDD date is kept — the router's
# newest-wins ordering depends on it; anything else (including the ~2% of
# random hex prefixes that happen to be all digits) is stripped, because the
# engine's date-key regex would otherwise misread it as the file's date.
_UPLOAD_PREFIX_RE = re.compile(r"^([0-9a-fA-F]{8})-")


class PerRunError(E.ReconError):
    """The upload cannot be run as-is; message says exactly why."""


# ----------------------------------------------------------------------
# Collect + stage
# ----------------------------------------------------------------------

def _plausible_yyyymmdd(s: str) -> bool:
    try:
        d = datetime.strptime(s, "%Y%m%d")
    except ValueError:
        return False
    return 1990 <= d.year <= 2100


def staged_name(filename: str, strip_prefix: bool = True) -> str:
    if strip_prefix:
        m = _UPLOAD_PREFIX_RE.match(filename)
        if m and not _plausible_yyyymmdd(m.group(1)):
            rest = filename[m.end():]
            if rest:
                return rest
    return filename


def collect_sources(paths):
    """Resolve the given paths to (spreadsheets, ignored) lists of absolute
    file paths. Directories are scanned one level deep (matching the engine's
    route_folder). Missing paths fail loud."""
    spreadsheets, ignored = [], []
    for p in dict.fromkeys(os.path.abspath(p) for p in paths):
        if os.path.isdir(p):
            for name in sorted(os.listdir(p)):
                fp = os.path.join(p, name)
                if os.path.isdir(fp):
                    ignored.append(fp)  # one-level scan; preflight warns if it holds spreadsheets
                    continue
                if not os.path.isfile(fp):
                    continue
                (spreadsheets if name.lower().endswith(SPREADSHEET_EXTS) else ignored).append(fp)
        elif os.path.isfile(p):
            (spreadsheets if p.lower().endswith(SPREADSHEET_EXTS) else ignored).append(p)
        else:
            raise PerRunError(f"input path does not exist: {p}")
    # A file named explicitly AND found via its parent directory is one source.
    spreadsheets = list(dict.fromkeys(spreadsheets))
    ignored = list(dict.fromkeys(ignored))
    if not spreadsheets:
        raise PerRunError(
            "no spreadsheet files (.xlsx/.xlsm/.xlsb/.csv) found in: "
            + ", ".join(os.path.abspath(p) for p in paths))
    return spreadsheets, ignored


def _skipped_dir_warnings(ignored):
    """Folders inside the upload are not scanned (one level deep, matching the
    engine's route_folder) — but a folder holding spreadsheets must be called
    out, never silently dropped."""
    out = []
    for ig in ignored:
        if os.path.isdir(ig):
            n = sum(1 for f in os.listdir(ig)
                    if f.lower().endswith(SPREADSHEET_EXTS))
            if n:
                out.append(
                    f"{ig} is a folder; folders are scanned one level deep, so "
                    f"its {n} spreadsheet file(s) were NOT read — pass it "
                    "explicitly if it belongs to this run")
    return out


def _sha256(path, bufsize=1 << 20):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        while True:
            chunk = fh.read(bufsize)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def stage(spreadsheets, input_dir, strip_prefix=True):
    """Copy source files into input_dir under their staged names. Returns a
    list of per-file manifest entries. Duplicate staged names fail loud."""
    os.makedirs(input_dir, exist_ok=False)
    entries, taken = [], {}
    for src in spreadsheets:
        name = staged_name(os.path.basename(src), strip_prefix)
        # Case-insensitive: on macOS/Windows two names differing only in case
        # are one file on disk, and copy2 would silently overwrite the first.
        key = name.lower()
        if key in taken:
            raise PerRunError(
                f"filename collision: {src} and {taken[key]} both stage as "
                f"'{name}' (names must be unique ignoring case) — rename one "
                "and re-run")
        taken[key] = src
        dst = os.path.join(input_dir, name)
        shutil.copy2(src, dst)
        entries.append({
            "source": src,
            "staged_as": name,
            "bytes": os.path.getsize(dst),
            "sha256": _sha256(dst),
            "role": E.classify_file(name) or "UNROUTED",
            "account": E.infer_account(name),
        })
    return entries


# ----------------------------------------------------------------------
# Pre-flight
# ----------------------------------------------------------------------

def preflight(entries):
    """Fail loud on a structurally unusable upload; return warnings for
    anything odd but runnable."""
    table = "\n".join(f"  {e['staged_as']:<60s} -> {e['role']}" for e in entries)

    missing = sorted(r for r in E.HARD_REQUIRED_ROLES
                     if not any(e["role"] == r for e in entries))
    if missing:
        raise PerRunError(
            f"no file in this upload routed to hard-required role(s) "
            f"{missing}. Files routed:\n" + table)

    # One run reconciles ONE account: any file naming a different account
    # token (an ST, MET, receipts export, ...) would silently cross-pollute
    # the pool, so the conflict check spans every routed file, not just BSL.
    accounts = sorted({e["account"] for e in entries if e["account"]})
    if len(accounts) > 1:
        raise PerRunError(
            f"mixed-account upload: files name accounts {accounts}. "
            "One run reconciles one account — split the upload.\n" + table)

    warnings = []
    if not any(e["account"] for e in entries
               if e["role"] in ("BSL", "ALL_DATA")):
        warnings.append("no account token recognized in BSL/ALL_DATA "
                        "filenames; engine will report account UNKNOWN")
    for e in entries:
        if e["role"] == "UNROUTED":
            warnings.append(f"{e['staged_as']} matched no router rule; the "
                            "engine will not read it")
    return warnings


# ----------------------------------------------------------------------
# Manifest
# ----------------------------------------------------------------------

def _git_commit():
    try:
        out = subprocess.run(
            ["git", "rev-parse", "HEAD"],
            cwd=os.path.dirname(os.path.abspath(__file__)),
            capture_output=True, text=True, timeout=10)
        return out.stdout.strip() if out.returncode == 0 else None
    except (OSError, subprocess.SubprocessError):
        return None


def write_manifest(run_dir, run_id, argv, entries, ignored, warnings):
    manifest = {
        "run_id": run_id,
        "created_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "argv": argv,
        "files": entries,
        "ignored_non_spreadsheets": ignored,
        "warnings": warnings,
        "code_versions": {
            "recon_engine.py": _sha256(E.__file__),
            "recon_audit.py": _sha256(os.path.join(
                os.path.dirname(os.path.abspath(E.__file__)), "recon_audit.py")),
            "git_commit": _git_commit(),
        },
    }
    path = os.path.join(run_dir, "manifest.json")
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, indent=2)
    return path


# ----------------------------------------------------------------------
# Orchestration
# ----------------------------------------------------------------------

def perform_run(paths, runs_root="./runs", run_id=None, strip_prefix=True,
                present=True, argv=None):
    """Stage one upload and run the engine over it. Returns (exit_code,
    report_dict). Staging/pre-flight problems raise PerRunError."""
    if run_id is None:
        run_id = datetime.now(timezone.utc).strftime("run_%Y%m%d_%H%M%SZ")
    # Reject anything that could escape the runs root ('.', '..', separators).
    if not re.fullmatch(r"(?!\.+$)[A-Za-z0-9._-]+", run_id):
        raise PerRunError(f"run id '{run_id}' is not a safe folder name")

    run_dir = os.path.abspath(os.path.join(runs_root, run_id))
    if os.path.exists(run_dir):
        raise PerRunError(
            f"run folder already exists: {run_dir} — runs are immutable; "
            "pick a new --run-id")
    input_dir = os.path.join(run_dir, "input")
    output_dir = os.path.join(run_dir, "outputs")

    spreadsheets, ignored = collect_sources(paths)
    os.makedirs(run_dir)
    try:
        entries = stage(spreadsheets, input_dir, strip_prefix)
        warnings = preflight(entries) + _skipped_dir_warnings(ignored)
        manifest_path = write_manifest(run_dir, run_id, argv or [], entries,
                                       ignored, warnings)
    except BaseException:
        # Nothing irreplaceable exists yet (staged files are copies of
        # still-present sources), so free the run-id instead of leaving a
        # half-built folder behind. Once the engine runs, everything stays.
        shutil.rmtree(run_dir, ignore_errors=True)
        raise

    print(f"run {run_id}: staged {len(entries)} file(s) -> {input_dir}")
    for e in entries:
        print(f"  {e['staged_as']:<60s} -> {e['role']}")
    for w in warnings:
        print(f"  WARNING: {w}")

    audit_failed = False
    try:
        runlog = E.run(input_dir, output_dir, present=present)
    except E.ReconError as exc:
        # run() writes the runlog *before* raising on an audit failure; if
        # the runlog is there, this was the audit gate, not a crash.
        runlog = _load_runlog(output_dir)
        if runlog is None:
            raise
        audit_failed = True
        print(f"AUDIT GATE: {exc}")

    audit_status = ((runlog.get("audit") or {}).get("status"))
    report = {
        "run_id": run_id,
        "run_dir": run_dir,
        "manifest": manifest_path,
        "account": runlog.get("account"),
        "bsl_count": runlog.get("bsl_count"),
        "recon_summary": runlog.get("recon_summary"),
        "unwind_summary": runlog.get("unwind_summary"),
        "audit": audit_status,
        "recon_workbook": runlog.get("recon_workbook"),
        "unwind_workbook": runlog.get("unwind_workbook"),
        "runlog": os.path.join(output_dir, f"{runlog.get('account')}_runlog.json"),
    }
    return (2 if audit_failed else 0), report


def _load_runlog(output_dir):
    if not os.path.isdir(output_dir):
        return None
    logs = sorted(n for n in os.listdir(output_dir) if n.endswith("_runlog.json"))
    if not logs:
        return None
    with open(os.path.join(output_dir, logs[0]), encoding="utf-8") as fh:
        return json.load(fh)


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Stage one upload into an immutable run folder and "
                    "reconcile it (see module docstring).")
    ap.add_argument("paths", nargs="+",
                    help="Upload files and/or folders for ONE account")
    ap.add_argument("--runs-root", default="./runs",
                    help="Parent folder for run folders (default ./runs)")
    ap.add_argument("--run-id", default=None,
                    help="Run folder name (default: UTC timestamp)")
    ap.add_argument("--no-strip-upload-prefix", action="store_true",
                    help="Keep leading 8-hex-char upload prefixes")
    ap.add_argument("--no-present-gate", action="store_true",
                    help="Do not fail the run on audit failure (debugging)")
    args = ap.parse_args(argv)

    try:
        code, report = perform_run(
            args.paths, runs_root=args.runs_root, run_id=args.run_id,
            strip_prefix=not args.no_strip_upload_prefix,
            present=not args.no_present_gate,
            argv=list(argv) if argv is not None else sys.argv[1:])
    except E.ReconError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print(json.dumps(report, indent=2))
    if report["audit"] in ("PASS", "SKIPPED"):
        print(f"OK: audit {report['audit']}; workbooks in {report['run_dir']}/outputs")
    elif code == 0:
        print(f"WARNING: audit {report['audit']} but the present gate is "
              "off; workbooks were written anyway", file=sys.stderr)
    else:
        print("AUDIT FAILED: outputs are quarantined — the workbooks in "
              "outputs/ are NOT approved for delivery; the runlog names "
              "every failed check", file=sys.stderr)
    return code


if __name__ == "__main__":
    sys.exit(main())
