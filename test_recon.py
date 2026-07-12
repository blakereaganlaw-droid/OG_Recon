#!/usr/bin/env python3
"""
Unit tests for the reconciliation engine.

Covers Section 7 primitives, Section 4 router, Section 5 binder, Section 8
pool dedup, and an end-to-end synthetic run through the forward pipeline
+ independent audit (Section 16 build order steps 1-8, validated on
a synthetic account since real UT source files are not present here).

Run:  python3 -m unittest test_recon -v
"""

import os
import shutil
import tempfile
import unittest
from datetime import date

import recon_engine as E
import recon_audit as A
import run_recon as R


class TestPrimitives(unittest.TestCase):
    def test_N(self):
        self.assertEqual(E.N(None), "")
        self.assertEqual(E.N("  x "), "x")
        self.assertEqual(E.N(5), "5")

    def test_cents(self):
        self.assertEqual(E.cents("123.45"), 12345)
        self.assertEqual(E.cents("$1,234.56"), 123456)
        self.assertEqual(E.cents("(100.00)"), -10000)
        self.assertEqual(E.cents("-50"), -5000)
        self.assertEqual(E.cents(0), 0)
        self.assertEqual(E.cents(10.0), 1000)
        self.assertIsNone(E.cents("abc"))
        self.assertIsNone(E.cents(None))
        self.assertIsNone(E.cents(""))
        # falsy-zero must be preserved, not dropped
        self.assertEqual(E.cents("0.00"), 0)

    def test_parse_date(self):
        self.assertEqual(E.parse_date("2024-03-05"), date(2024, 3, 5))
        self.assertEqual(E.parse_date("03/05/2024"), date(2024, 3, 5))
        self.assertEqual(E.parse_date("3/5/24"), date(2024, 3, 5))
        self.assertEqual(E.parse_date("2024-03-05T00:00:00.000+00:00"), date(2024, 3, 5))
        self.assertEqual(E.parse_date(45356), date(2024, 3, 5))  # Excel serial
        self.assertIsNone(E.parse_date("not a date"))
        self.assertIsNone(E.parse_date(None))

    def test_datetime_before_date(self):
        from datetime import datetime
        self.assertEqual(E.parse_date(datetime(2024, 3, 5, 12, 0)), date(2024, 3, 5))

    def test_znorm(self):
        self.assertEqual(E.znorm("ab-12_34"), "AB1234")
        self.assertEqual(E.znorm(" x.y "), "XY")

    def test_digit_runs(self):
        self.assertEqual(E.digit_runs("ab12345cd678", 5), {"12345"})
        self.assertEqual(E.digit_runs("12 34", 5), set())

    def test_payer_tokens_excludes_bai2_labels(self):
        toks = E.payer_tokens("SENDING CO NAME ACME WIDGETS COMPANY")
        self.assertIn("ACME", toks)
        self.assertIn("WIDGETS", toks)
        self.assertNotIn("COMPANY", toks)
        self.assertNotIn("NAME", toks)
        self.assertNotIn("SENDING", toks)

    def test_spn_of(self):
        self.assertEqual(E.spn_of("REC-SPN 12345-1"), "SPN12345")
        self.assertEqual(E.spn_of("nothing"), "")

    def test_is_mid(self):
        self.assertTrue(E.is_mid("8012345678"))
        self.assertTrue(E.is_mid("2000123456"))
        self.assertFalse(E.is_mid("6500000097"))  # Heartland excluded
        self.assertFalse(E.is_mid("1234567890"))

    def test_sibling(self):
        self.assertTrue(E.sibling("1234567", "1234568"))
        self.assertTrue(E.sibling("12345670", "12345699"))
        self.assertFalse(E.sibling("1234567", "1234567"))  # equal not sibling
        self.assertFalse(E.sibling("123", "124"))          # too short

    def test_signed_lag(self):
        self.assertEqual(E.signed_lag(date(2024, 3, 10), date(2024, 3, 5)), 5)
        self.assertIsNone(E.signed_lag(None, date(2024, 3, 5)))

    def test_reference_equal(self):
        self.assertTrue(E.reference_equal("ABC-123456", "abc123456"))
        self.assertTrue(E.reference_equal("REF987654", "XREF987654X"))  # containment >=6
        self.assertFalse(E.reference_equal("1234567", "1234568"))       # sibling conflict
        self.assertFalse(E.reference_equal("", "x"))

    def test_reference_tie(self):
        self.assertTrue(E.reference_tie("inv 45678 x", "y45678"))
        self.assertFalse(E.reference_tie("1234567", "1234568"))  # siblings never tie

    def test_date_bands(self):
        self.assertEqual(E.date_band(2), "STRONG")
        self.assertEqual(E.date_band(6), "MODERATE")
        self.assertEqual(E.date_band(12), "WEAK")
        self.assertEqual(E.date_band(20), "SUSPICIOUS")
        self.assertEqual(E.date_band(40), "REJECT")
        self.assertTrue(E.date_ok_merchant(3))
        self.assertFalse(E.date_ok_merchant(10))


class TestRouter(unittest.TestCase):
    def test_classify(self):
        self.assertEqual(E.classify_file("20240101_FHB_UTC_BSL_UNR.xlsx"), "BSL")
        self.assertEqual(E.classify_file("FHB_UTC_All_Data.xlsx"), "ALL_DATA")
        self.assertEqual(E.classify_file("Account_ST_open.xlsx"), "ST")
        self.assertEqual(E.classify_file("MET_All.xlsx"), "MET")
        self.assertEqual(E.classify_file("Receivables_Receipts.xlsx"), "RECEIPTS")
        self.assertEqual(E.classify_file("Edison_Payments.xlsx"), "EDISON_PAY")
        self.assertIsNone(E.classify_file("random.txt"))

    def test_bsl_excludes_all_data(self):
        # A file that is All_Data must not be misrouted to BSL.
        self.assertEqual(E.classify_file("FHB_UTC_BSL_All_Data.xlsx"), "ALL_DATA")

    def test_infer_account(self):
        self.assertEqual(E.infer_account("20240101_FHB_UTC_BSL.xlsx"), "FHB_UTC")
        self.assertEqual(E.infer_account("Regions_UTM_BSL.xlsx"), "REGIONS_UTM")


class TestBinder(unittest.TestCase):
    def test_content_first(self):
        rows = [
            ("Date", "Amount", "Reference"),
            ("2024-03-05", "100.00", "REF123456"),
            ("2024-03-06", "200.00", "REF223456"),
        ]
        specs = {
            "date": E._rs(True, ["Date"], E.pred_date),
            "amount": E._rs(True, ["Amount"], E.pred_signed_amount),
            "reference": E._rs(False, ["Reference"], E.pred_reference),
        }
        m, hi = E.bind_columns(rows, specs)
        self.assertEqual(hi, 0)
        self.assertEqual(m["date"], 0)
        self.assertEqual(m["amount"], 1)
        self.assertEqual(m["reference"], 2)

    def test_missing_required_raises(self):
        rows = [("Foo", "Bar"), ("x", "y"), ("a", "b")]
        specs = {"amount": E._rs(True, ["Amount"], E.pred_signed_amount)}
        with self.assertRaises(E.InvalidSourceData):
            E.bind_columns(rows, specs)


class TestPoolDedup(unittest.TestCase):
    def test_keep_largest_and_borrow_counterparty(self):
        # Classic total-plus-invoice-splits repeat: largest == sum(rest).
        e_total = E._mk_entry("T1", 10000, date(2024, 3, 5), "REF1", "", "AR", "UNR", True, "RECEIPTS")
        e_split1 = E._mk_entry("T1", 4000, date(2024, 3, 5), "REF1", "ACME CORP", "AR", "UNR", True, "RECEIPTS")
        e_split2 = E._mk_entry("T1", 6000, date(2024, 3, 5), "REF1", "", "AR", "UNR", True, "RECEIPTS")
        out = E._dedup_keep_largest([e_split1, e_total, e_split2], lambda e: e.id)
        self.assertEqual(len(out), 1)
        self.assertEqual(out[0].amount_cents, 10000)          # kept the total
        self.assertEqual(out[0].counterparty, "ACME CORP")     # borrowed

    def test_deposit_correction_detection(self):
        class _B:
            additional_info = ""
            line_info = ""
        b = _B()
        b.line_info = "Line 1 , 2026-06-10 DEPOSIT CORRECTION CREDIT"
        self.assertTrue(E._is_deposit_correction(b))
        b2 = _B()
        b2.additional_info = "CORRECTED DEPOSIT 0123"
        self.assertTrue(E._is_deposit_correction(b2))
        b3 = _B()
        b3.additional_info = "MERCHANT SERVICEDEPOSIT"
        b3.line_info = "Line 2"
        self.assertFalse(E._is_deposit_correction(b3))

    def test_amount_distinctive(self):
        self.assertTrue(E.amount_distinctive(245546969))   # 2,455,469.69
        self.assertTrue(E.amount_distinctive(-154638))     # (1,546.38)
        self.assertFalse(E.amount_distinctive(500000))     # 5,000.00 round
        self.assertFalse(E.amount_distinctive(6500))       # 65.00 too small
        self.assertFalse(E.amount_distinctive(-13000))     # (130.00)

    def test_convera_payables_only(self):
        class _B:
            amount_cents = -59690
            additional_info = "CONVERA TENN DEBITS 260703"
            line_info = "Line 59 , 2026-07-03"
            transaction_type = "Wire"
            reference_raw = ""
            customer_reference = ""
            account_servicer_reference = ""
        b = _B()
        ap = E._mk_entry("1394", -59690, date(2026, 7, 2), "1394", "Stichting EHEDG", "AP", "UNR", True, "ST")
        ext = E._mk_entry("999", -59690, date(2026, 7, 2), "999", "Someone", "EXT", "UNR", True, "MET")
        self.assertTrue(E._type_gate_ok(b, ap))
        self.assertFalse(E._type_gate_ok(b, ext))

    def test_ext_stale_candidate_bar(self):
        # Owner rule (2026-07-12): External STs entered 12+ days before the
        # BSL are barred even as Candidates; 8-11 days may still surface;
        # BSL-before-ST unbounded; non-EXT sources untouched.
        class _B:
            pass
        b = _B()
        b.date = date(2026, 6, 12)
        mk = lambda src_, d: E._mk_entry("X1", 1000, d, "R", "C", src_, "UNR", True, "ST")
        self.assertTrue(E._ext_stale_barred(b, mk("EXT", date(2026, 5, 31))))   # 12d stale
        self.assertTrue(E._ext_stale_barred(b, mk("EXT", date(2026, 1, 1))))    # very stale
        self.assertFalse(E._ext_stale_barred(b, mk("EXT", date(2026, 6, 1))))   # 11d stale
        self.assertFalse(E._ext_stale_barred(b, mk("EXT", date(2026, 7, 30))))  # ST after BSL
        self.assertFalse(E._ext_stale_barred(b, mk("AR", date(2025, 1, 1))))    # non-EXT
        e = mk("EXT", None)
        self.assertFalse(E._ext_stale_barred(b, e))                             # no date

    def test_chargeback_mid_gate(self):
        # Owner rule (2026-07-12): negative chargeback/merchant-fee lines
        # pair ONLY with STs carrying the SAME MID — wrong or absent MID is
        # barred even as a Candidate.
        class _B:
            amount_cents = -6000
            additional_info = ("MERCHANT SERVICECHARGEBACK2603178028920588"
                               "ENTRY DESC: CHARGEBACK ID NUMBER: 8028920588")
            line_info = "Line 129 , 2026-03-17"
            transaction_type = "Automated clearing house"
            reference_raw = "8028920588"
            customer_reference = ""
            account_servicer_reference = "8028920588"
        b = _B()
        wrong = E._mk_entry("1027659", -6000, date(2026, 6, 30), "8035758468", "FY26 CC chargeback", "EXT", "UNR", True, "MET")
        right = E._mk_entry("999999", -6000, date(2026, 3, 17), "8028920588", "Chargeback", "EXT", "UNR", True, "MET")
        nomid = E._mk_entry("888888", -6000, date(2026, 3, 17), "REF123456", "Something", "EXT", "UNR", True, "MET")
        self.assertTrue(E._is_card_fee_debit(b))
        self.assertFalse(E._type_gate_ok(b, wrong))   # different MID
        self.assertTrue(E._type_gate_ok(b, right))    # same MID
        self.assertFalse(E._type_gate_ok(b, nomid))   # no MID at all

    def test_distinct_receipts_sharing_label_all_kept(self):
        # Real FHB Master case: two receipts share one transaction-number
        # label but do NOT fit the total-plus-splits signature — both stay,
        # ids disambiguated by amount (base_id preserves the bridge join).
        a = E._mk_entry("SPN070326 ACH HRSA", 1820327, date(2026, 7, 3), "SPN070326 ACH HRSA", "HRSA", "AR", "UNR", True, "ST")
        b = E._mk_entry("SPN070326 ACH HRSA", 10241147, date(2026, 7, 3), "SPN070326 ACH HRSA", "HRSA", "AR", "UNR", True, "ST")
        out = E._dedup_keep_largest([a, b], lambda e: e.id)
        self.assertEqual(len(out), 2)
        self.assertEqual(sorted(e.id for e in out),
                         ["SPN070326 ACH HRSA [102411.47]",
                          "SPN070326 ACH HRSA [18203.27]"])
        self.assertTrue(all(e.base_id == "SPN070326 ACH HRSA" for e in out))


# ---- synthetic end-to-end fixture ------------------------------------

def _write_xlsx(path, sheets):
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)
    for title, rows in sheets:
        ws = wb.create_sheet(title=title[:31])
        for r in rows:
            ws.append(list(r))
    wb.save(path)


def _build_fhb_utc_inputs(d, bsl_name="20240101_FHB_UTC_BSL_UNR.xlsx",
                          st_name="20240101_FHB_UTC_Account_ST.xlsx"):
    """Synthetic FHB_UTC account: 2 matchable lines + 1 review line."""
    # BSL: three open bank lines for FHB_UTC.
    bsl = [
        ("Transaction Date", "Amount", "Line Number", "Account Servicer Reference",
         "Additional Information", "Transaction Code"),
        ("2024-03-10", "150.00", "1", "REF100200", "ACH CREDIT VENDOR", "142"),   # exact 1:1
        ("2024-03-12", "300.00", "2", "GRP500600", "DEPOSIT", "174"),             # 1:M group
        ("2024-03-15", "999.99", "3", "NADA000111", "UNKNOWN THING", "142"),      # review
    ]
    _write_xlsx(os.path.join(d, bsl_name), [("Exported", bsl)])
    # ST: open external transactions.
    st = [
        ("Transaction Date", "Amount", "Transaction Number", "Source", "Reference", "Counterparty"),
        ("2024-03-09", "150.00", "ST1", "EXT", "REF100200", "VENDOR INC"),        # matches line 1
        ("2024-03-11", "120.00", "ST2", "EXT", "GRP500600", "PAYER A"),           # part of group
        ("2024-03-11", "180.00", "ST3", "EXT", "GRP500600", "PAYER A"),           # part of group
        ("2024-03-01", "999.99", "ST4", "GL", "NADA000111", "JOURNAL"),           # journal: never matches
    ]
    _write_xlsx(os.path.join(d, st_name), [("Exported", st)])
    return "FHB_UTC"


class TestEndToEnd(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.mkdtemp()
        self.out = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.d, ignore_errors=True)
        shutil.rmtree(self.out, ignore_errors=True)

    def _build_inputs(self):
        return _build_fhb_utc_inputs(self.d)

    def test_forward_and_audit(self):
        account = self._build_inputs()
        runlog = E.run(self.d, self.out, present=True)
        self.assertEqual(runlog["account"], "FHB_UTC")
        self.assertEqual(runlog["bsl_count"], 3)
        summ = runlog["recon_summary"]
        # line 1 -> Match (exact 1:1); line 2 -> Match (1:M group); line 3 -> Review (journal only)
        self.assertEqual(summ["matches"], 2)
        self.assertEqual(summ["reviews"], 1)
        # audit must pass
        self.assertEqual(runlog["audit"]["status"], "PASS",
                         msg=str(runlog["audit"].get("failures")))

    def test_conservation_every_bsl_once(self):
        self._build_inputs()
        runlog = E.run(self.d, self.out, present=True)
        s = runlog["recon_summary"]
        self.assertEqual(s["matches"] + s["candidates"] + s["reviews"], runlog["bsl_count"])

    def test_journal_never_matches(self):
        self._build_inputs()
        runlog = E.run(self.d, self.out, present=True)
        # The 999.99 line has only a GL/journal counterpart -> must be Review.
        recon_path = runlog["recon_workbook"]
        tabs, _ = A._read_output_tabs(recon_path)
        review = [r for r in tabs["Review Notes"][3:] if any(A._N(c) for c in r)]
        self.assertTrue(any("999.99" in A._N(r[2]) for r in review))


class TestColumnRobustness(unittest.TestCase):
    """Column movement/renaming/insertion must never change results silently:
    identical output for benign rearrangement, loud named errors for genuine
    ambiguity. Pins the guarantees proven by the mutation sweep."""

    def setUp(self):
        self.a = tempfile.mkdtemp()   # baseline inputs
        self.b = tempfile.mkdtemp()   # mutated inputs
        self.oa = tempfile.mkdtemp()
        self.ob = tempfile.mkdtemp()

    def tearDown(self):
        for d in (self.a, self.b, self.oa, self.ob):
            shutil.rmtree(d, ignore_errors=True)

    def test_leading_date_key_validates_dates(self):
        # An 8-digit account number is not a date; the real stamp wins.
        self.assertEqual(E._leading_date_key("acct_99999999_st_20250101.xlsx"),
                         "20250101")
        self.assertEqual(E._leading_date_key("20240101_FHB_UTC_BSL.xlsx"),
                         "20240101")
        self.assertEqual(E._leading_date_key("undated.xlsx"), "00000000")

    def test_bind_columns_no_header_raises(self):
        rows = [("Foo", "Bar"), ("2024-01-01", "10.00")]
        specs = {"amount": E._rs(True, ["Amount"], E.pred_signed_amount)}
        with self.assertRaises(E.InvalidSourceData):
            E.bind_columns(rows, specs)

    def test_optional_role_blind_tie_unbound(self):
        # Two columns identical on content and header: an optional role must
        # go unbound, never bind leftmost-by-position.
        rows = [("Date", "Amount", "X", "Y"),
                ("2024-01-01", "10.00", "REF1", "REF9"),
                ("2024-01-02", "20.00", "REF2", "REF8")]
        specs = {
            "date": E._rs(True, ["Date"], E.pred_date),
            "amount": E._rs(True, ["Amount"], E.pred_signed_amount),
            "reference": E._rs(False, ["Reference"], E.pred_reference),
        }
        m, _hi = E.bind_columns(rows, specs)
        self.assertNotIn("reference", m)

    def test_pred_met_description(self):
        self.assertTrue(E.pred_met_description("d:8812345 | r:991 | UT FOUNDATION"))
        self.assertFalse(E.pred_met_description("d:63363 | r:197960"))  # ID-column stub
        self.assertFalse(E.pred_met_description("something | else"))
        self.assertFalse(E.pred_met_description("see attached"))

    def test_mid_master_multi_mid_and_conflict(self):
        gl = "01-1234567-123456-123456"
        p = os.path.join(self.a, "MID_Master.xlsx")
        _write_xlsx(p, [("Sheet1", [("8012345678", "8098765432", gl)])])
        rf = E.RoutedFile("MID_MASTER", p, "MID_Master.xlsx", "all")
        out = E.load_mid_master(rf)
        # Every MID in the row maps; no rightmost-wins column dependence.
        self.assertEqual(out["mid_gl"]["8012345678"], gl)
        self.assertEqual(out["mid_gl"]["8098765432"], gl)
        # Two distinct GL strings in one row: genuinely ambiguous -> loud.
        gl2 = "02-7654321-654321-654321"
        p2 = os.path.join(self.a, "MID_Master_bad.xlsx")
        _write_xlsx(p2, [("Sheet1", [("8012345678", gl, gl2)])])
        with self.assertRaises(E.InvalidSourceData):
            E.load_mid_master(E.RoutedFile("MID_MASTER", p2, "MID_Master_bad.xlsx", "all"))

    def test_multi_file_newest_wins_and_tie_fails_loud(self):
        # Distinct date stamps: newest wins, runs clean.
        _build_fhb_utc_inputs(self.a)
        _build_fhb_utc_inputs(self.a, bsl_name="20230101_FHB_UTC_BSL_UNR.xlsx",
                              st_name="20230101_FHB_UTC_Account_ST.xlsx")
        runlog = E.run(self.a, self.oa, present=True)
        self.assertEqual(runlog["roles_bound"]["BSL"]["file"],
                         "20240101_FHB_UTC_BSL_UNR.xlsx")
        # Same date stamp on two BSL files: one would be silently ignored ->
        # must fail loud instead.
        _build_fhb_utc_inputs(self.b)
        _build_fhb_utc_inputs(self.b, bsl_name="20240101_v2_FHB_UTC_BSL_UNR.xlsx",
                              st_name="20240101_v2_FHB_UTC_Account_ST.xlsx")
        with self.assertRaises(E.InvalidSourceData):
            E.run(self.b, self.ob, present=True)


    def test_column_shuffle_end_to_end_identical(self):
        # The core guarantee: permuted columns + preamble rows + a decoy
        # constant-date column produce cell-for-cell identical output.
        account = _build_fhb_utc_inputs(self.a)
        base = E.run(self.a, self.oa, present=True)

        def _mutate(src, dst, perm, preamble, decoy_at):
            from openpyxl import load_workbook, Workbook
            wb = load_workbook(src)
            rows = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
            out = Workbook(); out.remove(out.active)
            ws = out.create_sheet("Exported")
            for p in preamble:
                ws.append(p)
            for i, r in enumerate(rows):
                new = [r[j] for j in perm]
                if decoy_at is not None:
                    new.insert(decoy_at, "As Of" if i == 0 else "2024-04-01")
                ws.append(new)
            out.save(dst)

        _mutate(os.path.join(self.a, "20240101_FHB_UTC_BSL_UNR.xlsx"),
                os.path.join(self.b, "20240101_FHB_UTC_BSL_UNR.xlsx"),
                [5, 3, 1, 0, 4, 2], [("University of Tennessee",), ()], 2)
        _mutate(os.path.join(self.a, "20240101_FHB_UTC_Account_ST.xlsx"),
                os.path.join(self.b, "20240101_FHB_UTC_Account_ST.xlsx"),
                [5, 4, 3, 2, 1, 0], [], None)
        mut = E.run(self.b, self.ob, present=True)

        self.assertEqual(mut["audit"]["status"], "PASS",
                         msg=str(mut["audit"].get("failures")))
        self.assertEqual(base["recon_summary"], mut["recon_summary"])
        ta, _ = A._read_output_tabs(base["recon_workbook"])
        tb, _ = A._read_output_tabs(mut["recon_workbook"])
        self.assertEqual(ta, tb)

    def test_all_data_never_loaded_and_never_crashes(self):
        # ALL_DATA is recognized but not loaded (forward-only engine): even
        # two undated All_Data workbooks (a tie the newest-wins check would
        # reject for loaded roles) must not abort the forward run.
        _build_fhb_utc_inputs(self.a)
        _write_xlsx(os.path.join(self.a, "FHB_UTC_All_Data.xlsx"),
                    [("Recon History", [("junk",)])])
        _write_xlsx(os.path.join(self.a, "FHB_UTC_v2_All_Data.xlsx"),
                    [("Recon History", [("junk",)])])
        runlog = E.run(self.a, self.oa, present=True)
        self.assertEqual(runlog["audit"]["status"], "PASS")
        self.assertIn("not loaded", runlog["all_data"])
        self.assertNotIn("ALL_DATA", runlog["roles_bound"])


    def test_audit_binds_same_columns_as_engine(self):
        # 'Post Date' header plus an inserted date-parseable decoy column:
        # both engine and audit must bind the exact-alias column -> PASS.
        bsl = [
            ("Trade Date", "Post Date", "Amount", "Line Number",
             "Account Servicer Reference", "Additional Information", "Transaction Code"),
            ("2024-03-09", "2024-03-10", "150.00", "1", "REF100200", "ACH CREDIT", "142"),
            ("2024-03-11", "2024-03-12", "300.00", "2", "GRP500600", "DEPOSIT", "174"),
        ]
        st = [
            ("Transaction Date", "Amount", "Transaction Number", "Source", "Reference", "Counterparty"),
            ("2024-03-09", "150.00", "ST1", "EXT", "REF100200", "VENDOR INC"),
            ("2024-03-11", "120.00", "ST2", "EXT", "GRP500600", "PAYER A"),
            ("2024-03-11", "180.00", "ST3", "EXT", "GRP500600", "PAYER A"),
        ]
        _write_xlsx(os.path.join(self.a, "20240101_FHB_UTC_BSL_UNR.xlsx"),
                    [("Exported", bsl)])
        _write_xlsx(os.path.join(self.a, "20240101_FHB_UTC_Account_ST.xlsx"),
                    [("Exported", st)])
        runlog = E.run(self.a, self.oa, present=True)
        self.assertEqual(runlog["roles_bound"]["BSL"]["columns"]["date"], 1)
        self.assertEqual(runlog["audit"]["status"], "PASS",
                         msg=str(runlog["audit"].get("failures")))


class TestPerRunScript(unittest.TestCase):
    """run_recon.py: stage one upload -> immutable run folder -> engine."""

    def setUp(self):
        self.upload = tempfile.mkdtemp()
        self.runs = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.upload, ignore_errors=True)
        shutil.rmtree(self.runs, ignore_errors=True)

    def test_staged_name_prefix_rules(self):
        # Hex upload prefix stripped; plausible YYYYMMDD date prefix kept.
        self.assertEqual(R.staged_name("933782d6-FHB_UTC_BSL.xlsx"),
                         "FHB_UTC_BSL.xlsx")
        self.assertEqual(R.staged_name("20240101-FHB_UTC_BSL.xlsx"),
                         "20240101-FHB_UTC_BSL.xlsx")
        # All-digit but NOT a plausible date (month 20) -> an unlucky random
        # hex prefix; must be stripped or it poisons newest-wins ordering.
        self.assertEqual(R.staged_name("30201090-FHB_UTC_BSL.xlsx"),
                         "FHB_UTC_BSL.xlsx")
        self.assertEqual(R.staged_name("FHB_UTC_BSL.xlsx"), "FHB_UTC_BSL.xlsx")
        self.assertEqual(R.staged_name("933782d6-x.xlsx", strip_prefix=False),
                         "933782d6-x.xlsx")

    def test_perform_run_end_to_end(self):
        _build_fhb_utc_inputs(
            self.upload,
            bsl_name="933782d6-20240101_FHB_UTC_BSL_UNR.xlsx",
            st_name="ab12cd34-20240101_FHB_UTC_Account_ST.xlsx")
        # An unrouted spreadsheet and a non-spreadsheet ride along.
        _write_xlsx(os.path.join(self.upload, "random_notes.xlsx"),
                    [("Sheet1", [("a",)])])
        with open(os.path.join(self.upload, "readme.txt"), "w") as fh:
            fh.write("not a spreadsheet")

        code, report = R.perform_run([self.upload], runs_root=self.runs,
                                     run_id="run_test")
        self.assertEqual(code, 0)
        self.assertEqual(report["account"], "FHB_UTC")
        self.assertEqual(report["audit"], "PASS")

        run_dir = os.path.join(self.runs, "run_test")
        # Upload prefixes stripped in staging.
        self.assertTrue(os.path.exists(os.path.join(
            run_dir, "input", "20240101_FHB_UTC_BSL_UNR.xlsx")))
        self.assertTrue(os.path.exists(report["recon_workbook"]))
        self.assertTrue(os.path.exists(report["runlog"]))

        import json
        with open(os.path.join(run_dir, "manifest.json")) as fh:
            manifest = json.load(fh)
        roles = {e["staged_as"]: e["role"] for e in manifest["files"]}
        self.assertEqual(roles["20240101_FHB_UTC_BSL_UNR.xlsx"], "BSL")
        self.assertEqual(roles["20240101_FHB_UTC_Account_ST.xlsx"], "ST")
        self.assertEqual(roles["random_notes.xlsx"], "UNROUTED")
        self.assertTrue(all(len(e["sha256"]) == 64 for e in manifest["files"]))
        self.assertTrue(any("random_notes.xlsx" in w for w in manifest["warnings"]))
        self.assertEqual(len(manifest["ignored_non_spreadsheets"]), 1)
        cv = manifest["code_versions"]
        self.assertEqual(len(cv["recon_engine.py"]), 64)
        self.assertEqual(len(cv["recon_audit.py"]), 64)

    def test_individual_file_args_and_parent_dir_dedup(self):
        # Documented primary usage: file arguments, including a file passed
        # both explicitly and via its parent directory (one source, no
        # spurious self-collision).
        _build_fhb_utc_inputs(self.upload)
        bsl = os.path.join(self.upload, "20240101_FHB_UTC_BSL_UNR.xlsx")
        st = os.path.join(self.upload, "20240101_FHB_UTC_Account_ST.xlsx")
        code, report = R.perform_run([bsl, st, self.upload, bsl],
                                     runs_root=self.runs, run_id="rfiles")
        self.assertEqual(code, 0)
        self.assertEqual(report["audit"], "PASS")

    def test_audit_fail_exits_2_and_quarantines(self):
        # Force the independent audit to FAIL: perform_run must recover the
        # runlog (written before the gate raises), report FAIL, exit 2 —
        # and the workbooks stay on disk for forensics.
        import recon_audit
        _build_fhb_utc_inputs(self.upload)
        real_audit = recon_audit.audit
        recon_audit.audit = lambda *a, **k: {
            "status": "FAIL", "checks": {}, "failures": ["forced"]}
        try:
            self.assertEqual(
                R.main([self.upload, "--runs-root", self.runs,
                        "--run-id", "rfail"]), 2)
            out = os.path.join(self.runs, "rfail", "outputs")
            self.assertTrue(os.path.exists(
                os.path.join(out, "FHB_UTC_reconciliation.xlsx")))
            # Gate off: same failing audit -> exit 0, run completes.
            code, report = R.perform_run([self.upload], runs_root=self.runs,
                                         run_id="rfail2", present=False)
            self.assertEqual(code, 0)
            self.assertEqual(report["audit"], "FAIL")
        finally:
            recon_audit.audit = real_audit

    def test_no_strip_flag_via_cli(self):
        _build_fhb_utc_inputs(
            self.upload,
            bsl_name="933782d6-20240101_FHB_UTC_BSL_UNR.xlsx",
            st_name="ab12cd34-20240101_FHB_UTC_Account_ST.xlsx")
        self.assertEqual(
            R.main([self.upload, "--runs-root", self.runs, "--run-id", "rk",
                    "--no-strip-upload-prefix"]), 0)
        self.assertTrue(os.path.exists(os.path.join(
            self.runs, "rk", "input",
            "933782d6-20240101_FHB_UTC_BSL_UNR.xlsx")))

    def test_missing_bsl_fails_loud(self):
        _build_fhb_utc_inputs(self.upload)
        os.remove(os.path.join(self.upload, "20240101_FHB_UTC_BSL_UNR.xlsx"))
        with self.assertRaises(R.PerRunError):
            R.perform_run([self.upload], runs_root=self.runs, run_id="r1")

    def test_mixed_accounts_fail_loud(self):
        _build_fhb_utc_inputs(self.upload)
        _build_fhb_utc_inputs(self.upload,
                              bsl_name="20240101_Regions_UTM_BSL_UNR.xlsx",
                              st_name="20240101_Regions_UTM_Account_ST.xlsx")
        with self.assertRaises(R.PerRunError):
            R.perform_run([self.upload], runs_root=self.runs, run_id="r1")

    def test_mixed_account_st_alone_fails_loud(self):
        # A wrong-account ST with the right-account BSL would cross-pollute
        # the pool; the guard must span every routed file, not just BSL.
        _build_fhb_utc_inputs(self.upload,
                              st_name="20240101_Regions_UTM_Account_ST.xlsx")
        with self.assertRaises(R.PerRunError):
            R.perform_run([self.upload], runs_root=self.runs, run_id="r1")

    def test_case_only_collision_fails_loud(self):
        other = tempfile.mkdtemp()
        try:
            _build_fhb_utc_inputs(self.upload)
            _build_fhb_utc_inputs(other,
                                  bsl_name="20240101_fhb_utc_bsl_unr.xlsx",
                                  st_name="20240101_FHB_UTC_ACCOUNT_st.xlsx")
            with self.assertRaises(R.PerRunError):
                R.perform_run([self.upload, other], runs_root=self.runs,
                              run_id="r1")
        finally:
            shutil.rmtree(other, ignore_errors=True)

    def test_failed_preflight_frees_run_id(self):
        # An unusable upload must not leave a half-built folder or burn the id.
        _build_fhb_utc_inputs(self.upload)
        os.remove(os.path.join(self.upload, "20240101_FHB_UTC_BSL_UNR.xlsx"))
        with self.assertRaises(R.PerRunError):
            R.perform_run([self.upload], runs_root=self.runs, run_id="r1")
        self.assertFalse(os.path.exists(os.path.join(self.runs, "r1")))
        # Same id now works with a good upload.
        _build_fhb_utc_inputs(self.upload)
        code, _ = R.perform_run([self.upload], runs_root=self.runs,
                                run_id="r1")
        self.assertEqual(code, 0)

    def test_bad_run_ids_rejected(self):
        _build_fhb_utc_inputs(self.upload)
        for bad in ("..", ".", "a/b", ""):
            with self.assertRaises(R.PerRunError):
                R.perform_run([self.upload], runs_root=self.runs, run_id=bad)

    def test_nonexistent_path_fails_loud(self):
        with self.assertRaises(R.PerRunError):
            R.perform_run([os.path.join(self.upload, "nope.xlsx")],
                          runs_root=self.runs, run_id="r1")

    def test_subfolder_with_spreadsheets_warned(self):
        _build_fhb_utc_inputs(self.upload)
        nested = os.path.join(self.upload, "nested")
        os.makedirs(nested)
        _write_xlsx(os.path.join(nested, "20240201_FHB_UTC_BSL_UNR.xlsx"),
                    [("Exported", [("a",)])])
        code, report = R.perform_run([self.upload], runs_root=self.runs,
                                     run_id="rsub")
        self.assertEqual(code, 0)
        import json
        with open(os.path.join(report["run_dir"], "manifest.json")) as fh:
            manifest = json.load(fh)
        self.assertIn(nested, manifest["ignored_non_spreadsheets"])
        self.assertTrue(any("NOT read" in w for w in manifest["warnings"]))

    def test_collision_fails_loud(self):
        other = tempfile.mkdtemp()
        try:
            _build_fhb_utc_inputs(self.upload)
            _build_fhb_utc_inputs(other)  # same filenames, different folder
            with self.assertRaises(R.PerRunError):
                R.perform_run([self.upload, other], runs_root=self.runs,
                              run_id="r1")
        finally:
            shutil.rmtree(other, ignore_errors=True)

    def test_runs_are_immutable(self):
        _build_fhb_utc_inputs(self.upload)
        code, _ = R.perform_run([self.upload], runs_root=self.runs,
                                run_id="r1")
        self.assertEqual(code, 0)
        with self.assertRaises(R.PerRunError):
            R.perform_run([self.upload], runs_root=self.runs, run_id="r1")

    def test_main_exit_codes(self):
        # Empty upload -> exit 1 (fail loud, no run executed).
        empty = tempfile.mkdtemp()
        try:
            self.assertEqual(
                R.main([empty, "--runs-root", self.runs, "--run-id", "r1"]), 1)
        finally:
            shutil.rmtree(empty, ignore_errors=True)
        # Good upload -> exit 0.
        _build_fhb_utc_inputs(self.upload)
        self.assertEqual(
            R.main([self.upload, "--runs-root", self.runs, "--run-id", "r2"]), 0)


class TestRealDataShapes(unittest.TestCase):
    """Regression tests for the real FHB Master export shapes (2026-07-10
    validation run): router tokens, integer cells, NA placeholders, the MET
    scope join + ST bridge, and the ORT deposit-group pass."""

    def setUp(self):
        self.d = tempfile.mkdtemp()
        self.out = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.d, ignore_errors=True)
        shutil.rmtree(self.out, ignore_errors=True)

    def test_router_real_filenames(self):
        # "_st" must not swallow All_Status / Rosetta_Stone.
        cases = {
            "20260710_Oracle_CM_FHB_Master_BSL_UNR.xlsx": "BSL",
            "20260710_Oracle_CM_FHB_Master_ST_UNR.xlsx": "ST",
            "20260710_Oracle_OTBI_MET_All_Accounts_All_Status.xlsx": "MET",
            "20260710_FHB_Master_BAI2.xlsx": "BAI2",
            "ORT_Misc_All.xlsb": "ORT_MISC",
            "UT_Recon_SPN_Receivables_Rosetta_Stone_Part2_1.xlsx": "RELATIONSHIP_MAP",
            "UT_Recon_Data_Relationship_Map_20260702_1.xlsx": "RELATIONSHIP_MAP",
        }
        for name, want in cases.items():
            self.assertEqual(E.classify_file(name), want, msg=name)
        self.assertEqual(E.infer_account("20260710_FHB_Master_BAI2.xlsx"),
                         "FHB_MASTER")

    def test_integer_reference_cells_bind(self):
        # Oracle exports deliver references/ids as raw ints; parse_date must
        # not swallow small ints as Excel serials and poison the binder.
        self.assertTrue(E.pred_reference(1390))
        self.assertTrue(E.pred_reference("852256"))
        self.assertFalse(E.pred_reference("2026-07-09"))
        from datetime import datetime as _dt
        self.assertFalse(E.pred_reference(_dt(2026, 7, 9)))
        rows = [
            ("Date", "Amount (USD)", "Reference", "Transaction Number", "Source",
             "Payment Reference Number"),
            ("2026-07-10", "10.00", 1390, 17, "External", ""),
            ("2026-07-09", "20.00", 104000120, 1027519, "External", 260238),
        ]
        m, _hi = E.bind_columns(rows, E.ST_ROLES, filename="st.xlsx")
        self.assertEqual(m["transaction_number"], 3)
        self.assertEqual(m["reference"], 2)

    def test_na_placeholder_never_ties(self):
        self.assertEqual(E.clean_ref("NA"), "")
        self.assertEqual(E.clean_ref("n/a"), "")
        self.assertEqual(E.clean_ref("852256"), "852256")
        a = E._mk_entry("T1", 100, date(2026, 7, 1), "NA", "", "EXT", "UNR", True, "ST")
        self.assertEqual(a.reference, "")
        self.assertFalse(E.reference_equal("NA", a.reference))

    def test_account_of_bank_name(self):
        self.assertEqual(E.account_of_bank_name("FHB - Master Account"), "FHB_MASTER")
        self.assertEqual(E.account_of_bank_name("Regions - UTM"), "REGIONS_UTM")
        self.assertIsNone(E.account_of_bank_name("TRUIST BANK - Chattanooga"))

    def _met_rows(self):
        return [
            ("CBE_BANK_ACCOUNT_NAME", "CET_REFERENCE_TEXT", "CET_STATUS",
             "CET_TRANSACTION_DATE", "CET_TRANSACTION_ID", "AMOUNT",
             "TRANSACTION_DATE", "CET_DESCRIPTION", "DEPOSIT_ID", "RECEIPT_ID"),
            # our account: deposit 900 = two open receipts 400 + 600
            ("FHB - Master Account", "DEPREF77", "UNR", "2026-07-01", 501,
             "400.00", "2026-07-01", "d:900 | r:11 | PAYER A", 900, 11),
            ("FHB - Master Account", "DEPREF77", "UNR", "2026-07-01", 502,
             "600.00", "2026-07-01", "d:900 | r:12 | PAYER A", 900, 12),
            # same trx id as the ST export -> must bridge, not duplicate
            ("FHB - Master Account", "REF100200", "UNR", "2026-07-02", 601,
             "150.00", "2026-07-02", "d:901 | r:13 | VENDOR", 901, 13),
            # other account: identical amounts — must never enter the pool
            ("FHB - UTC", "DEPREF77", "UNR", "2026-07-01", 701,
             "1000.00", "2026-07-01", "d:902 | r:14 | PAYER A", 902, 14),
        ]

    def _build(self, bsl_rows):
        _write_xlsx(os.path.join(self.d, "20260710_FHB_Master_BSL_UNR.xlsx"),
                    [("Exported", bsl_rows)])
        st = [
            ("Date", "Amount (USD)", "Reference", "Transaction Number", "Source", "Counterparty"),
            ("2026-07-02", "150.00", "REF100200", 601, "External", "VENDOR INC"),
        ]
        _write_xlsx(os.path.join(self.d, "20260710_FHB_Master_ST_UNR.xlsx"),
                    [("Exported", st)])
        _write_xlsx(os.path.join(self.d, "20260710_MET_All_Accounts.xlsx"),
                    [("Miscellaneous External Transact", self._met_rows())])

    def test_met_scope_bridge_and_deposit_group(self):
        bsl = [
            ("Date", "Amount (USD)", "Reference", "Additional Information",
             "Account Servicer Reference", "Transaction Type", "Statement", "Transaction Code"),
            # bundled deposit line: no reference, sums d:900 (400+600)
            ("2026-07-03", "1000.00", "NA", "DEPOSIT 011", "NA",
             "Miscellaneous", "Line 1 , 2026-07-03", "174"),
            # exact 1:1 to the bridged ST
            ("2026-07-03", "150.00", "REF100200", "ACH CREDIT", "REF100200",
             "Automated clearing house", "Line 2 , 2026-07-03", "142"),
        ]
        self._build(bsl)
        runlog = E.run(self.d, self.out, present=True)
        self.assertEqual(runlog["audit"]["status"], "PASS",
                         msg=str(runlog["audit"].get("failures")))
        # scope: 3 of 4 MET rows are ours; the FHB-UTC row is excluded
        self.assertEqual(runlog["met_scope"], {"rows_total": 4, "rows_in_account": 3})
        # bridge: trx 601 exists in both ST and MET -> one pool entry
        self.assertEqual(runlog["met_bridged_to_st"], 1)
        self.assertEqual(runlog["pool_total"], 3)
        # ACH line -> P3 exact Match; deposit line -> deposit-type consistency
        # confers NOTHING (owner doctrine 2026-07-11): the exact sum alone
        # surfaces as an amount-only Candidate, never a Match.
        self.assertEqual(runlog["recon_summary"],
                         {"matches": 1, "candidates": 1, "reviews": 0})
        self.assertEqual(runlog["forward_pass_counts"].get("P4_deposit_group"), 1)
        tabs_c, _ = A._read_output_tabs(runlog["recon_workbook"])
        cand = [r for r in tabs_c["Candidate Matches"][3:] if any(A._N(c) for c in r)]
        self.assertIn("AMOUNT_ONLY_GROUP", A._N(cand[0][A.COL_EXPL]))
        # the bridged entry carries the ORT coordinates
        tabs, _ = A._read_output_tabs(runlog["recon_workbook"])
        match_rows = [r for r in tabs["Matches"][3:] if any(A._N(c) for c in r)]
        by_amount = {A._N(r[2]): r for r in match_rows}
        self.assertEqual(A._N(by_amount["150.00"][A.COL_DEP]), "901")  # ORT d:
        self.assertEqual(A._N(by_amount["150.00"][A.COL_REC]), "13")   # ORT r:

    def test_deposit_auto_rec_split_is_candidate(self):
        rows = self._met_rows()
        rows[2] = ("FHB - Master Account", "DEPREF77", "REC", "2026-07-01", 502,
                   "600.00", "2026-07-01", "d:900 | r:12 | PAYER A", 900, 12)
        _write_xlsx(os.path.join(self.d, "20260710_MET_All_Accounts.xlsx"),
                    [("Miscellaneous External Transact", rows)])
        bsl = [
            ("Date", "Amount (USD)", "Reference", "Additional Information",
             "Account Servicer Reference", "Transaction Type", "Statement", "Transaction Code"),
            ("2026-07-03", "1000.00", "NA", "DEPOSIT 011", "NA",
             "Miscellaneous", "Line 1 , 2026-07-03", "174"),
        ]
        _write_xlsx(os.path.join(self.d, "20260710_FHB_Master_BSL_UNR.xlsx"),
                    [("Exported", bsl)])
        runlog = E.run(self.d, self.out, present=True)
        self.assertEqual(runlog["recon_summary"]["candidates"], 1)
        tabs, _ = A._read_output_tabs(runlog["recon_workbook"])
        cand = [r for r in tabs["Candidate Matches"][3:] if any(A._N(c) for c in r)]
        self.assertIn("already-closed", A._N(cand[0][A.COL_EXPL]))


    def test_router_token_hardening(self):
        # Short bare tokens match whole name segments only.
        self.assertIsNone(E.classify_file("Market_Analysis.xlsx"))   # 'ar' in 'Market'
        self.assertIsNone(E.classify_file("Smart_Data.xlsx"))
        self.assertEqual(E.classify_file("ORT_Departments.xlsx"), "DEPT_INFO")
        self.assertEqual(E.classify_file("ORT_Chart_Of_Accounts.xlsx"), "CHART_OF_ACCOUNTS")
        self.assertEqual(E.classify_file("ORT_AR_All.xlsb"), "ORT_AR")
        self.assertEqual(E.classify_file("UT_MID_Master_Consolidated.xlsx"), "MID_MASTER")
        self.assertEqual(E.classify_file("20260710_FHB_Master_BAI2.xlsx"), "BAI2")  # bai+digits

    def test_bai2_enrichment(self):
        bsl = [
            ("Date", "Amount (USD)", "Reference", "Additional Information",
             "Account Servicer Reference", "Transaction Type", "Statement", "Transaction Code"),
            ("2026-07-03", "150.00", "NA", "MERCHANT SERVICEDEPOSIT", "NA",
             "Automated clearing house", "Line 1 , 2026-07-03", "142"),
            ("2026-07-04", "75.00", "NA", "AMBIG", "NA",
             "Automated clearing house", "Line 2 , 2026-07-04", "142"),
        ]
        _write_xlsx(os.path.join(self.d, "20260710_FHB_Master_BSL_UNR.xlsx"),
                    [("Exported", bsl)])
        bai = [
            ("Post Date", "Transaction Description", "Amount", "Debit/Credit",
             "Bank Reference", "Customer Reference", "BAI Code", "DETAIL1", "DETAIL2"),
            ("2026-07-03", "ACH CREDIT RECEIVED", "150.00", "Credit",
             "ACH123", "8036121500", "142", "MERCHANT SERVICEDEPOSIT   26", "FULL ADDENDA TEXT"),
            # two BAI2 rows tie on (date, cents) with unrelated details: never guess
            ("2026-07-04", "X", "75.00", "Credit", "B1", "", "142", "ROW ONE", ""),
            ("2026-07-04", "Y", "75.00", "Credit", "B2", "", "142", "ROW TWO", ""),
        ]
        _write_xlsx(os.path.join(self.d, "20260710_FHB_Master_BAI2.xlsx"),
                    [("CSVEXP", bai)])
        st = [("Date", "Amount (USD)", "Reference", "Transaction Number", "Source", "Counterparty"),
              ("2026-07-02", "150.00", "8036121500", 601, "External", "V")]
        _write_xlsx(os.path.join(self.d, "20260710_FHB_Master_ST_UNR.xlsx"),
                    [("Exported", st)])
        runlog = E.run(self.d, self.out, present=True)
        self.assertEqual(runlog["bai2_enrichment"],
                         {"joined": 1, "ambiguous": 1, "no_hit": 0})
        # The BAI2 Customer Reference (a MID) reached the line's addenda:
        # the line classifies MERCHANT and the MID ties the ST reference.
        self.assertEqual(runlog["bsl_by_lane"].get("MERCHANT"), 1)

    def test_mid_directory(self):
        _write_xlsx(os.path.join(self.d, "ORT_Departments.xlsx"),
                    [("Report (3)", [
                        ("Campus Name", "Department Name", "Campus Bank Account",
                         "Dept Bank Account", "Campus Bank Name", "Dept Bank Name", "Credit Card Mid"),
                        ("Knoxville", "Ticket Office", "1", "2",
                         "FHB - Master Account", "FHB - UTIA", "8036121500"),
                        ("Knoxville", "No Mid Dept", "1", "2",
                         "FHB - Master Account", "FHB - Master Account", "N/A"),
                    ])])
        rf = E.RoutedFile("DEPT_INFO", os.path.join(self.d, "ORT_Departments.xlsx"),
                          "ORT_Departments.xlsx", "Report")
        rows, _ = E.read_rows(rf)
        m, hi = E.bind_columns(rows, E.DEPT_INFO_ROLES, filename=rf.filename)
        loaded = {"DEPT_INFO": {"rows": rows, "map": m, "header_index": hi}}
        d = E._mid_directory(loaded, "FHB_MASTER")
        self.assertEqual(d["8036121500"]["department"], "Ticket Office")
        self.assertEqual(d["8036121500"]["home_account"], "FHB_UTIA")
        self.assertEqual(len(d), 1)  # N/A row never enters




    def test_directional_date_doctrine(self):
        # BSL may PRECEDE the ST by any amount (entry lag, no ceiling);
        # BSL trailing the ST beyond the band is stale — never a Match.
        self.assertTrue(E.date_ok_directional(-300))   # ST after BSL: valid
        self.assertTrue(E.date_ok_directional(0))
        self.assertTrue(E.date_ok_directional(7))
        self.assertFalse(E.date_ok_directional(8))     # ST 8+ days before BSL
        self.assertFalse(E.date_ok_directional(300))
        self.assertTrue(E.date_ok_directional(None))   # unknown dates not gated

    def test_type_gate(self):
        b = E.make_bsl("1", date(2026, 7, 1), 1000, "R", "R", "", "Miscellaneous", "399")
        cc = E._mk_entry("S1", 1000, date(2026, 7, 1), "R", "", "EXT", "UNR",
                         True, "ST", transaction_type="Credit Card")
        eft = E._mk_entry("S2", 1000, date(2026, 7, 1), "R", "", "EXT", "UNR",
                          True, "ST", transaction_type="EFT")
        self.assertFalse(E._type_gate_ok(b, cc))
        self.assertTrue(E._type_gate_ok(b, eft))


if __name__ == "__main__":
    unittest.main(verbosity=2)
