#!/usr/bin/env python3
"""
Unit tests for the reconciliation engine.

Covers Section 7 primitives, Section 4 router, Section 5 binder, Section 8
pool dedup, and an end-to-end synthetic run through the forward/backward
pipeline + independent audit (Section 16 build order steps 1-8, validated on
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
        self.assertTrue(E.date_ok_state(-100))     # BSL precedes ST: no ceiling
        self.assertFalse(E.date_ok_state(25))       # ST precedes BSL >20d
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
        e_total = E._mk_entry("T1", 10000, date(2024, 3, 5), "REF1", "", "AR", "UNR", True, "RECEIPTS")
        e_split = E._mk_entry("T1", 4000, date(2024, 3, 5), "REF1", "ACME CORP", "AR", "UNR", True, "RECEIPTS")
        out = E._dedup_keep_largest([e_split, e_total], lambda e: e.id)
        self.assertEqual(len(out), 1)
        self.assertEqual(out[0].amount_cents, 10000)          # kept the total
        self.assertEqual(out[0].counterparty, "ACME CORP")     # borrowed


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


class TestEndToEnd(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.mkdtemp()
        self.out = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.d, ignore_errors=True)
        shutil.rmtree(self.out, ignore_errors=True)

    def _build_inputs(self):
        # BSL: three open bank lines for FHB_UTC.
        bsl = [
            ("Transaction Date", "Amount", "Line Number", "Account Servicer Reference",
             "Additional Information", "Transaction Code"),
            ("2024-03-10", "150.00", "1", "REF100200", "ACH CREDIT VENDOR", "142"),   # exact 1:1
            ("2024-03-12", "300.00", "2", "GRP500600", "DEPOSIT", "174"),             # 1:M group
            ("2024-03-15", "999.99", "3", "NADA000111", "UNKNOWN THING", "142"),      # review
        ]
        _write_xlsx(os.path.join(self.d, "20240101_FHB_UTC_BSL_UNR.xlsx"),
                    [("Exported", bsl)])
        # ST: open external transactions.
        st = [
            ("Transaction Date", "Amount", "Transaction Number", "Source", "Reference", "Counterparty"),
            ("2024-03-09", "150.00", "ST1", "EXT", "REF100200", "VENDOR INC"),        # matches line 1
            ("2024-03-11", "120.00", "ST2", "EXT", "GRP500600", "PAYER A"),           # part of group
            ("2024-03-11", "180.00", "ST3", "EXT", "GRP500600", "PAYER A"),           # part of group
            ("2024-03-01", "999.99", "ST4", "GL", "NADA000111", "JOURNAL"),           # journal: never matches
        ]
        _write_xlsx(os.path.join(self.d, "20240101_FHB_UTC_Account_ST.xlsx"),
                    [("Exported", st)])
        return "FHB_UTC"

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


if __name__ == "__main__":
    unittest.main(verbosity=2)
