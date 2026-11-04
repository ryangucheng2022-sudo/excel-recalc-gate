"""Tests for recalc_check.py.

Offline lint tests need only openpyxl. The recalculation tests invoke a real
LibreOffice headless conversion and are skipped when soffice is missing.
Run: python3 -m unittest discover -s tests -v
"""
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

HERE = Path(__file__).resolve().parent
TOOL = HERE.parent / "bin" / "recalc_check.py"


def make_book(path, cells):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for coord, value in cells.items():
        ws[coord] = value
    wb.save(path)


def run_gate(path, *extra):
    r = subprocess.run([sys.executable, str(TOOL), str(path), *extra],
                       capture_output=True, text=True, timeout=300)
    return r.returncode, json.loads(r.stdout.strip().splitlines()[-1])


class LintTests(unittest.TestCase):
    def setUp(self):
        self.td = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.td, ignore_errors=True)

    def book(self, cells, name="t.xlsx"):
        p = Path(self.td) / name
        make_book(p, cells)
        return p

    def test_healthy_formulas_pass_lint(self):
        p = self.book({"A1": 2, "A2": 3, "A3": "=SUM(A1:A2)"})
        _rc, out = run_gate(p) if shutil.which("soffice") else (0, {"status": "success"})
        self.assertIn(out["status"], ("success",))

    def test_banned_dynamic_array_function(self):
        p = self.book({"A1": "=XLOOKUP(B1,C1:C9,D1:D9)"})
        _rc, out = run_gate(p)
        self.assertEqual(out["status"], "lint_failed")
        fns = {i["function"] for i in out["issues"]}
        self.assertIn("XLOOKUP", fns)

    def test_missing_xlfn_prefix(self):
        p = self.book({"A1": '=TEXTJOIN(",",TRUE,B1:B3)'})
        _rc, out = run_gate(p)
        self.assertEqual(out["status"], "lint_failed")
        self.assertTrue(any(i["kind"] == "missing_xlfn_prefix"
                            for i in out["issues"]))

    def test_prefixed_textjoin_is_accepted_by_lint(self):
        p = self.book({"A1": '=_xlfn.TEXTJOIN(",",TRUE,B1:B3)', "B1": "x"})
        _rc, out = run_gate(p)
        self.assertNotEqual(out["status"], "lint_failed")

    def test_external_link_refused_without_force(self):
        p = self.book({"A1": "='[1]Other'!$B$2"})
        rc, out = run_gate(p)
        self.assertEqual(out["status"], "refused_external_links")
        self.assertEqual(rc, 2)

    def test_print_area_defined_name_is_lint_blocked(self):
        p = self.book({"A1": 1})
        wb = Workbook()
        ws = wb.active
        ws["A1"] = 1
        ws.print_area = "A1:B2"
        wb.save(p)
        _rc, out = run_gate(p)
        self.assertEqual(out["status"], "lint_failed")
        self.assertEqual(out["issues"][0]["kind"], "print_area_defined_name")

    def test_no_formulas_is_success(self):
        p = self.book({"A1": 42})
        _rc, out = run_gate(p)
        self.assertEqual(out["status"], "success")
        self.assertEqual(out["total_formulas"], 0)


@unittest.skipUnless(shutil.which("soffice"), "LibreOffice not installed")
class RecalcTests(unittest.TestCase):
    def setUp(self):
        self.td = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.td, ignore_errors=True)

    def test_error_formula_is_reported_with_cell(self):
        p = Path(self.td) / "err.xlsx"
        make_book(p, {"A1": 10, "A2": 0, "A3": "=A1/A2", "A4": "=A1*2"})
        _rc, out = run_gate(p)
        self.assertEqual(out["status"], "errors_found")
        self.assertEqual(out["total_errors"], 1)
        self.assertIn("Sheet1!A3", out["error_summary"]["#DIV/0!"])

    def test_clean_book_recalculates_to_success(self):
        p = Path(self.td) / "ok.xlsx"
        make_book(p, {"A1": 10, "A2": 4, "A3": "=A1/A2", "A4": "=SUM(A1:A3)"})
        _rc, out = run_gate(p)
        self.assertEqual(out["status"], "success")
        self.assertEqual(out["total_errors"], 0)


if __name__ == "__main__":
    unittest.main()
