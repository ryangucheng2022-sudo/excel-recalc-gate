#!/usr/bin/env python3
"""Formula delivery gate for Excel workbooks (self-implemented; mechanism
absorbed from the official xlsx skill's recalc discipline, 2026-10-28).

Checks, in order:
1. Formula lint — banned dynamic-array functions (XLOOKUP/XMATCH/SORT/
   FILTER/UNIQUE/SEQUENCE: LibreOffice cannot evaluate them and a recalc
   can silently report 0 errors), and post-2007 functions written without
   the required `_xlfn.` prefix (TEXTJOIN/CONCAT/IFS/SWITCH/MAXIFS/MINIFS).
2. External-link guard — formulas referencing another file (`[n]...`)
   lose their links and cached values if the workbook is re-saved and
   recalculated. Refuses to recalc unless --force is given.
3. Recalculation — copies the workbook to a temp dir, has LibreOffice
   recompute and rewrite it, then reads back cached values and reports
   every formula cell holding an Excel error value.

Usage: python3 recalc_check.py <file.xlsx> [--force] [--timeout 120]
Output: one JSON object. status: success | errors_found | lint_failed |
refused_external_links | recalc_failed. Exit 0 unless the check itself
could not run (recalc_failed / refused / missing file exit 2).
The source file is never modified.
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

ERROR_VALUES = {"#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "#GETTING_DATA"}
BANNED = ["XLOOKUP", "XMATCH", "SORT", "FILTER", "UNIQUE", "SEQUENCE"]
PREFIXED = ["TEXTJOIN", "CONCAT", "IFS", "SWITCH", "MAXIFS", "MINIFS"]


def lint_formulas(wb) -> tuple[int, list[dict]]:
    total = 0
    issues: list[dict] = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                total += 1
                body = v.upper()
                for fn in BANNED:
                    if re.search(rf"(?<![A-Z0-9_.]){fn}\s*\(", body):
                        issues.append({"cell": f"{ws.title}!{cell.coordinate}",
                                       "kind": "banned_function", "function": fn, "formula": v})
                for fn in PREFIXED:
                    # bare use = not preceded by _xlfn. and not part of a longer name
                    if re.search(rf"(?<!_XLFN\.)(?<![A-Z0-9_.]){fn}\s*\(", body) and fn not in BANNED:
                        # CONCAT also appears inside _xlfn.CONCAT — the lookbehind covers it;
                        # plain CONCAT( without prefix is the failing form.
                        issues.append({"cell": f"{ws.title}!{cell.coordinate}",
                                       "kind": "missing_xlfn_prefix", "function": fn, "formula": v})
                # External link: a [book]Sheet! reference, e.g. '[1]Other'!$B$2
                # or 'C:\[Book.xlsx]Sheet'!A1. Structured table refs ([[#All],[Col])
                # are excluded by the # check.
                if re.search(r"\[[^\]#\[]+\][^\n]*!", v):
                    issues.append({"cell": f"{ws.title}!{cell.coordinate}",
                                   "kind": "external_link", "formula": v})
    return total, issues


def print_area_names(path: Path) -> list[str]:
    """Defined names Excel is known to repair-prompt over when emitted by
    openpyxl (root-caused 2026-10-28: Excel's own repaired copy of a demo
    workbook had discarded exactly the _xlnm.Print_Area definition)."""
    import zipfile
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("xl/workbook.xml").decode("utf-8", "replace")
    except Exception:
        return []
    return [n for n in ("_xlnm.Print_Area", "_xlnm.Print_Titles") if n in xml]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--force", action="store_true", help="recalc even with external links")
    ap.add_argument("--timeout", type=int, default=120)
    args = ap.parse_args()

    src = Path(args.file)
    if not src.exists():
        print(json.dumps({"status": "recalc_failed", "error": f"file not found: {src}"}))
        return 2

    wb = load_workbook(src)
    pa = print_area_names(src)
    if pa:
        print(json.dumps({"status": "lint_failed", "issues": [
            {"kind": "print_area_defined_name", "names": pa,
             "note": "openpyxl-emitted Print_Area/Print_Titles definitions can make "
                     "Excel show a repair prompt; remove the print area (keep page "
                     "setup fitToPage) or set it inside Excel instead"}]},
            ensure_ascii=False))
        return 0
    total_formulas, issues = lint_formulas(wb)
    lint_issues = [i for i in issues if i["kind"] != "external_link"]
    ext_links = [i for i in issues if i["kind"] == "external_link"]
    if lint_issues:
        print(json.dumps({"status": "lint_failed", "total_formulas": total_formulas,
                          "issues": lint_issues[:100]}, ensure_ascii=False))
        return 0
    if ext_links and not args.force:
        print(json.dumps({"status": "refused_external_links", "total_formulas": total_formulas,
                          "external_link_cells": [i["cell"] for i in ext_links[:100]],
                          "note": "re-saving would strip external links and cached values; "
                                  "copy values out first, or pass --force to accept the loss"},
                         ensure_ascii=False))
        return 2
    if total_formulas == 0:
        print(json.dumps({"status": "success", "total_formulas": 0, "total_errors": 0,
                          "note": "no formulas in workbook"}))
        return 0

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        work = td / src.name
        shutil.copy2(src, work)
        r = subprocess.run(
            ["soffice", "--headless", "--norestore", "--convert-to", "xlsx",
             "--outdir", str(td / "out"), str(work)],
            capture_output=True, text=True, timeout=args.timeout)
        out_file = td / "out" / src.name
        if r.returncode != 0 or not out_file.exists():
            print(json.dumps({"status": "recalc_failed",
                              "error": (r.stderr or r.stdout or "soffice failed")[-500:]}))
            return 2
        wb2 = load_workbook(out_file, data_only=True)
        errors: dict[str, list[str]] = {}
        checked = 0
        for ws in wb.worksheets:
            ws2 = wb2[ws.title] if ws.title in wb2.sheetnames else None
            if ws2 is None:
                continue
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if isinstance(v, str) and v.startswith("="):
                        checked += 1
                        got = ws2[cell.coordinate].value
                        if isinstance(got, str) and got in ERROR_VALUES:
                            errors.setdefault(got, []).append(f"{ws.title}!{cell.coordinate}")
        total_errors = sum(len(v) for v in errors.values())
        print(json.dumps({
            "status": "errors_found" if total_errors else "success",
            "total_formulas": total_formulas, "formulas_checked": checked,
            "total_errors": total_errors,
            "error_summary": {k: v[:100] for k, v in errors.items()},
            "external_links_forced": bool(ext_links) or False,
        }, ensure_ascii=False))
        return 0


if __name__ == "__main__":
    sys.exit(main())
