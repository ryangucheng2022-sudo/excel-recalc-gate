#!/usr/bin/env python3
"""Generate the fixed sample workbooks used by CI (and local checks).

Samples (written to the directory given as argv[1], default ./samples):
  ok.xlsx       clean formulas                       -> gate verdict: success
  err.xlsx      a formula that errors on recalc      -> gate verdict: errors_found
  printarea.xlsx  openpyxl-set print area            -> gate verdict: lint_failed
                (openpyxl-emitted Print_Area definitions can make Excel
                show a repair prompt; the gate flags them by design)
  dynarray.xlsx a dynamic-array formula (SORT)       -> gate verdict: lint_failed
                (dynamic-array functions are banned by the delivery
                standard this gate enforces)
"""
import os, sys
from openpyxl import Workbook

def build(outdir):
    os.makedirs(outdir, exist_ok=True)
    wb = Workbook(); ws = wb.active; ws['A1'] = 2; ws['A2'] = 3; ws['A3'] = '=SUM(A1:A2)'
    wb.save(os.path.join(outdir, 'ok.xlsx'))
    wb = Workbook(); ws = wb.active; ws['A1'] = 1; ws['A2'] = 0; ws['A3'] = '=A1/A2'
    wb.save(os.path.join(outdir, 'err.xlsx'))
    wb = Workbook(); ws = wb.active; ws['A1'] = 5; ws['A2'] = '=A1*2'; ws.print_area = 'A1:B2'
    wb.save(os.path.join(outdir, 'printarea.xlsx'))
    wb = Workbook(); ws = wb.active; ws['A1'] = 3; ws['A2'] = 1; ws['A3'] = 2; ws['B1'] = '=SORT(A1:A3)'
    wb.save(os.path.join(outdir, 'dynarray.xlsx'))
    print('samples written to', outdir)

if __name__ == '__main__':
    build(sys.argv[1] if len(sys.argv) > 1 else 'samples')
