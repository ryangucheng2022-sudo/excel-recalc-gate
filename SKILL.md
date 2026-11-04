# Skill: excel-recalc-gate

Verify any formula-bearing Excel workbook before delivering it.

## When to use

Before delivering, attaching or publishing any `.xlsx` / `.xlsm` workbook
that contains formulas and was generated or edited programmatically
(openpyxl, xlsxwriter, pandas, an agent, …).

## Procedure

1. Run: `python3 bin/recalc_check.py <file.xlsx>`
2. Read the JSON verdict:
   - `success` → deliver.
   - `errors_found` → fix the named cells (`error_summary` lists them) and
     re-run; do not deliver.
   - `lint_failed` → replace banned functions (`INDEX`/`MATCH` instead of
     `XLOOKUP`; sort/filter/dedupe in code), add the `_xlfn.` prefix where
     reported, or remove the print-area defined name; then re-run.
   - `refused_external_links` → freeze the external-link cells' values
     first; only re-run with `--force` when losing the links is explicitly
     acceptable.
   - `recalc_failed` → the check did not run (missing file / LibreOffice);
     resolve that before delivering anything.
3. Never save a workbook that was opened with `data_only=True`.

## Notes

- The tool never modifies the source file; recalculation runs on a copy.
- Full rules: `docs/delivery-standards.md`.
