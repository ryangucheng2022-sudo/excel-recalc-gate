# Changelog

## v1.0.0 (2026-11-04)

Initial open-source release.

- **Formula lint**: banned dynamic-array functions (XLOOKUP/XMATCH/SORT/
  FILTER/UNIQUE/SEQUENCE), missing `_xlfn.` prefixes on post-2007
  functions, and `Print_Area`/`Print_Titles` defined names that trigger
  Excel repair prompts in generated files.
- **External-link guard**: refuses to recalculate workbooks with
  external-link formulas unless `--force` is passed, because re-saving
  would strip the links and their cached values.
- **Real recalculation**: headless LibreOffice recompute on a temp copy;
  every formula cell is read back and error values are reported by sheet
  and coordinate. The source file is never modified.
- **JSON verdict** with five statuses (`success`, `errors_found`,
  `lint_failed`, `refused_external_links`, `recalc_failed`) and CI-friendly
  exit codes.
- Companion delivery-standards checklist in `docs/`.
- 9 tests (lint/verdict cases offline; recalculation cases use a real
  LibreOffice when available).
