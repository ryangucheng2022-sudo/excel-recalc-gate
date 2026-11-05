# excel-recalc-gate

English | [简体中文](README.zh-CN.md)

**A delivery gate for Excel workbooks that contain formulas.** Before you ship a generated `.xlsx`, this tool lints the formulas, recalculates the workbook with headless LibreOffice on a throwaway copy, and reports every formula cell that evaluates to an Excel error — as one JSON verdict. Ship only on `success`.

## The problem it solves

Libraries like openpyxl write formulas **without cached values**. Until a real spreadsheet engine opens and recalculates the file, every formula cell is empty when read by pandas, preview panes, PDF converters — and sometimes by your recipient. Worse, some formulas fail *silently*:

- Dynamic-array functions (`XLOOKUP`, `SORT`, `FILTER`, …) can't be evaluated by LibreOffice at all; a recalc may report zero errors while filling only the anchor cell.
- Post-2007 functions (`TEXTJOIN`, `IFS`, `MAXIFS`, …) break unless written with the `_xlfn.` prefix.
- Formulas with external links (`'[1]Other'!$B$2`) lose their links *and* cached values the moment the file is re-saved.
- A `Print_Area` defined name emitted by openpyxl can make Excel show a "content needs repair" prompt — root-caused by diffing Excel's own repaired copy of a workbook.

If you generate workbooks with scripts or AI agents, these failures reach your reader before they reach you. This gate catches them first.

## Who is it for

| Role | Typical scenarios | What you get |
|---|---|---|
| **Finance & FP&A** | Models, budgets, forecasts and board packs built or assembled by scripts | No workbook leaves your desk with uncomputed or error-valued formulas; the gate verdict is a JSON artifact you can attach to your QA trail |
| **Data analysts & engineers** | Scheduled pipelines that emit `.xlsx` reports for business users | A CI-friendly exit code and machine-readable report: failing cells are named (`Sheet!A3`, `#DIV/0!`), not buried in a GUI |
| **AI agent builders** | Agents that generate spreadsheets with openpyxl/xlsxwriter | A deterministic pre-delivery check an agent can run and self-repair against — lint rules encode the exact function-compatibility traps agents fall into |
| **Anyone shipping Excel** | One-off exports you send to people who open them in Excel, WPS or a phone preview | The file your reader opens shows numbers, not blanks — because a real engine has already computed it once |

## How the gate decides

Spreadsheet applications cache formula results inside the file. A workbook can therefore display values that no longer match its formulas — after edits by another tool, a broken external link, or a partial save — and look perfectly fine on screen.

The gate runs three steps on a copy of the workbook:

1. **Lint.** Structural checks first. If the workbook carries external links, re-saving would strip the links and their cached values, so the file is reported instead of rewritten.
2. **Recompute.** A headless LibreOffice recalculation forces every formula to re-evaluate.
3. **Read back.** The cached values are read from the recomputed file and scanned for error values (#VALUE!, #REF!, …).

Verdicts: `success` (safe to deliver), `errors_found` (formula errors present after recalculation), `lint_failed` (the file should not be re-saved at all). A delivery workflow should treat anything except `success` as a stop, not a warning.

## Usage

```bash
python3 bin/recalc_check.py report.xlsx
python3 bin/recalc_check.py report.xlsx --force     # recalc despite external links (accepts losing them)
python3 bin/recalc_check.py report.xlsx --timeout 120
```

Output is a single JSON object:

```json
{"status": "success", "total_formulas": 128, "formulas_checked": 128, "total_errors": 0}
```

| status | meaning | exit code |
|---|---|---|
| `success` | lint clean and zero error values after recalculation | 0 |
| `errors_found` | recalculation produced error values; `error_summary` names the cells | 0 |
| `lint_failed` | banned function, missing `_xlfn.` prefix, or a `Print_Area`/`Print_Titles` defined name | 0 |
| `refused_external_links` | external-link formulas present and `--force` not given | 2 |
| `recalc_failed` | the check itself could not run (missing file, LibreOffice failure) | 2 |

The source file is **never modified** — recalculation happens on a temp copy.

## Companion checklist

[docs/delivery-standards.md](docs/delivery-standards.md) distills the modeling and editing rules that pair with the gate: which functions to use, how to structure assumptions, why you must never save a workbook opened with `data_only=True`, and why generated workbooks should not carry a print area.

## Requirements

- Python 3.9+ with `openpyxl`
- LibreOffice (`soffice`) on PATH for the recalculation step (lint-only checks work without it)

## Tests

```bash
python3 -m unittest discover -s tests -v
# 9 cases: 7 lint/verdict cases + 2 real-recalc cases (skipped if soffice is absent)
```

## Version

v1.0.0 — see [CHANGELOG.md](CHANGELOG.md).

## License

MIT — see [LICENSE](LICENSE).
