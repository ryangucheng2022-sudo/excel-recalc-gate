# Excel delivery standards (companion checklist)

Distilled from production use. The gate (`bin/recalc_check.py`) enforces the
machine-checkable parts; this document covers the human/agent discipline
around it.

## The gate is not optional for formula-bearing workbooks

Any workbook that contains formulas must pass `recalc_check.py` with
`status: success` before it is delivered. `errors_found`, `lint_failed` and
`refused_external_links` all mean: fix first, or consciously accept and
document the exception — never ship silently.

Why: a workbook written by openpyxl carries formulas with **no cached
values**. Until a real engine recalculates it, previews, pandas reads and
PDF exports show blanks. The gate performs that first real calculation on a
copy and reads the results back.

## Function compatibility red lines

- Prefer the Excel 2007 generation of functions: `SUMIFS`, `INDEX`, `MATCH`,
  `IFERROR`, `SUMPRODUCT`.
- `TEXTJOIN`, `CONCAT`, `IFS`, `SWITCH`, `MAXIFS`, `MINIFS` must be written
  with the `_xlfn.` prefix (e.g. `=_xlfn.TEXTJOIN(...)`) when emitted by a
  library.
- Never emit dynamic-array functions (`XLOOKUP`, `XMATCH`, `SORT`, `FILTER`,
  `UNIQUE`, `SEQUENCE`) into a file other engines must open: LibreOffice
  cannot evaluate them and may silently fill only the anchor cell. Do
  lookups with `INDEX`/`MATCH`; sort, filter and de-duplicate in your code
  before writing cells.

## Modeling rules (models: P&L, budgets, incentive math, forecasts)

- Totals and derived values are **formulas**, never numbers pre-computed in
  Python — changing an input must ripple through the sheet.
- Parameters (tax rates, coefficients, growth rates, prices) live in
  clearly labeled assumption cells that formulas reference. No magic
  constants inside formulas (`=B5*1.05` is a defect).
- A consistent color language for models: manual inputs blue font
  (`0000FF`), formulas black, intra-workbook links green (`008000`),
  external links purple, key assumptions / to-fill cells yellow fill.
- Number formats: zero displays as `-`, negatives in parentheses,
  percentages stored as decimals and shown with one decimal, and the unit
  (yuan / thousand / ten-thousand) written in the header — never left for
  the reader to guess.
- Every hand-entered number carries its source next to it (cell comment or
  adjacent note): document and page, ledger ID, or "user-provided".

## Editing existing workbooks

- Read twice: once for formulas, once for cached values (`data_only=True`),
  and compare the two views.
- **Never save** a workbook opened with `data_only=True` — it permanently
  replaces every formula with its last cached value.
- Back up before editing. If the workbook has external links, inventory the
  linked cells first and freeze their cached values before any re-save.
- Respect the workbook's own existing conventions (input marking, formats,
  definitions); a generic standard does not override a file's local rules.

## File-generation traps

- **Do not set a print area** in generated workbooks. The `_xlnm.Print_Area`
  / `_xlnm.Print_Titles` defined names emitted by openpyxl are a known
  trigger for Excel's "content needs repair" prompt (root-caused by diffing
  Excel's repaired copy against the original: everything survived except
  the print definitions). Page setup such as `fitToPage` is fine; let the
  user set print areas inside Excel. The gate lints for this.
- If a LibreOffice conversion is part of your pipeline, make it the **last**
  step — do not open and re-save the converted file with openpyxl afterward.

## Data layer vs. presentation layer

- Keep data flat (one row per record, single header row) where formulas and
  programs read it; assemble multi-level, merged-header presentation sheets
  from the data layer by script. Never hand-edit numbers in the
  presentation layer — fix the data layer and regenerate.
- In presentation headers, color is a category code: same-type columns
  share a color, one color means one thing across all sheets, and a new
  category gets a new, recorded color — never an improvised one.
