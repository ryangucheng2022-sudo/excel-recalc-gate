# CI kit

This folder holds a ready-to-enable GitHub Actions setup, kept outside
`.github/workflows/` so it can be reviewed (and versioned) like any other
file in the repo.

**Enable it:** copy `ci.yml` to `.github/workflows/ci.yml` and push.

Two jobs:
- `lint-and-unit` — unit tests on a plain runner. The cases that need a real
  LibreOffice recalculation skip themselves when `soffice` is absent.
- `real-recalc` — installs LibreOffice, runs the full suite including the
  real recalculation path, then pushes the fixed samples from
  `make_samples.py` (clean / error formula / print area / dynamic array)
  through the gate CLI and asserts each verdict: clean -> success,
  error formula -> errors_found, openpyxl print area -> lint_failed,
  banned dynamic-array function -> lint_failed.

`make_samples.py` can also be run locally: `python ci/make_samples.py samples`.
