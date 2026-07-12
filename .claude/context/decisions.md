# Architectural Decisions

## DB Auth — Windows Auth Only
**Decision:** Always use `trusted_connection=True`. Never pass SQL username/password.
**Why:** SQL auth fails with domain error 18456 on this server. Windows Auth is the only working method in this environment.

## `pandas_utilis.py` — No Formatting Changes
**Decision:** Formatting and output changes go in the runner (`run_test_utilization_monthly.py`) or test file only. Never in `pandas_utilis.py`.
**Why:** `pandas_utilis.py` is a shared utility used across test areas. Formatting concerns belong at the call site, not in the shared layer.

## Session-Scoped DB Fixtures
**Decision:** DB stored procedure calls are wrapped in session-scoped pytest fixtures so they execute once per run and are shared across all tests.
**Why:** SP calls are expensive. Running them per-test would make the suite slow and add unnecessary DB load.

## EmpNo as String
**Decision:** EmpNo is always normalized to `str` via `str(x).split('.')[0]`.
**Why:** Excel reads numeric EmpNos as floats (e.g. `12345.0`), while DB returns them as ints or strings. The normalization strips `.0` and preserves leading zeros, ensuring consistent joins.

## Europe = Single Office Code
**Decision:** Brussels and UK both map to `Office_Code='Europe'` in the DB.
**Why:** The SP groups them under one code. Excel has them as separate sheets/rows, so the comparison filters DB rows by Excel EmpNos before merging to avoid spurious mismatches.

## `grand_total_label` Varies by Office
**Decision:** DataScience and AppliedResearch use `grand_total_label="TOTAL"` instead of `"OFFICE TOTAL"`.
**Why:** Those offices use a different Excel template with a different label for the grand total row.

## Run from `Code/` Directory
**Decision:** All test commands are run from the `Code/` working directory.
**Why:** Relative file paths in test files (fixtures, outputs, SQL files) are all anchored to `Code/`. Running from anywhere else breaks path resolution.

## SSH Remote Only
**Decision:** Use `git@github.com:cresearch-se/automation.git` (SSH), not HTTPS.
**Why:** HTTPS fails on this server without a PAT configured. SSH key is set up and works.
