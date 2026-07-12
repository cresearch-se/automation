# TeamworkDB — Utilization Tests

## Overview
Monthly QA validation that compares an Excel fixture (exported from Power BI / business team) against SQL Server stored procedure results.

## Key Files
- `tests/TeamworkDB/test_utilization_monthly.py` — all test logic + normalizers
- `tests/TeamworkDB/run_test_utilization_monthly.py` — runner: calls pytest, parses output, generates Excel + HTML report
- `tests/TeamworkDB/fixtures/Utilization_YYYYMM.xlsx` — monthly fixture (only this changes each month)
- `tests/TeamworkDB/output/` — generated reports land here

## DB Details
- **Server:** `SQLT1COFIN` (env var: `T1_DB_SERVER`)
- **Database:** `ReportDevl`
- **Stored Procedure:** `SP_Utilization_Validation`
- **Params:** `StartDate`, `EndDate`, `Type` (`'Summary'` or `'Detail'`)
- **Auth:** Windows Auth only — `trusted_connection=True` (SQL auth fails with domain error 18456)

## Session-Scoped Fixtures (fetched once, shared across tests)
| Fixture | SP Type | Date Range |
|---|---|---|
| `db_summary_monthly` | Summary | Monthly |
| `db_summary_ytd` | Summary | YTD |
| `db_detail_monthly` | Detail | Monthly |
| `db_detail_ytd` | Detail | YTD |

## Test Groups
| Group | What It Checks |
|---|---|
| `test_format_validations_by_location` | Offices present, titles present + correct order, no blanks/zeros |
| `test_totals_validation_by_location` | Subtotals and grand totals match computed sums |
| `test_format_validations_by_employee` | No blanks, unique EmpNo, subtotals per title |
| `test_db_comparison_us_monthly/europe_monthly/us_ytd/europe_ytd` | Location-level DB comparison |
| `test_db_comparison_employee_monthly` / `_ytd` | Employee-level DB comparison (parametrized × 9 offices) |

## Key Constants
```python
NUMERIC_COLS = ['Target_Hours', 'Target_Rev', 'Actual_Hours', 'Standard_Rev']
JOIN_KEYS_BY_EMPLOYEE = ["EmpNo"]
```

## Office / Region Rules
- Europe = Brussels + UK — both map to `Office_Code='Europe'` in DB
- Europe DB rows are filtered by Excel EmpNos before merge
- DataScience and AppliedResearch use `grand_total_label="TOTAL"` (not `"OFFICE TOTAL"`)

## EmpNo Handling
Canonical type is `str` — normalized via `str(x).split('.')[0]` to strip `.0` from floats and preserve leading zeros.

## Monthly Update — Only One Line Changes
```python
FIXTURE_FILE = "tests/TeamworkDB/fixtures/Utilization_YYYYMM.xlsx"
```
Sheet names and date ranges derive automatically from the filename.

## Run Commands (from `Code/` with venv active)
```bash
# Full report
python tests/TeamworkDB/run_test_utilization_monthly.py

# Format + totals only
python -m pytest tests/TeamworkDB/test_utilization_monthly.py -k "format_validations or totals_validation" -v --no-cov -s --tb=short

# Employee monthly comparison
python -m pytest tests/TeamworkDB/test_utilization_monthly.py -k "test_db_comparison_employee_monthly" -v --no-cov -s --tb=short

# Employee YTD comparison
python -m pytest tests/TeamworkDB/test_utilization_monthly.py -k "test_db_comparison_employee_ytd" -v --no-cov -s --tb=short
```
> Note: add `--override-ini=addopts=` when running pytest directly (pyproject.toml injects `--cov` via addopts).

## Known Issues

### False "MISSING IN DB" + "MISSING IN XLS" duplicates
**Root cause:** Employees exist in DB with NULL `Target_Hours` — the `pandas_utilis.py` logic uses `numeric_cols[0]` (i.e., `Target_Hours`) as the "employee present in DB" indicator. A NULL value triggers MISSING IN DB even when the employee IS in the DB.

**Fix applied (monthly tests):**
```python
# In run_employee_comparison, filter df_db before merge:
df_db = df_db[df_db[NUMERIC_COLS[0]].notna()].copy()
```

**Status for YTD tests:** Still under investigation. A debug print (currently commented out) prints the full merged table — uncomment to diagnose.
