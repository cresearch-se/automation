# TODO / Work In Progress

## Open Issues

### YTD False MISSING IN DB — Under Investigation
- **What:** Employee YTD comparison tests fire duplicate "MISSING IN DB" + "MISSING IN XLS" errors for employees who exist in the DB but have NULL `Target_Hours`.
- **Status:** Monthly tests have the fix (`df_db = df_db[df_db[NUMERIC_COLS[0]].notna()].copy()`). YTD path still under investigation.
- **Debug aid:** There is a commented-out print statement in `run_employee_comparison` that dumps the full merged table — uncomment to see what's happening in the YTD merge.
- **Next step:** Apply the same `notna()` filter to the YTD branch and verify it eliminates the false positives without dropping legitimate mismatches.

---

## Recurring Monthly Task

### Utilization Monthly Run
Each month, update the fixture file and run the full report:
1. Drop the new `Utilization_YYYYMM.xlsx` into `tests/TeamworkDB/fixtures/`
2. Update `FIXTURE_FILE` in `test_utilization_monthly.py` to point to the new file
3. Run: `python tests/TeamworkDB/run_test_utilization_monthly.py`
4. Review Excel + HTML output in `tests/TeamworkDB/output/`

> See [teamworkdb.md](teamworkdb.md) for full run commands and test group details.

---

## Backlog / Nice to Have
- Nothing tracked yet — add items here as they come up.
