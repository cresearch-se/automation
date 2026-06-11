---
name: update-utilization
description: Refresh the monthly TeamworkDB utilization fixture from Power BI and run the validation tests. Use when the user wants to do the monthly utilization run, refresh the utilization fixture, export the utilization report, or run the utilization tests for a new month.
---

# Update Utilization (monthly)

Automates the monthly TeamworkDB utilization workflow end to end: export the
Power BI paginated report, save it as the fixture, repoint the test at it, and
run the suite. Replaces the manual "download from Power BI → rename → drop in
fixtures → edit FIXTURE_FILE → run" steps.

## Prerequisites

- `config/creds/powerbi.env` must exist with valid service-principal values.
  Template: `config/powerbi.env.example`. If the file is missing or values are
  blank, tell the user to copy the template to `config/creds/powerbi.env` and
  fill in `PBI_TENANT_ID`, `PBI_CLIENT_ID`, `PBI_CLIENT_SECRET` from their dev,
  then stop — don't guess credentials.
- Run from the repo root (`Code/`) with the venv active:
  `source .venv/bin/activate`

## Steps

1. **First run / creds changed — verify access first:**
   ```bash
   python tests/TeamworkDB/refresh_utilization_fixture.py --check
   ```
   Confirms the credentials get a token and the report is reachable. If this
   warns about capacity, the workspace needs Premium/PPU/Fabric for paginated
   export — relay that to the user.

2. **Run the full refresh + tests** (default targets the *previous* month):
   ```bash
   python tests/TeamworkDB/refresh_utilization_fixture.py
   ```
   - Target a specific period: append the YYYYMM, e.g. `... 202505`.
   - Download + patch without running tests: add `--no-run`.
   - Forcing a month *older* than the report's default also needs
     `--period-param <name>` once the report's period parameter name is known.

3. **Report back** to the user:
   - which `YYYYMM` was exported and the fixture path
   - that `FIXTURE_FILE` was repointed in `test_utilization_monthly.py`
   - the test results and where the Excel/HTML report landed
     (`tests/TeamworkDB/output/`)

## Notes

- The export uses the report's **default parameters**, which match the latest
  period the dropdown shows — so the normal monthly run needs no parameter name.
- The export logic lives in `src/cornerstone_automation/utils/powerbi_utils.py`;
  the orchestrator is `tests/TeamworkDB/refresh_utilization_fixture.py`.
- If a credential is missing the script raises a clear error naming the missing
  key — surface that message to the user verbatim.
