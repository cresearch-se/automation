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

## QCFootprints: Read the DB directly, never copy/paste from SSMS
**Decision:** The DB export Excel is produced by `extract_db_footprints.py` (pyodbc, Windows Auth, cell-by-cell write), not by copying SSMS results into Excel.
**Why:** Pasting broke on a title starting with a straight quote - Excel swallowed 17 tickets into one cell, spilled rows, blanked flags and flattened accents. All of those false differences vanished with the direct extract.

## QCFootprints: Long text goes into continuation columns, not truncated
**Decision:** `Full Description History` / `Latest Description Entry` are split into `<col> (part N)` columns of 32,000 chars; the compare joins them back.
**Why:** Excel cells hold 32,767 chars. Truncating hid password terms past the cut (IR-57761) and gave false 'NO PASSWORD WORD FOUND'.

## QCFootprints: Fixed-name tracked outputs, one report
**Decision:** Fixed file names, overwritten each run, tracked in git (fixtures + `output/qc_footprints_report.xlsx`); `--sample` writes an ignored `_sample.xlsx`.
**Why:** Claude works from the repo copy (apps/DB unreachable from the server) and the user wanted no timestamped clutter.

## QCFootprints: Credential values are never written to the report
**Decision:** Credential_Review shows `[REDACTED]` for values and only the words around the label.
**Why:** The report is tracked in a git repo. Context still shows usernames/emails - revisit before wider sharing.

## QCFootprints: `--add` instead of `--tickets` for extra attachment tickets
**Decision:** `extract_api_attachments.py --add IR-n` fetches the ticket and keeps everything already cached; `--tickets` builds the file from only the named tickets.
**Why:** Fetching one missing ticket with `--tickets` would have overwritten the 12,892-ticket file; the default scope would re-draw the random 500 sample and re-fetch ~500 tickets.

## QCFootprints: Credential scan ignores `PW-` project codes and URL slugs
**Decision:** `is_project_code_or_slug()` skips `pw -`/`PW-xxxx` and values preceded by `-`, `/` or `_`; `plausible_value()` accepts plain words after an explicit `label:`.
**Why:** `PW-Silver` etc. are project codes (352 false hits among unflagged tickets, 344 of them this); and 'Password: <plain word>' (IR-57400) was missed by the digit/symbol rule.

## QCFootprints: flag-check detail kept out of git
**Decision:** The tracked report's `Flag_Check_No` sheet lists ticket numbers and verdicts only; the masked context goes to the git-ignored `output/qc_footprints_flag_check.xlsx`.
**Why:** Context text includes usernames/emails around the password label; the report lives in a repo. (Credential_Review still shows context for flagged tickets - revisit.)

## QCFootprints: flag-check context file is now tracked (2026-11-11, user request)
**Decision:** Superseded 2026-11-12: there is now ONE output file. The `Flag_Check_No` sheet in `output/qc_footprints_report.xlsx` carries the masked context itself; the separate `qc_footprints_flag_check.xlsx` was removed (git rm). Earlier: that file was committed, superseding the "keep out of git" decision.
**Why:** The user wants to pull it instead of re-running the compare; Credential_Review in the tracked report already shows the same kind of masked context (usernames/emails, values `[REDACTED]`). The repo is the private cresearch-se org. Revisit if the repo's audience widens.
