# TODO / Work In Progress

## Open Issues

### YTD False MISSING IN DB — Under Investigation
- **What:** Employee YTD comparison tests fire duplicate "MISSING IN DB" + "MISSING IN XLS" errors for employees who exist in the DB but have NULL `Target_Hours`.
- **Status:** Monthly tests have the fix (`df_db = df_db[df_db[NUMERIC_COLS[0]].notna()].copy()`). YTD path still under investigation.
- **Debug aid:** There is a commented-out print statement in `run_employee_comparison` that dumps the full merged table — uncomment to see what's happening in the YTD merge.
- **Next step:** Apply the same `notna()` filter to the YTD branch and verify it eliminates the false positives without dropping legitimate mismatches.

### QCFootprints — open items (details in [qcfootprints.md](qcfootprints.md))
- [ ] Ask Kim/Ofmara: `-315` instead of 'Email Request' in Created/Updated By (522 cells); a fresh DB export; 7 tickets the DB `Possible Credential` flag missed (IR-58782, 59410, 71937, 73289, 78717, 84144, 87367); numeric prefix / `~` / `_` in D-drive file names; 4 tickets whose History Entry Count != entries in text. (Accent/quote/`?` issues were our copy-paste, NOT DB bugs - do not raise.)
- [ ] After the fresh DB export: re-run extract_db_footprints.py -> compare -> push; expect the 90 'changed after export' cells and 95 newer tickets to clear.
- [ ] Human review of the 67 LIKELY REAL PASSWORD tickets (Credential_Review sheet); decide whether the report (tracked in git, shows usernames/emails in masked context) should stay tracked.
- [ ] Gurpreet's question: scan attachment contents (about 489 zip/archive files) for passwords; Kim still working on zip attachments.
- [x] D-drive existence + health check done 2026-11-12 via check_attachment_files.py on Z:. Open: explain the ~2.6% size gap (ask devs), review 78 corrupt + 13 mismatch files, stage 3 (byte compare with the app download) and stage 4 (scan attachment contents).
- [ ] Optional builds (not requested yet): Credential_Flag_Mismatch sheet (unflagged tickets with a likely password, kept out of git); compare app Description vs DB Latest Description Entry / Full Description History; HTML report; pytest wrapper.
- [ ] Housekeeping: remove the plaintext password in a commented line of `src/cornerstone_automation/utils/api_utils.py` (also in git history) and rotate it; open a PR for `qc_footprints`; `backup-untracked-files` branch has no PR.
- Done 2026-11-11: root cause of the 6 'missing in DB' tickets (SSMS paste), direct DB extract script, split-history columns, attachments file complete (12,893), credential scan fixes (PW- codes, plain-word values), Excel lock-file ignore.

**Recurring QCFootprints refresh (in order):** `extract_db_footprints.py` -> `extract_api_tickets.py` (fresh cookie) -> `extract_api_attachments.py` (or `--add IR-n` for single tickets) -> `compare_footprints.py` (report closed in Excel) -> commit + push `fixtures/*` and `output/qc_footprints_report.xlsx`.

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
