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
- [x] D-drive existence + health check done 2026-11-12 via `check_attachment_files.py` on `Z:` (runs on the dev server; sheets explained in qcfootprints.md).
- [ ] Attachment file findings to resolve: (1) ~2.6% size gap, 27,386 files + 751 others (ask devs how the app computes 'Size (bytes)' / whether the export rewrites files); (2) the 2 PDFs that really fail (IR-59636, IR-85722) - devs compare with the app's copy (the other 76 suspicious files open fine); (3) 13 extension/content mismatches.
- [ ] Quick extra checks: existence of the 310 deleted-ticket paths; `--deep` on zip/Office files. DONE 2026-11-12 via open_check_files.py (MuPDF): of the 78 suspicious files only 2 PDFs (IR-59636, IR-85722) truly fail, 65 open with warnings (no %%EOF), plus a 200-file spot check of text-like files (0 failures); optional: bigger sample (`--sample 1734` = all).
- [ ] Stage 3: download attachments from the app (needs a fresh cookie, user runs) and compare size/hash with `Z:` - on the 78 + 13 + a random sample. Script not written.
- [ ] Stage 4: scan attachment contents for passwords (Gurpreet; needs go-ahead; zip waits for Kim; images need OCR - skipped; time 200 PDFs first, data is 24.4 GB / 32k files). Script not written.
- [ ] After re-running `compare_footprints.py`, re-run `check_attachment_files.py` (compare rewrites the report and drops its four sheets).
- [ ] Optional builds (not requested yet): Credential_Flag_Mismatch sheet (unflagged tickets with a likely password, kept out of git); compare app Description vs DB Latest Description Entry / Full Description History; HTML report; pytest wrapper.
- [ ] Housekeeping: ROTATE the password that was in a commented line of `api_utils.py` (line removed 2026-11-12, still in git history); `backup-untracked-files` branch has no PR. PR for `qc_footprints` exists (created by the user).
- **WAITING ON DEVS (as of 2026-11-13):** user explained all issues to the devs (Kim, Ofmara) and emailed them the README + report links (branch `qc_footprints`, commit 521920d). Pick up when they reply: (1) `-315` vs 'Email Request'; (2) how the app computes attachment 'Size (bytes)' / does the export rewrite files (the ~2.6% gap); (3) the 2 PDFs that fail to open (IR-59636, IR-85722) - is the app's copy complete; (4) fresh DB export, then re-run the full refresh; (5) 7 tickets the credential flag missed; (6) numeric prefix / `~` / `_` in file names; (7) 4 tickets with History Entry Count mismatch; (8) zip attachments (Kim). After answers: stage 3 (byte compare, needs fresh cookie) and/or stage 4 (password scan of attachments; Gurpreet's go-ahead; estimate ~1-2 h without OCR, unmeasured, time 200 PDFs first).
- Done 2026-11-13: `tests/QCFootprints/README.md` written (purpose, flow, how to run, every report sheet explained); keep in sync with sheets/categories/run steps.
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
