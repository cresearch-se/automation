# QCFootprints — Footprints app vs DB validation

Status as of 2026-11-11 (report commit `9ae7591`, branch `qc_footprints`). Rewritten as a current-state summary; the
history of how we got here is in git log and in the "Lessons" section.

## Workflow (see memory: code-here-run-on-local-machine)
- The Footprints app and the fpdb001 DB are NOT reachable from the dev server. Code is written here; the user runs it on the
  local machine, pushes results to GitHub, and Claude pulls and reviews them. Claude only sees what is pushed or pasted.
- Branch `qc_footprints`, push with `git push origin qc_footprints`. If a push is rejected:
  `git pull --no-rebase --no-edit origin qc_footprints` then push again.
- Outputs have FIXED names (overwritten each run, tracked in git, no timestamped copies):
  `fixtures/QCFootprints_<YYYYMMDD>.xlsx` (app tickets), `fixtures/QCFootprints_Attachments.xlsx`,
  `fixtures/QCFootprint_DB_Details.xlsx`, `output/qc_footprints_report.xlsx`.
- Do NOT have the report open in Excel while running the compare: Windows blocks the overwrite (compare exits with
  "Cannot write ... close it in Excel"). Excel lock files `~$*` are git-ignored.

## The three inputs and the compare
Key = ticket number `IR-nnnnn` (app `Record Number` <-> DB `Ticket Number`). The DB export is used so the DB is not queried on every run.

| Script (tests/QCFootprints/) | Source | Writes |
|---|---|---|
| `extract_api_tickets.py` (was extract_footprints.py) | Footprints search API | app ticket Excel |
| `extract_api_attachments.py` (was extract_attachments.py) | Footprints attachment API | attachments Excel |
| `extract_db_footprints.py` (new) | `SELECT * FROM IR_Ticket_Summary`, Windows Auth | DB export Excel |
| `compare_footprints.py` | reads the three files above | the report |

Typical full refresh, in this order: DB extract -> API tickets -> API attachments -> compare -> push report.

### extract_api_tickets.py
- POST `https://crcofpsc.cornerstone.com/footprints/servicedesk/api/search/advanced/run?_dc=..&limit=500&page=N&start=..`
  Body = saved search (`request_body.txt`, local, gitignored). 61 pages for ~30.3k tickets. Raw pages cached in `output/raw/`; use `--refresh` after a sample run.
- Auth: NTLM does NOT work. Use `--auth cookie` (full browser Cookie header, ~584 chars; a partial one redirects to login). Cookie expires. Needs `--insecure`.
- The app total drifts upward while running (new tickets), which can shift offset paging slightly.

### extract_api_attachments.py
- GET `.../api/item/attachment/<Item ID>/<ID>?limit=25&page=1&start=0` (Item ID = 21103 for all; ID = the ticket ID column). Response is a bare list or `{"data":[...]}`; each item has name,size,extension,creationDate,uri,mimeType (signed token stripped).
- Default scope = tickets the DB says have attachments + a random 500 of those it says have none (`--sample-none`, `--seed`). Others: `--all`, `--first N`, `--probe IR-n`, `--workers`, `--timeout`, `--retries`, `--refresh`.
- `--tickets IR-a,IR-b` builds the output from ONLY those tickets (it would overwrite the full file!). To add a ticket to the existing file use `--add IR-a,...`: it fetches the listed tickets and keeps everything already cached in `output/raw_att/` (local, gitignored). Verified 2026-11-11: file went 12,892 -> 12,893 tickets.
- Tickets with 25+ attachments: server ignores paging, handled with a page-signature check, BIG_LIMIT=500 fallback and MAX_PAGES=40. Responses cached per ticket with atomic writes. Failures are listed on the Not_Fetched sheet.

### extract_db_footprints.py
- Reads straight from the DB and writes the Excel cell by cell. Replaces SSMS copy/paste, which broke the data (see Lessons).
- Config in `config/db.env` (gitignored): `FOOTPRINTS_DB_SERVER`, `FOOTPRINTS_DB_DATABASE` (or `--server/--database`). `--check` = dry run, writes nothing. Takes a few minutes (large text columns).
- Prints row count, ticket range, duplicates, missing/extra columns vs what compare expects.
- Text longer than Excel's 32,767-char cell limit is split over continuation columns `Full Description History (part 2..N)` and `Latest Description Entry (part 2..4)` (32,000 chars each); `compare_footprints.clean_db` joins them back. The largest history (IR-73240, 5.7M chars, 89 entries) needs 179 columns - likely embedded content; the file is ~70 MB.

### compare_footprints.py
- Reads app file, DB file (cached as `output/db_clean_cache.pkl`, keyed on mtime+size; delete it to force a re-read), optional attachments file. Flags: `--sample` (writes ignored `_sample.xlsx`), `--attachments PATH`, `--no-attachments`, `--db PATH`.
- DB deleted tickets (`Flagged Deleted in FootPrints` = Yes, 1,002) are excluded. DB timestamps are UTC, app is US Eastern (converted via America/New_York, verified: 0 Created On mismatches). DB export cutoff = 2026-11-02 19:15:02 Eastern; highest DB ticket IR-88530.
- Report sheets: Summary, Column_Check, Mismatch_Categories, Value_Mismatches, Missing_in_DB, Missing_in_App, Deleted_in_DB, Attachment_Count_Check, Attachments_Not_Fetched, Attachment_Name_Check, Attachment_Size_Check, Attachment_Path_Check, Credential_Review.
- Mismatch categories: CHANGED AFTER DB EXPORT (app `Updated On` later than the cutoff - expected, not a defect), USER ID INSTEAD OF NAME (DB), ACCENTS / SPECIAL CHARACTERS, NON-ASCII REPLACED BY '?', QUOTES REMOVED, DB BLANK APP HAS VALUE, OTHER. Missing_in_DB: NEWER THAN DB EXPORT vs WITHIN DB RANGE - INVESTIGATE.
- Console is set to errors="replace" (cp1252 consoles crashed on odd characters). pandas 3.x: use `sheet_name=` keyword.

## DB export columns (30, A-AD)
A-T: Ticket Number, Ticket Title, Case Name, Case / Matter Number, Status, Priority, Responsible Officer, Officer Location, Ticket Created Date, Created By, Last Updated Date, Last Updated By, Requester First Name, Requester Last Name, Requester Email, Flagged Deleted in FootPrints, Ticket Internal ID, Has Attachments, Attachment Count, Attachments Not Exported.
U Total Attachment Size Recorded (MB) (really MiB); V Attachment File Names (joined ' | ', some URL-encoded); W Attachment Locations (same order) (`D:\FPextract\out\item\attachment\6257\6518\<internal ID>\<file>`); X History Entry Count; Y/Z First/Last Entry Date (Pacific); AA Last Entry Author; AB Latest Description Entry; AC Full Description History (`[n] date (Pacific) | Name: text`); AD Possible Credential (Yes/No).
Mapping DB <-> app: Ticket Internal ID<->ID, Ticket Title<->Title, Case Name<->Client Matter Name, Case / Matter Number<->Client Matter Number, Status, Priority, Responsible Officer, Officer Location, Ticket Created Date<->Created On, Created By, Last Updated Date<->Updated On, Last Updated By<->Updated By, Requester First/Last Name<->First/Last Name, Requester Email<->Email Address.
Current DB file: 31,232 rows (30,230 live + 1,002 deleted), IR-57299..IR-88530, no duplicates. History Entry Count matches the entries in the text for all but 4 tickets (IR-65843 3 vs 106, IR-74278 14 vs 12, IR-80590 10 vs 9, IR-79430 31 vs 30) - unexplained, low priority.

## Current results (report 9ae7591)
- App 30,325 tickets; DB live 30,230 compared. Value mismatches: 612 cells / 558 tickets, ONLY two causes:
  1. USER ID INSTEAD OF NAME: 522 cells - app 'Email Request' vs DB `-315` (single id, 513 Created By + 9 Updated By).
  2. CHANGED AFTER DB EXPORT: 90 cells / 43 tickets (edited in app after the cutoff). A fresh DB export clears these.
- Missing_in_DB: 95, all NEWER THAN DB EXPORT. Missing_in_App 0. Deleted_in_DB 0.
- Attachments (12,893 tickets checked; DB says attachments 12,393): counts, names and total sizes all match except the 4 tickets changed after export (IR-86397, IR-88389, IR-88459, IR-88524 - extra files created after the cutoff, verified from creation dates). Sample of 500 'DB says none' tickets: 0 have UI attachments. Attachments_Not_Fetched: 0. Path check: 32,265 D-drive paths, 0 pattern/folder-ID problems; file name identical for 26,794, differs for 5,471 (4,889 numeric prefix, 296 illegal chars -> `_`, 241 shortened with `~`, 45 combined).
- Credential review (DB flag Possible Credential = Yes: 267 incl. deleted, 241 live): 140 MEETING PASSCODE, 67 LIKELY REAL PASSWORD - REVIEW, 34 MENTIONS PASSWORD WORD ONLY, 0 NO PASSWORD WORD FOUND.

## Credential scan logic (compare_footprints.check_credentials)
- Scans `Full Description History` of live tickets flagged Yes. Values are NEVER written (report is tracked in git): likely password -> `[REDACTED]`; context shows username/email text, so think before sharing the report wider.
- Order: LIKELY (value after password/passwd/pwd/pw/passcode label, not meeting) > MEETING PASSCODE (`?pwd=` in a URL or after 'Meeting ID') > MENTIONS WORD ONLY > NO WORD FOUND.
- `plausible_value()`: value counts if it has a digit/symbol/inner capital, OR is a plain word >=5 letters after an explicit `label:`/`label =` and not filler (CRED_STOP, CRED_FILLER). Added after IR-57400 ('Password: <plain word>') was missed.
- `is_project_code_or_slug()`: ignores `PW-xxxx` (Cornerstone project code prefix) and URL slugs like `bitcoin-passwords-2021`. Without it the unflagged scan showed 352 false hits instead of 8.
- Heuristic only: a human must review the 67. A password written in words without a label is not caught.
- DB flag misses (scan of tickets with flag = No finds a likely real password): IR-58782, IR-59410, IR-71937, IR-73289, IR-78717, IR-84144, IR-87367 (patterns 'pw is', 'Password :', 'Password -', 'Your password is:'). The flag rate (~0.7%) is flat across ticket ranges, so there is no batch gap. Not built into the report yet (offered: a Credential_Flag_Mismatch sheet, kept out of git).
- Real credentials are stored in plain text in the DB ticket history (e.g. IR-57400) - a security finding in itself.

## Lessons (root causes found - keep so they are not re-investigated)
- **SSMS copy/paste into Excel corrupts the export.** A title starting with a straight quote (IR-65403 `"Remarkable Retail” ...`) made Excel treat everything up to the next quote as ONE cell: IR-65404..65420 were swallowed into one 32,789-char Title cell (past the 32,767 cell limit, spilling into the next row). That produced the 6 'WITHIN DB RANGE' missing tickets, 2,463 'spilled text' rows, 1,610 blank Possible Credential flags (long histories), and the accent flattening / '?' / stripped-quote / blank-title mismatches. None of these were DB bugs; all disappeared with the direct DB extract. Do not raise them with Kim/Ofmara.
- Excel cell limit 32,767 chars: why history is split into continuation columns. Truncating it made the credential scan miss text past the cut (IR-57761's Zoom `?pwd=` link -> falsely 'NO PASSWORD WORD FOUND').
- The report file must be closed in Excel before running the compare.
- DB column count: the export has 30 columns; an old `usecols=range(20)` hid U-AD for a while (extract_api_attachments.db_attachment_flags still uses range(20) on purpose - it only needs A-T).
- Writing scripts through shell heredocs halves backslashes; use `chr(92)` or normalise paths to `/` before regex.
- Push rejected after the other machine pushed: pull with `--no-rebase --no-edit` then push. A pull with unstaged deletions fails: `git restore` them first.

## Open items
- Ask Kim/Ofmara: why created/updated-by is `-315` instead of 'Email Request'; get a fresh DB export (clears the 90 CHANGED AFTER cells and 95 newer tickets); the 7 tickets the DB credential flag missed; numeric prefixes / `~` shortening / `_` replacement in D-drive file names; 4 tickets with History Entry Count mismatches.
- Gurpreet's question: scan attachment contents for passwords (about 489 zip/archive attachments; Kim is still working on zip attachments).
- Human review of the 67 LIKELY REAL PASSWORD tickets.
- D-drive existence/openability check: needs a machine that sees D: (script not written).
- Not built yet: compare app Description vs DB Latest Description Entry / Full History; Credential_Flag_Mismatch sheet; HTML report; pytest wrapper.
- Repo housekeeping: plaintext password in a commented line of `src/cornerstone_automation/utils/api_utils.py` (in git history) - remove and rotate; no PR yet for `qc_footprints`; `backup-untracked-files` branch has no PR.
