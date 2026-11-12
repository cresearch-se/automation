# QCFootprints — Footprints app vs DB validation

Status as of 2026-11-12 (branch `qc_footprints`; attachment FILE check added, see the last section). Rewritten as a current-state summary; the
history of how we got here is in git log and in the "Lessons" section.

## Workflow (see memory: code-here-run-on-local-machine)
- The Footprints app and the fpdb001 DB are NOT reachable from the dev server. Code is written here; the user runs it on the
  local machine, pushes results to GitHub, and Claude pulls and reviews them. Claude only sees what is pushed or pasted.
- Exception: `check_attachment_files.py` runs HERE because drive `Z:` (the export server's D drive) is mapped on this server. This server has no Excel viewer, so the user pulls and reviews the report locally.
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
- D-drive existence/health check: DONE 2026-11-12 (see the last section); open items are listed there.
- Not built yet: compare app Description vs DB Latest Description Entry / Full History; Credential_Flag_Mismatch sheet; HTML report; pytest wrapper.
- Repo housekeeping: plaintext password in a commented line of `src/cornerstone_automation/utils/api_utils.py` (in git history) - remove and rotate; no PR yet for `qc_footprints`; `backup-untracked-files` branch has no PR.

## Flag check: is Possible Credential = No/blank really right? (added 2026-11-11)
- `check_flag_no()` in compare_footprints.py scans the 29,989 live tickets with flag No/blank using the same `assess_text()` as the Yes list. Report gets a Summary block, a `Flag_Check_No` sheet that includes the masked context (usernames/emails, values [REDACTED]). Single output file: `output/qc_footprints_report.xlsx` (the separate flag_check file was removed 2026-11-12).
- Result on the current data: 29,516 NO PASSWORD WORD FOUND, 466 MENTIONS PASSWORD WORD ONLY (mostly 'reset your password' style), **7 LIKELY REAL PASSWORD** (the flag misses listed in the open items), 0 MEETING PASSCODE - the DB flags every Zoom `?pwd=` ticket, so its flag is not a plain keyword search (it flags 140 meeting-passcode tickets but misses 'pw is' / 'Password :' / 'Password -' wording and flags only 34 word-only tickets vs 466 unflagged ones).
- CRED_WORD_RX now ignores `PW-xxxx` project codes (was inflating unflagged word-only from 466 to 1,075).

## Attachment file validation on the drive (added 2026-11-12) - `check_attachment_files.py`
### How it runs
- Runs HERE on the dev server (unlike the other scripts): drive `Z:` = `\\25sqlt4cotest\d$` is the export server's D drive. DB paths `D:\FPextract\...` are mapped to `Z:\FPextract\...`. (`D:` on this server is an unrelated empty drive - do not use it.) Options: `--drive`, `--limit N` (trial run, not written), `--no-health`, `--deep` (reads every member of zip/Office files; not run yet), `--threads`.
- Inputs are fixtures only (UI attachments file + DB export, no app/DB access). Full run takes about 3 minutes (stage 1 ~30 s, stage 2 ~100 s).
- It ADDS four sheets to `output/qc_footprints_report.xlsx` (the single report file). Run `compare_footprints.py` FIRST: it rewrites the report and drops these sheets. Close the report in Excel first.
- Scope: the 32,266 paths on the 12,393 LIVE tickets that have attachments (12,893 tickets incl. the 500 'DB says none' sample). NOT covered: 310 paths on 164 tickets flagged deleted in the DB (no UI data; an existence-only check could be added).

### What was and was NOT verified
- Paths: DB path -> `Z:` file exists. (Path pattern / ticket-folder checks are in Attachment_Path_Check from compare.) The app has no D: path, so paths cannot be compared with the UI.
- Sizes: each file's size on disk vs the UI `Size (bytes)`, paired by attachment name (sorted by size when a name repeats); ticket total on disk vs DB `Total Attachment Size Recorded (MB)` (MiB; the DB has no per-file size).
- Health = reads CONTENT, not file names (extension only decides what to expect). For every non-empty file: first 16 bytes + last 2 KB. Checks: header matches type (pdf, zip family, OLE doc/xls/ppt/msg incl. old BIFF xls, jpg, png, gif, bmp, tif, rtf); end markers (pdf `%%EOF`, jpg FFD9, png IEND, gif `;`); zip/docx/xlsx/pptx opened with Python zipfile (table of contents only, Office files must contain `[Content_Types].xml`). Text-like types (eml, txt, csv, htm...) have no signature -> 'not checked'.
- NOT done: opening files in Word/Excel/Acrobat, reading the middle of files, extracting zip members (`--deep`), byte compare with the app's download. So the health categories are hints, not proof (the open check gives the verdict); every file could be opened and read (CANNOT READ FILE = 0) and no zip/Office file failed to open.

### Sheets added by this script (what each is for)
| Sheet | Use |
|---|---|
| `File_Check_Summary` | Counts for everything below; start here. |
| `Attachment_File_Check` | Existence/size problems. One row per file; a row with a BLANK attachment name is a ticket-level total row (Note: `DB x MB, disk y MiB`). Columns: Record Number, Attachment name, UI size (bytes), Disk size (bytes), DB path, Mapped path, Note, Category. |
| `Attachment_File_Health` | Files that look damaged (only problems are listed). Columns incl. Extension, Detail, Mapped path. |
| `Attachment_Orphan_Files` | Files in a ticket's folder that no DB path points to (empty so far). |

### Categories
- `Attachment_File_Check`: FILE MISSING ON DRIVE (DB path has no file) | ZERO-BYTE FILE | DISK ~2.6% SMALLER THAN UI (disk/UI between 0.973 and 0.976) | SIZE DIFFERS (UI vs DISK) (any other difference) | TICKET TOTAL SIZE != DB (ticket-level; same gap rolled up) | IN UI BUT NOT IN DB (app attachment without DB path) | CREATED AFTER DB EXPORT (same, but file newer than the export cutoff - expected) | CHANGED AFTER DB EXPORT (ticket total differs because of such a file - expected) | UNEXPECTED DB PATH | PATH COUNT != NAME COUNT (malformed DB values).
- `Attachment_File_Health`: NO END MARKER (CHECK - see File_Open_Check) (pdf without %%EOF, jpg/png/gif without end marker - only a hint, 75 files) | CORRUPT (BAD HEADER OR ZIP) (header not recognised or zip will not open, 3 files) | EXTENSION / CONTENT MISMATCH (e.g. `.png` that is a tif, 13) | CANNOT READ FILE (0). Renamed 2026-11-12: it used to be one category "CORRUPT OR TRUNCATED" (78), which overstated the problem.

### Results (first full run, 2026-11-12)
- 32,266 DB paths: ALL present, 0 zero-byte, 0 orphans, 0 bad paths.
- SIZES (the main open issue): only 4,129 files equal the UI size; 27,386 are about 2.6% smaller on disk (disk/UI ~0.974 = 38/39, i.e. UI = disk + 2 bytes per 76, like MIME line wrapping - only a hypothesis); 751 more are 3-8% smaller; 5,025 tickets' disk total != DB total (e.g. IR-57299: UI/DB 2.88 MB, disk 2.81 MiB). The DB total equals the sum of UI sizes, so the DISK is the odd one out. Cause unknown: either the app's size field overstates the size or the export rewrote the files. About 20,000 PDFs read fine, so it does not look like systematic truncation.
- HEALTH: 30,441 OK, 1,734 not checked (text-like), 13 extension/content mismatches (png that is tif/jpg, docx/xlsx that are old OLE, rtf that is OLE; probably mislabelled), 78 corrupt/truncated (68 pdf without `%%EOF`, 8 tiny gifs without trailer, 2 old xls with unrecognised header; IR-59636 is exactly 96 KiB = likely cut). A first run flagged 10 more xls that were only old BIFF headers; fixed in SIGS.
- Expected, not defects: 9 files / 1 ticket total created or changed after the DB export (a fresh export clears them).

### What still needs attention (priority order)
1. The size gap - decide via stage 3 and the devs' answer.
2. The 78 suspicious files - open in the app and on the drive (devs).
3. The 13 mismatched extensions.
4. Small: 9+1 after-export items; 310 deleted-ticket paths not checked.

### Planned next validations
- Quick, no one else needed: existence check of the 310 deleted-ticket paths; `--deep` on zip/Office files; readability check of the 1,734 text-like files; full read of the 78 files (is `%%EOF` really absent?).
- Stage 3 (needs a fresh browser cookie, run by the user): download attachments from the app and compare size and hash with the file on Z:. Run on the 78 suspicious + 13 mismatched + a random ~50-500 sample. Settles the size gap: downloaded size == disk size -> the app's size field is inflated; == UI size -> exported files differ.
- Stage 4 (Gurpreet's question, needs go-ahead): extract text from pdf/docx/xlsx/txt/eml and run `assess_text()`; masked context only. Zip waits for Kim (476 zip files, 2.3 GB); images (~7,000 png/jpg/gif) hold no text and would need OCR - skip and list as not scanned; scanned PDFs show 'no text found'. Data volume: 32,275 files, 24.4 GB, of which 13,658 PDFs = 18.4 GB. Estimate (not measured): roughly 30 min - 2 h with 16 threads over the share; time 200 PDFs first.
- Dev questions: how Footprints computes attachment 'Size (bytes)'; does the export decode/rewrite attachments; open the 78 files in the app; renaming rules (numeric prefix is NOT the ticket number or folder ID - looks like an attachment-level ID added when names clash; 5,472 of 32,266 paths carry it).

## Really opening files: `open_check_files.py` (added 2026-11-12)
- Opens files with real parsers on the server (no Word/Acrobat here): pypdf (every page + text extraction), Pillow (full pixel decode), xlrd (old .xls), Python email parser (.eml), decode/parse for txt/csv/htm/log. Libraries are in requirements.txt (pypdf, pillow, xlrd). Read-only.
- Targets: the 78 `NO END MARKER` + `CORRUPT (BAD HEADER OR ZIP)` files from `Attachment_File_Health`, optionally the 13 mismatches (`--include-mismatch`), plus a random spot check of the text-like files (`--sample N`, default 200 of 1,734, `--seed`). Run order: compare -> check_attachment_files -> open_check_files (compare rewrites the report and drops these sheets).
- Adds sheets `File_Open_Summary` (counts per source and extension) and `File_Open_Check` (one row per file: Result, Detail, earlier health finding, path; failures first). Results: OPENS OK / OPENS WITH WARNINGS (parser repaired something) / FAILS TO OPEN / NO PARSER (rtf, docx/xlsx-that-are-OLE, gz, rar, mp3, ics, vcf, mhtml - not opened).
- CORRECTION 2026-11-12 (user opened some of the 78 manually and they opened fine): the first run used pypdf only, which is strict and said 'Stream has ended unexpectedly' for 64 PDFs that viewers open without trouble. `open_pdf` now uses PyMuPDF (tolerant like Acrobat/Edge, rebuilds a damaged index) as the verdict and keeps pypdf only as fallback; requirements.txt has pymupdf. Do NOT claim '64 truncated PDFs' - that was wrong.
- Results 2026-11-12 (MuPDF): SUSPICIOUS 78 files -> 9 OPENS OK, 65 OPENS WITH WARNINGS, 4 FAILS TO OPEN. PDFs (68): 65 open and every page reads; they only lack the `%%EOF` end marker (a few also needed an index repair, e.g. IR-70963 with junk before %PDF) - cosmetic/structural, content readable. Only 2 PDFs really fail: IR-59636 (98,304 bytes) and IR-85722 (1,441,792 bytes) - MuPDF finds 0 pages; both sizes are exact multiples of 64 KiB (1.5x and 22x), so they look cut at a chunk boundary = truly truncated. All 8 tiny GIFs decode fine (false alarm: no trailer byte but valid). The 2 xls are not corrupt, they are XML/HTML saved as .xls (IR-73830, IR-84212). MISMATCH 13: png/jpg all decode (mislabelled, harmless), 6 not opened (no parser). TEXT SPOT CHECK 200: 191 open OK (156 eml, 16 csv, 16 txt, ...), 0 failures, 9 no parser (gz, rar, mp3, ics, vcf, mhtml, req).
- Conclusion: only 2 PDFs (IR-59636, IR-85722) are genuinely unusable on the drive; the app's copy still has to be compared (stage 3). Missing `%%EOF` is a weak signal, and the health check in check_attachment_files.py (`NO %%EOF`) over-reports - it is now its own category NO END MARKER (75) and only bad headers/zips (3) are called CORRUPT; real failures are listed in File_Open_Check. The size gap is a separate issue: ~20,000 files with the 2.6% gap open fine.
