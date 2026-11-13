# QCFootprints - Footprints app vs database vs attachment files

## 1. What we are testing

Footprints is the ticket system (tickets are numbered `IR-nnnnn`). Its data was exported from the database (table
`IR_Ticket_Summary`) together with the attachment files (on the export server's D drive). This folder checks that
the export is complete and correct, by comparing three sources against each other:

| Source | What it is | How we get it |
|---|---|---|
| **App (UI/API)** | What users see in Footprints today | Footprints API (`extract_api_tickets.py`, `extract_api_attachments.py`) |
| **DB export** | What was exported from the database | Direct DB query (`extract_db_footprints.py`) |
| **Attachment files** | The actual files in `D:\FPextract\out\item\attachment\6257\6518\<ticket internal ID>\<file>` | Read from drive `Z:` (`check_attachment_files.py`, `open_check_files.py`) |

What is checked:

1. **Ticket fields** - app vs DB, field by field (title, status, dates, created/updated by, requester, ...). Key: `IR-nnnnn`.
2. **Attachments, app vs DB** - number of attachments, file names, file sizes per ticket.
3. **Attachments, DB vs drive** - every DB path points to a file that exists, sizes match, files are not damaged, files really open.
4. **Credentials in ticket history** - tickets the DB flags `Possible Credential = Yes` are reviewed, and tickets flagged `No` are checked for missed passwords.

Not checked yet: the content of attachments (password scan inside files), a byte-for-byte compare with the app's download,
and the app's Description text vs the DB's description history. See section 6.

## 2. Flow of testing

```
extract_db_footprints.py      DB          ->  fixtures/QCFootprint_DB_Details.xlsx
extract_api_tickets.py        App API     ->  fixtures/QCFootprints_<YYYYMMDD>.xlsx
extract_api_attachments.py    App API     ->  fixtures/QCFootprints_Attachments.xlsx
            |
            v
compare_footprints.py         reads the three files  ->  output/qc_footprints_report.xlsx  (REWRITES the report)
            |
            v
check_attachment_files.py     DB paths + UI data vs files on Z:  ->  ADDS 4 sheets to the report
            |
            v
open_check_files.py           really opens the suspicious files    ->  ADDS 2 sheets to the report
```

**Run order matters:** `compare_footprints.py` rewrites the whole report and drops the sheets added by the two later
scripts, so after every compare, re-run `check_attachment_files.py` and then `open_check_files.py`.

All output files have **fixed names** (overwritten each run, tracked in git). Close the report in Excel before running
anything that writes it, otherwise Windows blocks the write.

## 3. Where each script runs

| Script | Runs on | Why |
|---|---|---|
| `extract_db_footprints.py` | Your own machine | Needs the DB (Windows Auth) |
| `extract_api_tickets.py`, `extract_api_attachments.py` | Your own machine | The app is not reachable from the dev server; needs a browser cookie |
| `compare_footprints.py` | Anywhere (reads the three Excel files only) | |
| `check_attachment_files.py`, `open_check_files.py` | The dev server | Drive `Z:` (= `\\25sqlt4cotest\d$`, the export server's D drive) is mapped there |

Run everything from the repo root with the virtual env active. The libraries the file-open check needs (pymupdf, pillow,
pypdf, xlrd) are in `requirements.txt`; the scripts also use pandas, openpyxl, requests and pyodbc from the project's venv.

## 4. How to run

### One-time setup
- `config/db.env` (git-ignored): `FOOTPRINTS_DB_SERVER`, `FOOTPRINTS_DB_DATABASE` (or pass `--server` / `--database`).
- `config/footprints.env` (git-ignored): `FP_COOKIE` = the full `Cookie` header copied from the browser (DevTools >
  Network > Request Headers). NTLM does not work for the app. The cookie expires, so copy a fresh one for each run.
- `tests/QCFootprints/request_body.txt` (git-ignored): the saved-search POST body copied from the browser
  (DevTools > Network > the `run` request > Payload).
- Never put passwords or cookies in committed files.

### Full refresh (in this order)
```bash
# 1. DB export (a few minutes). --check = dry run, writes nothing
python tests/QCFootprints/extract_db_footprints.py

# 2. App tickets (about 61 pages for ~30k tickets). Needs a fresh cookie and --insecure
python tests/QCFootprints/extract_api_tickets.py --auth cookie --insecure
#    --probe fetches 5 rows to test the setup; --refresh ignores cached pages

# 3. App attachments (tickets the DB says have attachments + a random 500 that have none)
python tests/QCFootprints/extract_api_attachments.py --auth cookie --insecure
#    --add IR-1,IR-2  adds single tickets to the existing file;  --all = every ticket

# 4. Compare (close the report in Excel first)
python tests/QCFootprints/compare_footprints.py
#    --sample = quick 25-row check (writes a separate, untracked _sample file)

# 5. Attachment files on drive Z: (dev server, about 3 minutes)
python tests/QCFootprints/check_attachment_files.py
#    --limit N = trial on N tickets, not written;  --no-health = skip the health stage;  --deep = read every byte of zip/Office files

# 6. Really open the suspicious files (dev server)
python tests/QCFootprints/open_check_files.py
#    --sample N = random text-like files to spot-check (default 200);  --include-mismatch adds the extension mismatches
```

Then commit and push the fixtures and `output/qc_footprints_report.xlsx`. GitHub does not display `.xlsx`; download
the file (Download raw file) and open it in Excel.

If a push is rejected because the other machine pushed first: `git pull --no-rebase --no-edit origin qc_footprints`, then push again.

## 5. Reading the output: `output/qc_footprints_report.xlsx`

Start with **Summary** and **File_Check_Summary**; they hold the counts. Every other sheet lists only problems, so an
empty sheet means nothing was found.

### Written by `compare_footprints.py`

| Sheet | What it shows |
|---|---|
| `Summary` | Counts of tickets compared, mismatches per category, missing tickets, attachment and credential results. Start here. |
| `Column_Check` | Per column: how many tickets match / mismatch. Shows which fields have problems. |
| `Mismatch_Categories` | Mismatches grouped by cause (see categories below). |
| `Value_Mismatches` | One row per differing cell: ticket, column, app value, DB value, category. |
| `Missing_in_DB` | Tickets in the app but not in the DB export. `NEWER THAN DB EXPORT` is expected; `WITHIN DB RANGE - INVESTIGATE` is a real gap. |
| `Missing_in_App` | Tickets in the DB but not in the app. |
| `Deleted_in_DB` | Tickets the DB flags as deleted that still exist in the app (information). |
| `Attachment_Count_Check` | Tickets where the number of attachments differs between app and DB. |
| `Attachments_Not_Fetched` | Tickets whose attachments could not be fetched from the app (re-run the attachment extract). |
| `Attachment_Name_Check` | Attachment names that differ between app and DB. |
| `Attachment_Size_Check` | Total attachment size per ticket, app vs DB. |
| `Attachment_Path_Check` | DB attachment paths that do not follow the expected folder pattern, and file names that differ between the app and the path (numeric prefix, illegal characters replaced by `_`, `~` shortening). |
| `Credential_Review` | Tickets the DB flags `Possible Credential = Yes`, each classified (see below). Values are masked `[REDACTED]`, but the context can show user names and emails - think before sharing. |
| `Flag_Check_No` | Tickets flagged No/blank that mention a password word, to find flags the DB missed. Same masking. |

**Mismatch categories** (`Mismatch_Categories`, `Value_Mismatches`):
- `CHANGED AFTER DB EXPORT` - the app's `Updated On` is later than the DB export time. Expected, not a defect; a fresh DB export clears it.
- `USER ID INSTEAD OF NAME` - the DB holds an id (e.g. `-315`) where the app shows a name ('Email Request').
- `ACCENTS / SPECIAL CHARACTERS`, `NON-ASCII REPLACED BY '?'`, `QUOTES REMOVED` - text differs in characters only.
- `DB BLANK APP HAS VALUE` - the DB is empty where the app has data.
- `OTHER` - anything else; look at these first.

**Credential classes** (`Credential_Review`, `Flag_Check_No`): `LIKELY REAL PASSWORD - REVIEW` (a value after a
password label; a human must check) | `MEETING PASSCODE` (Zoom/Teams style, not a secret) | `MENTIONS PASSWORD WORD ONLY`
(e.g. 'reset your password') | `NO PASSWORD WORD FOUND`. This is a heuristic scan, not proof.

### Added by `check_attachment_files.py`

| Sheet | What it shows |
|---|---|
| `File_Check_Summary` | Counts for the three sheets below. Start here. |
| `Attachment_File_Check` | Existence and size problems, one row per file. A row with a **blank attachment name** is a ticket-level total row (`DB x MB, disk y MiB`). Columns: Record Number, Attachment name, UI size (bytes), Disk size (bytes), DB path, Mapped path, Note, Category. |
| `Attachment_File_Health` | Files that look damaged, judged by reading the file content (header, end marker, zip index), not the name. Only problems are listed. |
| `Attachment_Orphan_Files` | Files in a ticket's folder that no DB path points to (empty so far). |

Categories in `Attachment_File_Check`:
- `FILE MISSING ON DRIVE` - the DB path has no file.
- `ZERO-BYTE FILE` - the file is empty.
- `DISK ~2.6% SMALLER THAN UI` - disk size / UI size is 0.973-0.976. Cause still unknown (see section 6).
- `SIZE DIFFERS (UI vs DISK)` - any other size difference.
- `TICKET TOTAL SIZE != DB` - ticket-level row; the same gap added up.
- `IN UI BUT NOT IN DB` - an app attachment has no DB path.
- `CREATED AFTER DB EXPORT` / `CHANGED AFTER DB EXPORT` - the file or ticket is newer than the export. Expected, not a defect.
- `UNEXPECTED DB PATH`, `PATH COUNT != NAME COUNT` - malformed DB values.

Categories in `Attachment_File_Health` (a flag is a hint; every flagged PDF or image is then really opened):
- `FLAGGED BUT OPENS FINE (INFO)` - e.g. a PDF without the `%%EOF` end marker; it opens, so it is not a defect.
- `NO END MARKER AND CANNOT BE OPENED` - looks truncated and does not open. **Real problem.**
- `CORRUPT (BAD HEADER OR ZIP)` - the header or zip structure is wrong (so far: XML/HTML saved with an `.xls` extension, i.e. mislabelled).
- `EXTENSION / CONTENT MISMATCH` - the content is a different type than the extension (e.g. `.png` that is a TIFF). Usually harmless.
- `CANNOT READ FILE` - the file could not be read at all.

### Added by `open_check_files.py`

| Sheet | What it shows |
|---|---|
| `File_Open_Summary` | Counts per source (flagged files / mismatches / text spot check) and per extension. |
| `File_Open_Check` | One row per file that was really opened: Result, Detail, the earlier health finding, path. **Failures are first.** Filter `Result = FAILS TO OPEN` to see the files that do not open. |

Results: `OPENS OK` | `OPENS WITH WARNINGS` (the parser repaired something, but the content reads) | `FAILS TO OPEN` |
`NO PARSER` (type we cannot open here, e.g. rtf, rar, mp3, not a failure). PDFs are opened with PyMuPDF (tolerant like a
real viewer; pypdf is too strict and gave false failures), images with Pillow, old `.xls` with xlrd, `.eml` with the mail parser.

## 6. Current status and open items (as of 2026-11-12)

- App 30,325 tickets, DB 30,230 live tickets compared. 612 value mismatches, only two causes: 522 `-315` instead of
  'Email Request', and 90 changed after the DB export (43 tickets). 95 tickets are newer than the export.
- Attachments: counts, names and total sizes match between app and DB except 4 tickets changed after the export.
- Drive: all 32,266 DB paths exist on `Z:`; none zero-byte; no orphans.
- **Size gap (main open issue):** 27,386 files are about 2.6% smaller on disk than the size the app reports (751 more by 3-8%).
  Unknown whether the app's size field is inflated or the export rewrote the files. About 20,000 of these PDFs read fine.
- **Files that do not open:** 2 PDFs (IR-59636, IR-85722; both sizes are exact multiples of 64 KiB, so they look cut off), and
  2 mislabelled `.xls` (IR-73830, IR-84212). 74 other flagged files open fine.
- 13 files have a different type than their extension (harmless so far).
- Credentials: of 241 flagged live tickets, 67 are `LIKELY REAL PASSWORD - REVIEW` (needs a human); 7 tickets the DB flag missed
  (IR-58782, IR-59410, IR-71937, IR-73289, IR-78717, IR-84144, IR-87367).

**Not built yet:**
- Stage 3: download attachments from the app and compare size/hash with the drive copy (needs a fresh cookie). This settles the size gap.
- Stage 4: scan attachment contents for passwords (needs a go-ahead; zip files wait for Kim; images need OCR). Rough estimate 1-2 hours without OCR, not measured.
- Existence check for the 310 paths on tickets the DB flags as deleted.
- Compare the app's Description text with the DB's description history; a `Credential_Flag_Mismatch` sheet; a pytest wrapper.

Detailed working notes, root causes and lessons learned: `.claude/context/qcfootprints.md`.

## 7. Things worth knowing (lessons)

- **Do not paste the DB export through SSMS into Excel.** A title starting with a quote made Excel swallow several rows into
  one cell, which produced fake missing tickets, blank flags and '?' characters. `extract_db_footprints.py` writes the Excel directly.
- Excel allows 32,767 characters per cell, so long text is split over continuation columns (`Full Description History (part 2..N)`);
  `compare_footprints.py` joins them back.
- DB timestamps are UTC, app timestamps are US Eastern; the compare converts them. Deleted-in-DB tickets are excluded.
- The DB attachment-size column is in MiB (labelled MB) and the DB has no per-file size.
- The report is tracked in git and can show user names and emails in the masked credential context. Review before sharing widely.
- Missing `%%EOF` does not mean a PDF is broken. Only trust `File_Open_Check`.
