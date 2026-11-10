# QCFootprints — Footprints app vs DB validation

## Workflow (code here, run locally — see memory: code-here-run-on-local-machine)
- App not reachable from the dev server. Scripts are written here, user runs them on the local machine and pastes output back.
- Branch: `qc_footprints` (tracks origin/main by user request; push with `git push origin qc_footprints`).

## Extractor: `tests/QCFootprints/extract_footprints.py`
- POST `https://crcofpsc.cornerstone.com/footprints/servicedesk/api/search/advanced/run?_dc=..&limit=500&page=N&start=(N-1)*500`
- Body = saved search definition (~17.9k chars) in `tests/QCFootprints/request_body.txt` (local only, gitignored).
- Auth: NTLM does NOT work (returns HTML login page). Use `--auth cookie` (browser Cookie header, ~584 chars; FP_COOKIE in `config/footprints.env`, env file was not being picked up — user pastes at prompt instead). Cookie expires.
- Needs `--insecure` (internal CA not trusted by Python requests).
- Response: `{totalCount, pageData:[{total, rowsData:[{rowValues:[40 values]}], columns:[40 {title,fieldName}]}]}`. Column names come from `columns[].title`.
- Output: raw `output/raw/page_NNN.json`, Excel `fixtures/QCFootprints_<YYYYMMDD>.xlsx` (both gitignored). Raw pages are cached: use `--refresh` (or delete output/raw) before a full run after a sample run.
- App total drifts upward (30307 -> 30309 -> 30322 over a few days); new tickets during a run can shift offset paging.
- 25-row sample run verified OK (2026-11-09). Full run (61 pages) not yet done.

## DB export: `tests/QCFootprints/fixtures/QCFootprint_DB_Details.xlsx` (50 MB, force-added to git)
- Sheet1 headers (20): Ticket Number, Ticket Title, Case Name, Case / Matter Number, Status, Priority, Responsible Officer, Officer Location, Ticket Created Date, Created By, Last Updated Date, Last Updated By, Requester First Name, Requester Last Name, Requester Email, Flagged Deleted in FootPrints, Ticket Internal ID, Has Attachments, Attachment Count, Attachments Not Exported.
- Sheet1 has 33,689 non-empty rows but only **31,226 real ticket rows** (those with Ticket Internal ID). The other 2,463 rows are **spilled ticket description text** (multi-line text broke rows; only 1–2 cells filled, interleaved among real rows). Must drop rows with null Ticket Internal ID. One junk cell is an openpyxl ArrayFormula, one a datetime.
- Real rows: Ticket Number unique (no dups), all `IR-nnnnn`; created 2021-01-04 .. 2026-11-03 (export ~Nov 3-4 2026).
- Flagged Deleted: No 30,224 / Yes 1,002. App total ~30,322 — likely the app excludes deleted tickets (unconfirmed).
- Priority: DB has literal string 'NULL' (26,768) plus real nulls; Status values Closed/Open/Pending; Officer Location codes (CRMP, CRNY, CRB, COR, CRCH, CRDC, CRLA, CRSF, CRUK, CRLO, NY).
- Sheet2: 30 rows, only a "Ticket Number" column containing header-name text (stray, ignore).
- Dates: DB datetimes are **5 hours ahead** of the app's (DB 2021-01-04 01:28:03.540 vs app 2021-01-03T20:28:03.540 for IR-57299) -> DB is UTC, app is Eastern; offset is 4h in DST. Needs tz conversion before comparing.

## Probable column mapping (unconfirmed with user)
DB Ticket Number <-> app Record Number (also Request Number); Ticket Internal ID <-> ID; Ticket Title <-> Title; Case Name <-> Client Matter Name; Case / Matter Number <-> Client Matter Number; Status <-> Status; Priority <-> Priority; Responsible Officer <-> Responsible Officer; Officer Location <-> Officer Location; Ticket Created Date <-> Created On; Created By <-> Created By; Last Updated Date <-> Updated On; Last Updated By <-> Updated By; Requester First/Last Name <-> First Name/Last Name; Requester Email <-> Email Address. DB-only: Flagged Deleted, Has Attachments, Attachment Count, Attachments Not Exported.

## First full comparison (report qc_footprints_report_20261109_122436.xlsx, 2026-11-09)
- App 30,325 tickets, DB live 30,224 (1,002 deleted excluded). All 30,224 DB live tickets found in app. Created On: 0 mismatches over the whole range -> DST handling (America/New_York) is correct.
- 958 mismatch cells / 774 tickets. Causes: (1) Created/Updated By 'Email Request' vs DB '-315' = 522 (DB holds user id, not name); (2) accents flattened in DB (Ł->L, ł->l) = 331 (Created By 199, First Name 132); (3) ticket changed in app after DB export (~50 tickets: Status 30, Updated On 43, Updated By ~10, matter/officer 1-2) -> clears with a fresh DB export; (4) Title: 3 leading/embedded quotes stripped, 1 non-ASCII -> '?', 1 title starting with '-' came through blank in the DB Excel (export artifact), 2 non-breaking hyphens; (5) Client Matter Number text values ('test', 'a5345') flagged falsely = BUG in compare script, fixed (text fallback when either side is non-numeric).
- Missing in DB: 101 = 95 newer than DB's highest ticket (IR-88530) + 6 inside range (IR-65403..65405, IR-65418..65420) -> ask dev.
- Script now: labels each mismatch with a Category, tags tickets whose app Updated On > DB export cutoff as CHANGED AFTER DB EXPORT, splits Missing_in_DB into NEWER THAN DB EXPORT / WITHIN DB RANGE, adds Mismatch_Categories sheet.
- Questions for the dev: why is created/updated-by -315 instead of 'Email Request'; why are non-ASCII characters flattened/'?'-replaced; why are tickets IR-65403..65405 and IR-65418..65420 absent from the DB export.

## Report publishing (decision 2026-11-09)
- Single report file with a fixed name, overwritten each run: `tests/QCFootprints/output/qc_footprints_report.xlsx`, tracked in git so everyone (and Claude) can read it from the repo. `--sample` runs write `output/qc_footprints_report_sample.xlsx` (ignored) so they never overwrite the full report. No timestamped copies, no `reports/` folder.
- `tests/QCFootprints/.gitignore` now ignores only output/raw/, probe/debug files, *.pkl cache, *_sample.xlsx and request_body.txt. fixtures/ (app extract + DB export) are tracked.
- Publish flow (user, local): git pull --autostash origin qc_footprints; run compare_footprints.py; git add tests/QCFootprints/output/qc_footprints_report.xlsx; commit; push.

## Dev validation notes (2026-11-10) and attachments
- Dev checklist: all IR tickets 1/1/2021-Nov 3 present; UI count == DB; spot check ID/case/title; attachment counts UI vs DB; attachment names identical UI vs DB field; open attachments (D drive); attachment location in DB field correct; D drive matches; flag tickets with real passwords in description (ignore plain 'change password'); question for Gurpreet: scan attachments for passwords? Kim still working on zip attachments.
- Checked from existing data: DB IR-57299..88530, only gaps = the 6 missing tickets; 31,226 rows - 1,002 deleted = 30,224 + 101 missing = 30,325 app. DB Has Attachments consistent with Attachment Count (12,556 tickets with, 18,670 without, 32,575 attachments; Attachments Not Exported all 0). DB has counts only, no names/paths.
- App saved search (#11901 'IR >2021 Tickets') has NO attachment fields and no 'Possible Credential' field (field picker: only Common 3 + Information Resources.Request 19). Ask Kim where Possible Credential comes from.
- Attachment API (found via browser DevTools): GET .../footprints/servicedesk/api/item/attachment/<Item ID>/<ID>?limit=25&page=1&start=0 -> {"data":[{name,size,extension,creationDate,uri,mimeType}],"success":true}. <Item ID>=21103 (same for all, = _metaId), <ID> = ticket 'ID' column (IR-57299 -> 7399). uri = item/attachment/<x>/<definitionId 6518>/<ID>/<file>?<signed token> (token stripped in output).
- Script `extract_attachments.py` (written + tested vs fake server, not yet run against real app): default scope = DB-flagged tickets with attachments + random 500 without; --all, --tickets, --first, --probe. Caches output/raw_att/<ID>.json (ignored), writes fixtures/QCFootprints_Attachments.xlsx (fixed name) (sheets Attachments, Ticket_Counts). Next: compare UI count vs DB Attachment Count in compare script; password scan of app Description (mask values - report is public in repo).
- compare_footprints.py now also reads fixtures/QCFootprints_Attachments.xlsx (if present; --attachments PATH, --no-attachments to skip): UI count (Ticket_Counts sheet) vs DB 'Attachment Count' per checked ticket. Report gets an ATTACHMENTS block in Summary + sheets Attachment_Count_Check (categories: CHANGED AFTER DB EXPORT / MORE / FEWER ATTACHMENTS IN UI THAN DB) and Attachments_Not_Fetched. Also flags sampled 'DB says none' tickets that have UI attachments. Tested with fake data only; stdout set to errors=replace (cp1252 consoles crashed on odd chars). Attachment NAME/path comparison still needs the DB-side file names/paths from Kim.
