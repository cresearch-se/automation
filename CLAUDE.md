# QA Automation Framework — Cornerstone Research

## Repo
- **Git:** `git@github.com:cresearch-se/automation.git` (SSH)
- **Default branch:** `main`
- **Working directory:** `~/QA_Automation/Code/` (this folder is the repo root)
- **Virtual env:** `.venv` — activate with `source .venv/bin/activate` (Linux) before running tests

---

## Framework Structure

```
Code/
├── src/cornerstone_automation/
│   ├── utils/
│   │   ├── api_utils.py        — HTTP GET/POST/PUT/DELETE with NTLM auth
│   │   ├── db_utils.py         — pyodbc DB connection + stored procedure execution
│   │   ├── excel_utils.py      — openpyxl-based Excel reading and column comparison
│   │   ├── json_utils.py       — simple JSON file reader
│   │   └── pandas_utilis.py    — pandas-based Excel/DB comparison utilities
│   └── sqls/
│       ├── loader.py                  — loads named SQL queries from .sql files
│       ├── ardent_queries.sql         — billable hours by office
│       ├── target_daily_queries.sql   — target vs actual hours
│       └── employee_details.sql       — employee lookups by code/name/empno
├── tests/
│   ├── TeamworkDB/              — utilization + TWDB analysis + profitability (MOST ACTIVE)
│   ├── ConsultingComp/          — API tests against CompWebAPI (staging)
│   ├── SelfEval/                — SharePoint self-eval file matching
│   ├── ShareHoldersApp/         — folder permissions + missing PRP validation
│   └── TimeOffSheets/           — time off Excel validation
└── config/
    ├── db.env                   — DB server/database names (loaded via dotenv)
    └── api.env                  — API username/password for NTLM auth
```

---

## Shared Utilities

### `api_utils.py`
- NTLM auth via `HttpNtlmAuth` (credentials from `config/api.env`)
- Functions: `get_request(url)`, `post_request(url, data)`, `put_request(url, data)`, `delete_request(url)`
- Returns `(status_code, json_body)` tuple

### `db_utils.py`
- `get_db_connection_from_env(server, database, trusted_connection=True)` — Windows Auth only (SQL auth fails with domain error 18456)
- `call_stored_procedure(conn, sp_name, named_params, as_dataframe=True)` — returns list of result sets as DataFrames

### `excel_utils.py`
- `read_excel(file_path, sheet_name)` — returns list of row dicts
- `get_excel_headers(file_path, sheet_name)` — returns header list
- `compare_columns_between_files(file1, headers1, file2, headers2)` — pairwise column comparison
- `find_key_header(headers, key_headers)` — fuzzy header matching
- `compare_rows_by_headers(r1, r2, headers, ...)` — row-level diff with noise column filtering

### `pandas_utilis.py`
- `read_excel_file`, `get_excel_sheet_names`, `find_column_by_keywords`
- `compare_db_to_excel(merged, key_cols, numeric_cols, tolerance)` — returns `[MISSING IN DB]`, `[MISSING IN XLS]`, `[VALUE MISMATCH]` errors
- **Important:** uses `numeric_cols[0]_DB` (first numeric col) as the "employee present in DB" indicator — NULL in that col fires MISSING IN DB
- `check_totals_match`, `safe_to_numeric`
- **Rule:** formatting/output changes go in the runner or test file ONLY — never in `pandas_utilis.py`

### `sqls/loader.py`
- `load_query(file_name, query_name)` — reads named queries marked with `-- query: name` in `.sql` files

---

## Test Areas

### 1. TeamworkDB — Utilization (MOST ACTIVE)

**Files:**
- `tests/TeamworkDB/test_utilization_monthly.py` — all test logic + normalizers
- `tests/TeamworkDB/run_test_utilization_monthly.py` — runner: calls pytest, parses output, generates Excel + HTML report to `tests/TeamworkDB/output/`

**Monthly update — only one line changes:**
```python
FIXTURE_FILE = "tests/TeamworkDB/fixtures/Utilization_YYYYMM.xlsx"
```
Everything else (sheet names, date ranges) derives automatically from the filename.

**DB connection:**
- Server: `SQLT1COFIN` (`T1_DB_SERVER` in db.env), Database: `ReportDevl`
- Stored procedure: `SP_Utilization_Validation` with params `StartDate`, `EndDate`, `Type` ('Summary' or 'Detail')
- Windows Auth only (`trusted_connection=True`)

**Session-scoped fixtures (fetch DB once per run, shared across tests):**
```python
db_summary_monthly   # Summary SP — monthly range
db_summary_ytd       # Summary SP — YTD range
db_detail_monthly    # Detail SP — monthly range
db_detail_ytd        # Detail SP — YTD range
```

**Test groups:**
- `test_format_validations_by_location` — offices present, titles present + correct order, no blanks/zeros
- `test_totals_validation_by_location` — subtotals and grand totals match computed sums
- `test_format_validations_by_employee` — no blanks, unique EmpNo, subtotals per title
- `test_db_comparison_us_monthly/europe_monthly/us_ytd/europe_ytd` — location-level DB comparison
- `test_db_comparison_employee_monthly` / `_ytd` — employee-level DB comparison (parametrized × 9 offices)

**Key constants:**
- `NUMERIC_COLS = ['Target_Hours', 'Target_Rev', 'Actual_Hours', 'Standard_Rev']`
- `JOIN_KEYS_BY_EMPLOYEE = ["EmpNo"]`
- Europe (Brussels/UK) both map to `Office_Code='Europe'` in DB — filtered by Excel EmpNos before merge
- DataScience/AppliedResearch use `grand_total_label="TOTAL"` (not "OFFICE TOTAL")
- EmpNo canonical type: `str` via `str(x).split('.')[0]` — strips `.0` from floats, preserves leading zeros

**Known issue — duplicate MISSING IN DB + MISSING IN XLS:**
Employees in DB with NULL `Target_Hours` fire false MISSING IN DB when absent from Excel.
Fix in `run_employee_comparison`: filter `df_db` before merge:
```python
df_db = df_db[df_db[NUMERIC_COLS[0]].notna()].copy()
```
Still under investigation for YTD tests — debug print exists (commented out) to print full merged table.

**Run commands (from `Code/` with venv active):**
```bash
# Full report (all groups + generates Excel/HTML)
python tests/TeamworkDB/run_test_utilization_monthly.py

# Individual groups
python -m pytest tests/TeamworkDB/test_utilization_monthly.py -k "format_validations or totals_validation" -v --no-cov -s --tb=short
python -m pytest tests/TeamworkDB/test_utilization_monthly.py -k "test_db_comparison_employee_monthly" -v --no-cov -s --tb=short
python -m pytest tests/TeamworkDB/test_utilization_monthly.py -k "test_db_comparison_employee_ytd" -v --no-cov -s --tb=short
```
Note: `--override-ini=addopts=` needed if running pytest directly (pyproject.toml has `--cov` in addopts).

---

### 2. ConsultingComp — API Tests

**Files:** `tests/ConsultingComp/test_api.py`, `test_employees_api.py`

- Tests against `appstaging.cornerstone.com/CompWebAPI`
- Uses NTLM auth (`api_utils.py` + `config/api.env`)
- Validates: locations, base salaries, configurations API endpoints
- Checks structure (required fields), uniqueness (locationIDs), expected values (country codes: BE, UK, US)

---

### 3. SelfEval — SharePoint File Matching

**Files:** `tests/SelfEval/test_selfeval_sharepoint_file_matches.py`

- Compares `AllOfficers-SelfEval.xlsx` vs `WOWD.xlsx` in `tests/SelfEval/fixtures/`
- Validates that PDF filenames generated from ReviewerLogin match entries in WOWD
- Uses `excel_utils.py` for reading

---

### 4. ShareHoldersApp — Folder Permissions

**Files:** `tests/ShareHoldersApp/test_folderPermissions.py`, `test_missingPRP.py`

- Parses text files with `FolderPath || Permissions` format (latin-1 encoding)
- Validates folder permission assignments
- Run from `Code/` folder: `python tests/ShareHoldersApp/test_folderPermissions.py`

---

### 5. TimeOffSheets — Excel Validation

**Files:** `tests/TimeOffSheets/test_timeoff_excel.py`, `test_timeoff_excel_basic.py`

- Validates time off Excel sheets using `excel_utils.py`
- Fixture files detected dynamically by name fragment in `fixtures/` directory

---

## Git Workflow

```bash
# Create feature branch
git checkout -b branch-name

# Stage and commit
git add <file>
git commit -m "description"

# Push
git push -u origin branch-name
# Then open PR on GitHub to merge into main
```

---

## Important Rules
1. **Never edit `pandas_utilis.py` for formatting/output** — changes go in the runner or test file only
2. **Windows Auth only** for DB — `trusted_connection=True`, no SQL username/password
3. **Run from `Code/` directory** — relative paths in test files assume this as the working directory
4. **Use `-k` flag** for parametrized tests, not `::function_name`
5. **SSH remote** — `git@github.com:cresearch-se/automation.git` (HTTPS fails on this server without a PAT)

---

## Context Files (Claude must maintain these)

Live notes are stored in `.claude/context/`. **At the start of every session, read all three files before doing any work.** Update them whenever something changes — don't wait to be asked.

| File | Update when |
|---|---|
| [`teamworkdb.md`](.claude/context/teamworkdb.md) | SP behavior changes, new fixture quirks, new test groups, bug root causes found |
| [`decisions.md`](.claude/context/decisions.md) | A non-obvious architectural or coding decision is made and the reason should be remembered |
| [`todo.md`](.claude/context/todo.md) | An issue is opened or closed, a task is started or finished, a new recurring step is discovered |

**Rules:**
- Update context files silently as work progresses — never wait to be asked
- Create new topic files as needed if a new area grows complex enough to deserve its own file
- Mark items in `todo.md` as done (or remove them) as soon as they are resolved — never leave stale entries
- Add to `decisions.md` any time a "why did we do it this way?" question comes up and gets answered
- Keep `teamworkdb.md` as the single source of truth for utilization test mechanics — if CLAUDE.md and `teamworkdb.md` conflict, `teamworkdb.md` wins (it is more detailed)
