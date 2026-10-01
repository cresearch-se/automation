# DSC Folder Permissions Validation

Validates that Windows NTFS folder permissions on DSC case folders match what is recorded in the database (`vw_lan_CaseFolderAccessUsers`).

---

## What It Does

1. **Pulls** all DSC case folder permission records from the database
2. **Reads** the actual Windows ACL permissions from the file servers
3. **Compares** both and reports any differences

### Difference Types in the Report

| DiffType | Meaning |
|---|---|
| `MISSING_IN_WINDOWS` | DB says a group should have access but Windows ACL does not have it |
| `MISSING_IN_DB` | Windows ACL has a group for a folder but DB has no record of it |
| `NO_ACCESS` | Could not read Windows ACL — no access to that server/folder, verification skipped |

---

## Folder Structure

```
tests/DSCFolderPermissions/
├── scripts/
│   └── Get-FolderPermissions.ps1   — PowerShell: reads Windows ACLs from file servers
├── output/                          — all generated files land here (gitignored)
│   ├── db_permissions.csv           — permissions from DB (Step 1 output)
│   ├── servers.txt                  — distinct servers from DB (Step 1 output)
│   ├── project_paths.txt            — \\SERVER\ProjectsXXX paths from DB (Step 1 output)
│   ├── folder_permissions.csv       — Windows ACLs (Step 2 output)
│   └── comparison_report.csv        — final diff report (Step 3 output)
├── get_db_permissions.py            — Python: pulls DB data and generates output files
├── compare_permissions.py           — Python: compares both CSVs and produces report
└── README.md                        — this file
```

---

## Prerequisites

### For Steps 1 and 3 (Python scripts — run from the repo)
- Python 3.x with the project venv active
- Network access to `25SQLT4COTEST` SQL Server (VPN or office network)
- `config/db.env` present with:
  ```
  DSC_DB_SERVER=25SQLT4COTEST
  DSC_DB_DATABASE=DSCCaseFolder
  ```

### For Step 2 (PowerShell script — run from any Windows machine with file server access)
- PowerShell 5.1 or later (built-in on Windows 10 / Server 2016+)
- Network access to the DSC file servers (`\\CRCHFS1`, `\\CRNYFS2`, etc.)
- A service/admin account with access to enumerate all file server shares (provided by the dev team)
- The script will prompt for username and password when it starts

---

## Execution Steps

### Step 1 — Pull permissions from the database

Run from the repo root with the venv active:

```bash
python tests/DSCFolderPermissions/get_db_permissions.py
```

**Output files generated in `output/`:**
- `db_permissions.csv` — full permissions data from the DB
- `servers.txt` — list of distinct file servers (used by Step 2)
- `project_paths.txt` — list of `\\SERVER\ProjectsXXX` paths (fallback if no service account)

---

### Step 2 — Read Windows ACL permissions from file servers

> Run this on a machine that has network access to the DSC file servers.

Copy these two files to that machine:
- `tests/DSCFolderPermissions/output/servers.txt`
- `tests/DSCFolderPermissions/scripts/Get-FolderPermissions.ps1`

Place both files in the same folder, then run:

```powershell
powershell -ExecutionPolicy Bypass -File .\Get-FolderPermissions.ps1
```

The script will prompt for the service account username and password, then scan all servers in `servers.txt`, enumerating every `Projects*` share and its case folders.

**Output file generated:**
- `folder_permissions.csv` — Windows ACL entries for all accessible case folders

Once done, copy `folder_permissions.csv` back into `tests/DSCFolderPermissions/output/` in the repo.

---

### Step 3 — Compare and generate report

Run from the repo root with the venv active:

```bash
python tests/DSCFolderPermissions/compare_permissions.py
```

**Output file generated:**
- `output/comparison_report.csv` — all differences with `DiffType`, `FolderPath`, `GroupName`, and `Detail` columns

---

## Config

Database connection details are stored in `config/db.env` (gitignored — local only):

```
DSC_DB_SERVER=25SQLT4COTEST
DSC_DB_DATABASE=DSCCaseFolder
```

---

## Notes

- The PowerShell script skips case folders it cannot access and marks them as `NO_ACCESS` in the CSV so they appear in the report rather than being silently ignored
- Inherited ACL entries (system defaults like `SYSTEM`, `Administrators`) are excluded from the comparison — only explicit group assignments are compared
- All output files are gitignored — do not commit them
