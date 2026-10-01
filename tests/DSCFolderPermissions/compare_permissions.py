"""
compare_permissions.py
Compares Windows folder permissions (from Get-FolderPermissions.ps1 output)
against database permissions (from get_db_permissions.py output).

Join key : FolderPath (case-insensitive UNC path)
Compare  : GroupName from DB vs Identity from Windows ACL
           (domain prefix stripped from Windows Identity before comparing,
            e.g. CORNERSTONE\\WC-BoxesGowrisankaran -> WC-BoxesGowrisankaran)

Difference types:
  MISSING_IN_WINDOWS  — DB says a group should have access, but Windows ACL has no matching entry
  MISSING_IN_DB       — Windows ACL has a group for that folder, but DB has no record of it
  (No rights-level diff because the DB view does not store permission levels)

Run from the repo root:
    python tests/DSCFolderPermissions/compare_permissions.py
"""

from pathlib import Path
import pandas as pd

OUTPUT_DIR        = Path(__file__).parent / "output"
FOLDER_PERMS_FILE = OUTPUT_DIR / "folder_permissions.csv"
DB_PERMS_FILE     = OUTPUT_DIR / "db_permissions.csv"
REPORT_FILE       = OUTPUT_DIR / "comparison_report.csv"


def strip_domain(identity: str) -> str:
    """CORNERSTONE\\GroupName  ->  GroupName"""
    return identity.split("\\")[-1].strip().upper()


def normalise_path(path: str) -> str:
    return path.strip().upper()


def main():
    for f in [FOLDER_PERMS_FILE, DB_PERMS_FILE]:
        if not f.exists():
            raise FileNotFoundError(
                f"{f} not found.\n"
                "Run Get-FolderPermissions.ps1 and get_db_permissions.py first."
            )

    folder_df = pd.read_csv(FOLDER_PERMS_FILE, dtype=str).fillna("")
    db_df     = pd.read_csv(DB_PERMS_FILE,     dtype=str).fillna("")

    # Separate NO_ACCESS rows — folders we couldn't read at all
    no_access_df  = folder_df[folder_df["FolderType"] == "NO_ACCESS"].copy()
    folder_cases  = folder_df[folder_df["FolderType"] == "Case"].copy()
    folder_cases  = folder_cases[folder_cases["IsInherited"] == "False"].copy()

    no_access_paths = set(normalise_path(r["FolderPath"]) for _, r in no_access_df.iterrows())

    print(f"Folder CSV (case, non-inherited) : {len(folder_cases)} rows")
    print(f"Folder CSV (no access)           : {len(no_access_df)} folders")
    print(f"DB CSV                           : {len(db_df)} rows")

    # Build set: (normalised_path, normalised_group)
    # Windows side
    windows_set = set()
    for _, row in folder_cases.iterrows():
        path  = normalise_path(row["FolderPath"])
        group = strip_domain(row["Identity"])
        windows_set.add((path, group))

    # DB side
    db_set = set()
    for _, row in db_df.iterrows():
        path  = normalise_path(row["FolderPath"])
        group = row["GroupName"].strip().upper()
        db_set.add((path, group))

    results = []

    # In DB but not in Windows
    for path, group in db_set - windows_set:
        if path in no_access_paths:
            results.append({
                "DiffType":  "NO_ACCESS",
                "FolderPath": path,
                "GroupName":  group,
                "Detail":    "No access to server/folder - Windows ACL could not be read, verification skipped",
            })
        else:
            results.append({
                "DiffType":    "MISSING_IN_WINDOWS",
                "FolderPath":  path,
                "GroupName":   group,
                "Detail":      "DB has this group for the folder but Windows ACL does not",
            })

    # In Windows but not in DB
    for path, group in windows_set - db_set:
        results.append({
            "DiffType":    "MISSING_IN_DB",
            "FolderPath":  path,
            "GroupName":   group,
            "Detail":      "Windows ACL has this group for the folder but DB does not",
        })

    report_df = pd.DataFrame(results).sort_values(["DiffType", "FolderPath"])

    counts = report_df["DiffType"].value_counts() if not report_df.empty else {}
    print("\n=== Comparison Summary ===")
    print(f"  MISSING_IN_WINDOWS : {counts.get('MISSING_IN_WINDOWS', 0)}")
    print(f"  MISSING_IN_DB      : {counts.get('MISSING_IN_DB', 0)}")
    print(f"  NO_ACCESS          : {counts.get('NO_ACCESS', 0)}  (could not verify - insufficient permissions)")
    print(f"  TOTAL DIFFS        : {len(report_df)}")
    if report_df.empty:
        print("  All permissions match!")

    print("""
NOTE: This report is DB-driven — it only checks folders the DB knows about.
  - MISSING_IN_WINDOWS / NO_ACCESS / RIGHTS_MISMATCH are reliable findings.
  - MISSING_IN_DB cannot be detected: folders that exist on the file server
    but are not in the DB are invisible without admin-level share enumeration.
    Please check with the team if admin access can be provided to cover this gap.
""")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    report_df.to_csv(REPORT_FILE, index=False, encoding="utf-8")
    print(f"Report saved to: {REPORT_FILE}")


if __name__ == "__main__":
    main()