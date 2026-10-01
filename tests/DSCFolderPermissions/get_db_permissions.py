"""
get_db_permissions.py
Connects to SQLT4COSTAGDW / DSCCaseFolder (Windows Auth), runs the permissions
query, and writes results to output/db_permissions.csv.

Run from the repo root:
    python tests/DSCFolderPermissions/get_db_permissions.py
"""

import os
import sys
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

# Locate repo root (two levels up from this file) and load config/db.env
REPO_ROOT = Path(__file__).parents[2]
load_dotenv(REPO_ROOT / "config" / "db.env")

sys.path.insert(0, str(REPO_ROOT / "src"))
from cornerstone_automation.utils.db_utils import get_db_connection_from_env, select_query

OUTPUT_DIR              = Path(__file__).parent / "output"
OUTPUT_FILE             = OUTPUT_DIR / "db_permissions.csv"
SERVERS_FILE            = OUTPUT_DIR / "servers.txt"
PROJECT_PATHS_FILE      = OUTPUT_DIR / "project_paths.txt"
IGNORED_PRINCIPALS_FILE = OUTPUT_DIR / "ignored_principals.csv"

SERVER   = os.getenv("DSC_DB_SERVER",   "SQLT4COSTAGDW")
DATABASE = os.getenv("DSC_DB_DATABASE", "DSCCaseFolder")

QUERY          = "SELECT * FROM [vw_lan_CaseFolderAccessUsers]"
IGNORED_QUERY  = "SELECT PrincipalName FROM [Lan_IgnoredPrincipal] WHERE IsActive = 1"


def main():
    print(f"Connecting to {SERVER} / {DATABASE} ...")
    conn = get_db_connection_from_env(SERVER, DATABASE, trusted_connection=True)

    print("Running permissions query ...")
    df = select_query(conn, QUERY, as_dataframe=True)

    print("Running ignored principals query ...")
    ignored_df = select_query(conn, IGNORED_QUERY, as_dataframe=True)
    conn.close()

    print(f"Permissions query returned {len(df)} rows")
    print(f"Ignored principals: {len(ignored_df)} entries")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8")
    print(f"Saved to: {OUTPUT_FILE}")

    ignored_df.to_csv(IGNORED_PRINCIPALS_FILE, index=False, encoding="utf-8")
    print(f"Ignored principals saved to: {IGNORED_PRINCIPALS_FILE}")

    # Extract distinct servers from FolderPath (e.g. \\CRCHFS1\Projects...) -> \\CRCHFS1
    servers = (
        df["FolderPath"]
        .dropna()
        .str.strip()
        .str.extract(r'^(\\\\[^\\]+)', expand=False)
        .dropna()
        .str.upper()
        .unique()
    )
    servers = sorted(servers)
    SERVERS_FILE.write_text("\n".join(servers), encoding="utf-8")
    print(f"Servers list ({len(servers)}) saved to: {SERVERS_FILE}")
    for s in servers:
        print(f"  {s}")

    # Extract distinct \\SERVER\ProjectsXXX paths (server + share/project level)
    project_paths = (
        df["FolderPath"]
        .dropna()
        .str.strip()
        .str.extract(r'^(\\\\[^\\]+\\[^\\]+)', expand=False)
        .dropna()
        .str.upper()
        .unique()
    )
    project_paths = sorted(project_paths)
    PROJECT_PATHS_FILE.write_text("\n".join(project_paths), encoding="utf-8")
    print(f"Project paths ({len(project_paths)}) saved to: {PROJECT_PATHS_FILE}")


if __name__ == "__main__":
    main()