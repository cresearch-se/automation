"""
Export the Footprints ticket data straight from the database into the DB export Excel that
compare_footprints.py reads. This replaces copy/paste from SSMS (which breaks on quotes, tabs and
line breaks inside titles/descriptions and shifts rows).

Run on the LOCAL machine (DB reachable with your Windows login), from the repo root:

    python tests/QCFootprints/extract_db_footprints.py

Output (overwritten each run, fixed name, tracked in git):
    tests/QCFootprints/fixtures/QCFootprint_DB_Details.xlsx

Connection (Windows Authentication only):
    Add these two lines to config/db.env (gitignored), or pass --server / --database:
        FOOTPRINTS_DB_SERVER=<server>
        FOOTPRINTS_DB_DATABASE=<database that holds IR_Ticket_Summary>

Options:
    --query "SELECT ..."   override the default  SELECT * FROM IR_Ticket_Summary
    --out PATH             write somewhere else
    --check                connect, run the query, print row/column counts, do not write a file
"""
import argparse
import os
import re
import sys
import warnings

import pandas as pd
from dotenv import load_dotenv

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
sys.path.insert(0, os.path.join(REPO, "src"))

from cornerstone_automation.utils.db_utils import get_db_connection_from_env  # noqa: E402

DEFAULT_OUT = os.path.join(HERE, "fixtures", "QCFootprint_DB_Details.xlsx")
DEFAULT_QUERY = "SELECT * FROM IR_Ticket_Summary"

# Columns compare_footprints.py needs; missing ones are reported, not fatal.
EXPECTED = [
    "Ticket Number", "Ticket Title", "Case Name", "Case / Matter Number", "Status", "Priority",
    "Responsible Officer", "Officer Location", "Ticket Created Date", "Created By",
    "Last Updated Date", "Last Updated By", "Requester First Name", "Requester Last Name",
    "Requester Email", "Flagged Deleted in FootPrints", "Ticket Internal ID", "Has Attachments",
    "Attachment Count", "Attachments Not Exported", "Total Attachment Size Recorded (MB)",
    "Attachment File Names", "Attachment Locations (same order)", "History Entry Count",
    "Full Description History", "Latest Description Entry", "Possible Credential",
]
EXCEL_CELL_LIMIT = 32767
PART_SIZE = 32000  # a little under the cell limit
LONG_COLS = ["Full Description History", "Latest Description Entry"]
ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")  # openpyxl rejects these


def clean_cell(v):
    """Strip control characters Excel cannot store and cut text at the Excel cell limit."""
    if isinstance(v, str):
        v = ILLEGAL_XML.sub("", v)
        if len(v) > EXCEL_CELL_LIMIT:
            return v[:EXCEL_CELL_LIMIT]
    return v


def main():
    load_dotenv(os.path.join(REPO, "config", "db.env"))
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--server", default=os.getenv("FOOTPRINTS_DB_SERVER"))
    ap.add_argument("--database", default=os.getenv("FOOTPRINTS_DB_DATABASE"))
    ap.add_argument("--query", default=DEFAULT_QUERY)
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--check", action="store_true", help="run the query and report only; write nothing")
    args = ap.parse_args()

    if not args.server or not args.database:
        sys.exit("Set FOOTPRINTS_DB_SERVER and FOOTPRINTS_DB_DATABASE in config/db.env "
                 "(or pass --server / --database).")

    print(f"Connecting to {args.server} / {args.database} (Windows Auth) ...")
    conn = get_db_connection_from_env(args.server, args.database, trusted_connection=True)
    try:
        print(f"Running: {args.query}")
        with warnings.catch_warnings():  # pandas prefers SQLAlchemy; a pyodbc connection works fine
            warnings.simplefilter("ignore", UserWarning)
            df = pd.read_sql(args.query, conn)
    finally:
        conn.close()
    print(f"Fetched {len(df):,} rows x {df.shape[1]} columns")

    missing = [c for c in EXPECTED if c not in df.columns]
    extra = [c for c in df.columns if c not in EXPECTED]
    if missing:
        print(f"[warn] columns the compare expects but the query did not return: {missing}")
    if extra:
        print(f"[info] extra columns (kept, ignored by the compare): {extra}")

    key = "Ticket Number"
    if key in df.columns:
        dups = df[key].duplicated().sum()
        print(f"Ticket range: {df[key].min()} .. {df[key].max()}   duplicate ticket numbers: {dups}")

    # Long history text does not fit one Excel cell: continue it in "<col> (part 2)", "(part 3)" ...
    # so nothing is lost (compare_footprints.py joins the parts back together).
    for col in LONG_COLS:
        if col not in df.columns:
            continue
        s = df[col].where(df[col].notna(), "").astype(str)
        n_parts = max(1, -(-int(s.str.len().max()) // PART_SIZE))
        if n_parts > 1:
            print(f"[info] {col}: {int((s.str.len() > PART_SIZE).sum())} values longer than {PART_SIZE} chars "
                  f"-> split over {n_parts} columns")
            pos = list(df.columns).index(col) + 1
            for i in range(2, n_parts + 1):
                piece = s.str.slice((i - 1) * PART_SIZE, i * PART_SIZE)
                df.insert(pos + i - 2, f"{col} (part {i})", piece.where(piece != "", None))
            df[col] = s.str.slice(0, PART_SIZE).where(df[col].notna(), None)

    obj_cols = [c for c in df.columns
                if not pd.api.types.is_numeric_dtype(df[c]) and not pd.api.types.is_datetime64_any_dtype(df[c])
                and not pd.api.types.is_bool_dtype(df[c])]
    too_long = {c: int((df[c].astype(str).str.len() > EXCEL_CELL_LIMIT).sum()) for c in obj_cols}
    too_long = {c: n for c, n in too_long.items() if n}
    if too_long:
        print(f"[warn] values longer than Excel's {EXCEL_CELL_LIMIT}-char cell limit were cut: {too_long}")
    for c in obj_cols:
        df[c] = df[c].map(clean_cell)
    # Control-character stripping above can only shorten text, so every part stays within the limit.

    # Excel cannot store timezone-aware datetimes
    for c in df.columns:
        if isinstance(df[c].dtype, pd.DatetimeTZDtype):
            df[c] = df[c].dt.tz_localize(None)

    if args.check:
        print("--check: no file written.")
        return

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    df.to_excel(args.out, index=False, sheet_name="Sheet1")
    print(f"Wrote {args.out}")


if __name__ == "__main__":
    main()
