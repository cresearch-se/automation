"""
Compare the Footprints app extract (Excel) against the DB export (Excel).

Run on the LOCAL machine from the repo root:

    # 25-row sanity check: only compares tickets that are in the app file
    python tests/QCFootprints/compare_footprints.py --sample

    # full comparison (after the full extract)
    python tests/QCFootprints/compare_footprints.py

Inputs
    app file : newest tests/QCFootprints/fixtures/QCFootprints_<date>.xlsx   (or --app PATH)
    DB file  : tests/QCFootprints/fixtures/QCFootprint_DB_Details.xlsx       (or --db PATH)

Rules applied
    * Key = app "Record Number"  <->  DB "Ticket Number"  (IR-nnnnn)
    * DB Sheet1 has spilled description-text rows: rows without a "Ticket Internal ID" are dropped.
    * DB tickets flagged "Deleted in FootPrints" = Yes are excluded from the comparison.
      If such a ticket exists in the app it is listed as DELETED IN DB (information, not a mismatch).
    * DB timestamps are UTC, app timestamps are US Eastern -> DB is converted to Eastern before comparing.
    * Blanks: '', 'NULL', None, NaN are all treated as blank.
    * Tickets in the DB export but not in the app (and vice versa) are reported as MISSING.

Output: tests/QCFootprints/output/qc_footprints_report_<YYYYMMDD_HHMMSS>.xlsx
"""
import argparse
import glob
import os
import sys
import time

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
FIXTURES_DIR = os.path.join(HERE, "fixtures")
OUTPUT_DIR = os.path.join(HERE, "output")
DEFAULT_DB_FILE = os.path.join(FIXTURES_DIR, "QCFootprint_DB_Details.xlsx")

KEY_APP = "Record Number"
KEY_DB = "Ticket Number"
DB_INTERNAL_ID = "Ticket Internal ID"
DB_DELETED_FLAG = "Flagged Deleted in FootPrints"
APP_TZ = "America/New_York"

# (DB column, App column, kind)   kind: text | email | number | datetime
MAPPING = [
    ("Ticket Internal ID", "ID", "number"),
    ("Ticket Title", "Title", "text"),
    ("Case Name", "Client Matter Name", "text"),
    ("Case / Matter Number", "Client Matter Number", "number"),
    ("Status", "Status", "text"),
    ("Priority", "Priority", "text"),
    ("Responsible Officer", "Responsible Officer", "text"),
    ("Officer Location", "Officer Location", "text"),
    ("Ticket Created Date", "Created On", "datetime"),
    ("Created By", "Created By", "text"),
    ("Last Updated Date", "Updated On", "datetime"),
    ("Last Updated By", "Updated By", "text"),
    ("Requester First Name", "First Name", "text"),
    ("Requester Last Name", "Last Name", "text"),
    ("Requester Email", "Email Address", "email"),
]

BLANK_TEXT = {"", "null", "none", "nan", "nat"}
EXCEL_CELL_LIMIT = 32767


# ---------------------------------------------------------------- loading
def newest_app_file():
    files = sorted(glob.glob(os.path.join(FIXTURES_DIR, "QCFootprints_2*.xlsx")),
                   key=os.path.getmtime)
    if not files:
        sys.exit("No app extract found in fixtures/ (QCFootprints_<date>.xlsx). Run extract_footprints.py first.")
    return files[-1]


def load_app(path):
    print(f"App file : {path}")
    df = pd.read_excel(path, sheet_name=0, dtype=object)
    missing = [a for _, a, _ in MAPPING + [(None, KEY_APP, None)] if a not in df.columns]
    if missing:
        sys.exit(f"App file is missing columns {missing}.\nColumns present: {list(df.columns)}")
    df["_key"] = df[KEY_APP].astype(str).str.strip()
    return df


def load_db(path):
    """Read DB Sheet1 (slow for 50 MB), cached as a pickle next to the output."""
    print(f"DB file  : {path}")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    cache = os.path.join(OUTPUT_DIR, "db_clean_cache.pkl")
    stamp = (os.path.getmtime(path), os.path.getsize(path))
    if os.path.exists(cache):
        cached = pd.read_pickle(cache)
        if cached.get("stamp") == stamp:
            print("  (using cached DB read)")
            return cached["raw"]
    print("  reading DB Excel (first time takes a few minutes)...")
    raw = pd.read_excel(path, sheet_name="Sheet1", usecols=range(20), dtype=object)
    pd.to_pickle({"stamp": stamp, "raw": raw}, cache)
    return raw


def clean_db(raw):
    missing = [d for d, _, _ in MAPPING + [(KEY_DB, None, None)] if d not in raw.columns]
    if missing:
        sys.exit(f"DB file is missing columns {missing}.\nColumns present: {list(raw.columns)}")
    n_all = len(raw)
    db = raw[raw[DB_INTERNAL_ID].notna()].copy()  # drops spilled description-text rows
    print(f"  DB rows: {n_all} non-empty, {n_all - len(db)} spilled-text rows dropped, {len(db)} real tickets")
    db["_key"] = db[KEY_DB].astype(str).str.strip()
    return db


# ---------------------------------------------------------------- normalizers
def is_blank_raw(s):
    t = s.astype(object).where(s.notna(), "")
    return t.astype(str).str.strip().str.lower().isin(BLANK_TEXT)


def norm_text(s):
    t = s.astype(object).where(s.notna(), "").astype(str)
    t = t.str.replace(r"\s+", " ", regex=True).str.strip()
    return t.where(~t.str.lower().isin(BLANK_TEXT), "")


def mismatch_series(app_s, db_s, kind, tol_seconds):
    """Return (boolean mismatch Series, app_normalized, db_normalized)."""
    if kind in ("text", "email"):
        a, b = norm_text(app_s), norm_text(db_s)
        if kind == "email":
            a, b = a.str.lower(), b.str.lower()
        return a != b, a, b

    if kind == "number":
        a, b = pd.to_numeric(app_s, errors="coerce"), pd.to_numeric(db_s, errors="coerce")
        bad = (~is_blank_raw(app_s) & a.isna()) | (~is_blank_raw(db_s) & b.isna())
        diff = (a.isna() ^ b.isna()) | ((a - b).abs() > 1e-6).fillna(False)
        return diff | bad, a, b

    if kind == "datetime":
        a = pd.to_datetime(app_s, errors="coerce", format="ISO8601")  # app: naive, US Eastern
        b = pd.to_datetime(db_s, errors="coerce")  # DB: naive, UTC
        b = b.dt.tz_localize("UTC").dt.tz_convert(APP_TZ).dt.tz_localize(None)
        bad = (~is_blank_raw(app_s) & a.isna()) | (~is_blank_raw(db_s) & b.isna())
        diff = (a.isna() ^ b.isna()) | ((a - b).abs() > pd.Timedelta(seconds=tol_seconds)).fillna(False)
        return diff | bad, a, b

    raise ValueError(kind)


# ---------------------------------------------------------------- compare
def compare(app, db_all, sample, tol_seconds):
    deleted_mask = db_all[DB_DELETED_FLAG].astype(str).str.strip().str.lower() == "yes"
    db_deleted, db = db_all[deleted_mask], db_all[~deleted_mask]
    print(f"  DB tickets flagged deleted (excluded): {len(db_deleted)} -> {len(db)} live tickets")

    app_dups = app["_key"].duplicated(keep=False)
    db_dups = db["_key"].duplicated(keep=False)
    if app_dups.any():
        print(f"[warn] {app_dups.sum()} app rows share a Record Number; keeping the first of each")
    if db_dups.any():
        print(f"[warn] {db_dups.sum()} DB rows share a Ticket Number; keeping the first of each")
    app_u = app.drop_duplicates("_key")
    db_u = db.drop_duplicates("_key")

    app_keys, db_keys, del_keys = set(app_u["_key"]), set(db_u["_key"]), set(db_deleted["_key"])

    deleted_in_app = sorted(app_keys & del_keys)
    missing_in_db = sorted(app_keys - db_keys - del_keys)
    missing_in_app = [] if sample else sorted(db_keys - app_keys)
    both = sorted(app_keys & db_keys)

    app_c = app_u[app_u["_key"].isin(both)].set_index("_key")
    db_c = db_u[db_u["_key"].isin(both)].set_index("_key").loc[app_c.index]

    mism_rows, col_stats = [], []
    for db_col, app_col, kind in MAPPING:
        a_raw, d_raw = app_c[app_col], db_c[db_col]
        bad, a_n, d_n = mismatch_series(a_raw, d_raw, kind, tol_seconds)
        stat = {
            "App column": app_col, "DB column": db_col, "Type": kind,
            "Compared": len(a_raw), "Mismatches": int(bad.sum()),
            "App blank": int(is_blank_raw(a_raw).sum()), "DB blank": int(is_blank_raw(d_raw).sum()),
        }
        col_stats.append(stat)
        for key in bad[bad].index:
            mism_rows.append({
                "Record Number": key, "App column": app_col, "DB column": db_col,
                "App value": a_raw.loc[key], "DB value": d_raw.loc[key],
                "App normalized": a_n.loc[key], "DB normalized": d_n.loc[key],
            })

    return {
        "missing_in_db": missing_in_db, "missing_in_app": missing_in_app,
        "deleted_in_app": deleted_in_app, "compared": len(both),
        "mismatches": pd.DataFrame(mism_rows), "col_stats": pd.DataFrame(col_stats),
        "n_app": len(app_u), "n_db_live": len(db_u), "n_db_deleted": len(db_deleted),
    }


# ---------------------------------------------------------------- report
def cap(df):
    df = df.copy()
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].map(lambda v: v[:EXCEL_CELL_LIMIT] if isinstance(v, str) else v)
    return df


def write_report(res, sample, app_path, db_path):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out = os.path.join(OUTPUT_DIR, f"qc_footprints_report_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
    summary = pd.DataFrame([
        ("Mode", "SAMPLE (only tickets in the app file)" if sample else "FULL"),
        ("App file", os.path.basename(app_path)), ("DB file", os.path.basename(db_path)),
        ("App tickets", res["n_app"]), ("DB live tickets (not deleted)", res["n_db_live"]),
        ("DB tickets flagged deleted (excluded)", res["n_db_deleted"]),
        ("Compared (in both)", res["compared"]),
        ("MISSING IN DB (in app, not in DB)", len(res["missing_in_db"])),
        ("MISSING IN APP (in DB, not in app)", "n/a (sample mode)" if sample else len(res["missing_in_app"])),
        ("DELETED IN DB but present in app", len(res["deleted_in_app"])),
        ("Value mismatches (cells)", len(res["mismatches"])),
        ("Tickets with a mismatch", res["mismatches"]["Record Number"].nunique() if len(res["mismatches"]) else 0),
    ], columns=["Item", "Value"])
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="Summary", index=False)
        res["col_stats"].to_excel(xw, sheet_name="Column_Check", index=False)
        cap(res["mismatches"]).to_excel(xw, sheet_name="Value_Mismatches", index=False)
        pd.DataFrame({"Record Number": res["missing_in_db"]}).to_excel(xw, sheet_name="Missing_in_DB", index=False)
        pd.DataFrame({"Record Number": res["missing_in_app"]}).to_excel(xw, sheet_name="Missing_in_App", index=False)
        pd.DataFrame({"Record Number": res["deleted_in_app"]}).to_excel(xw, sheet_name="Deleted_in_DB", index=False)
    return out, summary


def print_console(res, sample, summary):
    print("\n==== QC FOOTPRINTS COMPARISON ====")
    for item, value in summary.itertuples(index=False):
        print(f"{item:<42}: {value}")
    print("\n---- per-column check ----")
    print(res["col_stats"].to_string(index=False))
    mm = res["mismatches"]
    if len(mm):
        print("\n---- sample mismatches (up to 5 per column) ----")
        for col, grp in mm.groupby("App column", sort=False):
            print(f"\n[{col}]  {len(grp)} mismatches")
            for r in grp.head(5).itertuples(index=False):
                print(f"  {r[0]}: app={str(r[3])[:60]!r}  db={str(r[4])[:60]!r}")
    if res["missing_in_db"]:
        print(f"\nMISSING IN DB (first 10): {res['missing_in_db'][:10]}")
    if res["missing_in_app"]:
        print(f"MISSING IN APP (first 10): {res['missing_in_app'][:10]}")
    if res["deleted_in_app"]:
        print(f"DELETED IN DB but in app (first 10): {res['deleted_in_app'][:10]}")


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--app", help="app extract Excel (default: newest QCFootprints_<date>.xlsx)")
    p.add_argument("--db", default=DEFAULT_DB_FILE, help="DB export Excel")
    p.add_argument("--sample", action="store_true",
                   help="only compare tickets present in the app file; skip MISSING IN APP")
    p.add_argument("--tolerance", type=float, default=1.0, help="datetime tolerance in seconds (default 1)")
    args = p.parse_args()

    app_path = args.app or newest_app_file()
    app = load_app(app_path)
    if not args.sample and len(app) < 5000:
        print(f"[warn] app file has only {len(app)} rows. Without --sample every other DB ticket "
              "will be reported as MISSING IN APP.")
    db = clean_db(load_db(args.db))

    res = compare(app, db, args.sample, args.tolerance)
    out, summary = write_report(res, args.sample, app_path, args.db)
    print_console(res, args.sample, summary)
    print(f"\nReport written to: {out}")


if __name__ == "__main__":
    main()
