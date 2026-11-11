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
    attachments (optional): fixtures/QCFootprints_Attachments.xlsx from extract_api_attachments.py
        -> adds UI-vs-DB attachment count checks (sheets Attachment_Count_Check, Attachments_Not_Fetched);
           skip with --no-attachments

Rules applied
    * Key = app "Record Number"  <->  DB "Ticket Number"  (IR-nnnnn)
    * DB Sheet1 has spilled description-text rows: rows without a "Ticket Internal ID" are dropped.
    * DB tickets flagged "Deleted in FootPrints" = Yes are excluded from the comparison.
      If such a ticket exists in the app it is listed as DELETED IN DB (information, not a mismatch).
    * DB timestamps are UTC, app timestamps are US Eastern -> DB is converted to Eastern before comparing.
    * Blanks: '', 'NULL', None, NaN are all treated as blank.
    * Tickets in the DB export but not in the app (and vice versa) are reported as MISSING.

Output: tests/QCFootprints/output/qc_footprints_report.xlsx         (full run; same name every time, tracked in git)
        tests/QCFootprints/output/qc_footprints_report_sample.xlsx  (--sample run; not tracked)
"""
import argparse
import glob
import os
import re
import sys
import time
import unicodedata
from collections import Counter
from urllib.parse import unquote

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
FIXTURES_DIR = os.path.join(HERE, "fixtures")
OUTPUT_DIR = os.path.join(HERE, "output")
DEFAULT_DB_FILE = os.path.join(FIXTURES_DIR, "QCFootprint_DB_Details.xlsx")
DEFAULT_ATT_FILE = os.path.join(FIXTURES_DIR, "QCFootprints_Attachments.xlsx")

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
        sys.exit("No app extract found in fixtures/ (QCFootprints_<date>.xlsx). Run extract_api_tickets.py first.")
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
    # All columns (A-AD): the attachment names/paths, history text and 'Possible Credential' are in U-AD.
    raw = pd.read_excel(path, sheet_name="Sheet1", dtype=object)
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
    # extract_db_footprints.py continues text longer than one Excel cell in "<col> (part N)" columns
    for base in (DB_HISTORY, "Latest Description Entry"):
        parts = sorted((c for c in db.columns if re.fullmatch(re.escape(base) + r" \(part \d+\)", str(c))),
                       key=lambda c: int(re.search(r"part (\d+)", c).group(1)))
        if parts and base in db.columns:
            joined = db[base].where(db[base].notna(), "").astype(str)
            for c in parts:
                joined = joined + db[c].where(db[c].notna(), "").astype(str)
            db[base] = joined.where(joined != "", None)
            print(f"  joined {len(parts)} continuation column(s) of '{base}'")
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
        # Compare numerically when both sides are numbers; otherwise (blank or free text such as
        # 'test' or '31594.74 & 31594.83') compare the cleaned text.
        a, b = pd.to_numeric(app_s, errors="coerce"), pd.to_numeric(db_s, errors="coerce")
        a_t, b_t = norm_text(app_s), norm_text(db_s)
        both_num = a.notna() & b.notna()
        diff = a_t != b_t
        diff[both_num] = ((a - b).abs() > 1e-6)[both_num]
        a_out = a.astype(object).where(a.notna(), a_t)
        b_out = b.astype(object).where(b.notna(), b_t)
        return diff, a_out, b_out

    if kind == "datetime":
        a = pd.to_datetime(app_s, errors="coerce", format="ISO8601")  # app: naive, US Eastern
        b = pd.to_datetime(db_s, errors="coerce")  # DB: naive, UTC
        b = b.dt.tz_localize("UTC").dt.tz_convert(APP_TZ).dt.tz_localize(None)
        bad = (~is_blank_raw(app_s) & a.isna()) | (~is_blank_raw(db_s) & b.isna())
        diff = (a.isna() ^ b.isna()) | ((a - b).abs() > pd.Timedelta(seconds=tol_seconds)).fillna(False)
        return diff | bad, a, b

    raise ValueError(kind)


# ---------------------------------------------------------------- mismatch categories
CAT_STALE = "CHANGED AFTER DB EXPORT"
CAT_USER_ID = "USER ID INSTEAD OF NAME (DB)"
CAT_ACCENTS = "ACCENTS / SPECIAL CHARACTERS"
CAT_QUESTION = "NON-ASCII REPLACED BY '?' (DB)"
CAT_QUOTES = "QUOTES REMOVED"
CAT_DB_BLANK = "DB BLANK, APP HAS VALUE"
CAT_OTHER = "OTHER"

_LETTER_MAP = str.maketrans({"Ł": "L", "ł": "l", "Ø": "O", "ø": "o", "Đ": "D", "đ": "d", "ß": "ss", "Æ": "AE", "æ": "ae"})
_PUNCT_MAP = str.maketrans({
    "‐": "-", "‑": "-", "‒": "-", "–": "-", "—": "-", "−": "-",
    "‘": "'", "’": "'", "“": '"', "”": '"',
    " ": " ", " ": " ", " ": " ",
})
_QUOTE_CHARS = "\"'‘’“”"


def ascii_fold(s):
    s = str(s).translate(_LETTER_MAP).translate(_PUNCT_MAP)
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def categorize(app_val, db_val):
    a = "" if pd.isna(app_val) else str(app_val)
    d = "" if pd.isna(db_val) else str(db_val)
    if re.fullmatch(r"-?\d+", d.strip()) and not re.fullmatch(r"-?\d+", a.strip()):
        return CAT_USER_ID
    if ascii_fold(a).strip() == d.strip():
        return CAT_ACCENTS
    if re.sub(r"[^\x00-\x7f]+", "?", a).strip() == d.strip():
        return CAT_QUESTION
    strip_q = lambda t: re.sub("[" + re.escape(_QUOTE_CHARS) + "]", "", ascii_fold(t)).strip()
    if strip_q(a) == strip_q(d):
        return CAT_QUOTES
    if not d.strip() and a.strip():
        return CAT_DB_BLANK
    return CAT_OTHER


def ticket_number(key):
    m = re.search(r"(\d+)", str(key))
    return int(m.group(1)) if m else -1


# ---------------------------------------------------------------- compare
def compare(app, db_all, sample, tol_seconds):
    deleted_mask = db_all[DB_DELETED_FLAG].astype(str).str.strip().str.lower() == "yes"
    db_deleted, db = db_all[deleted_mask], db_all[~deleted_mask]
    print(f"  DB tickets flagged deleted (excluded): {len(db_deleted)} -> {len(db)} live tickets")

    # When was the DB exported? Latest timestamp in it (UTC -> Eastern, to match the app).
    stamps = pd.concat([pd.to_datetime(db_all["Ticket Created Date"], errors="coerce"),
                        pd.to_datetime(db_all["Last Updated Date"], errors="coerce")])
    export_cutoff = stamps.max().tz_localize("UTC").tz_convert(APP_TZ).tz_localize(None)
    db_max_number = max(ticket_number(k) for k in db_all[KEY_DB].astype(str))
    print(f"  DB export cutoff (latest timestamp in DB, Eastern): {export_cutoff};  highest ticket: IR-{db_max_number}")

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

    # Tickets whose app "Updated On" is later than the DB export were changed after the export:
    # any difference on them is expected, so they are labelled separately.
    app_updated = pd.to_datetime(app_c["Updated On"], errors="coerce", format="ISO8601")
    stale_keys = set(app_updated[app_updated > export_cutoff].index)

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
            cat = CAT_STALE if key in stale_keys else (
                categorize(a_raw.loc[key], d_raw.loc[key]) if kind in ("text", "email") else CAT_OTHER)
            mism_rows.append({
                "Record Number": key, "Category": cat, "App column": app_col, "DB column": db_col,
                "App value": a_raw.loc[key], "DB value": d_raw.loc[key],
                "App normalized": a_n.loc[key], "DB normalized": d_n.loc[key],
            })

    mism_df = pd.DataFrame(mism_rows)
    return {
        "export_cutoff": export_cutoff, "db_max_number": db_max_number, "stale_tickets": len(stale_keys),
        "missing_in_db": missing_in_db, "missing_in_app": missing_in_app,
        "deleted_in_app": deleted_in_app, "compared": len(both),
        "mismatches": mism_df, "col_stats": pd.DataFrame(col_stats),
        "n_app": len(app_u), "n_db_live": len(db_u), "n_db_deleted": len(db_deleted),
        "db_live": db_u, "stale_keys": stale_keys,
    }


# ---------------------------------------------------------------- attachment names / size / paths
DB_NAMES = "Attachment File Names"
DB_PATHS = "Attachment Locations (same order)"
DB_SIZE = "Total Attachment Size Recorded (MB)"
DB_HISTORY = "Full Description History"
DB_CRED = "Possible Credential"
LIST_SEP = " | "
# Expected DB path: D:\FPextract\out\item\attachment\<n>\<n>\<ticket internal ID>\<file>  (compared with / for \)
PATH_RX = re.compile(r"^D:/FPextract/out/item/attachment/\d+/\d+/(\d+)/([^/]+)$")
ILLEGAL_FILE_CHARS = re.compile(r'[\\/:*?"<>|]')
SIZE_TOLERANCE_MB = 0.011  # DB value is rounded to 2 decimals

CAT_NAME_DIFF = "NAME DIFFERENCE"
CAT_SIZE_DIFF = "SIZE DIFFERENCE"
CAT_P_SAME = "SAME NAME"
CAT_P_PREFIX = "NUMERIC PREFIX ADDED"
CAT_P_ILLEGAL = "ILLEGAL CHARACTERS REPLACED BY _"
CAT_P_SHORT = "NAME SHORTENED WITH ~"
CAT_P_SPECIAL = "SPECIAL CHARACTERS SIMPLIFIED"
CAT_P_OTHER = "OTHER"


def split_list(value):
    return [x for x in str(value).split(LIST_SEP) if x] if pd.notna(value) else []


def path_name_category(att_name, file_name):
    """How the file name inside a DB path differs from the attachment's own name."""
    core, prefixed = file_name, False
    if file_name != att_name and file_name != ILLEGAL_FILE_CHARS.sub("_", att_name):
        m = re.match(r"\d+_(.*)$", file_name, re.S)
        if m:
            core, prefixed = m.group(1), True
    if core == att_name:
        base = CAT_P_SAME
    elif core == ILLEGAL_FILE_CHARS.sub("_", att_name):
        base = CAT_P_ILLEGAL
    elif "~" in core:
        base = CAT_P_SHORT
    elif ascii_fold(core) == ascii_fold(att_name):
        base = CAT_P_SPECIAL
    else:
        base = CAT_P_OTHER
    if prefixed:
        return CAT_P_PREFIX if base == CAT_P_SAME else base + " + NUMERIC PREFIX"
    return base


def check_attachment_details(att_path, res, db, checked):
    """Attachment names, total size and D-drive paths: UI (attachment file) vs the DB's columns U-W.

    DB names are URL-decoded first (the export holds some as %3A, %5E, ...). A ticket whose app
    Updated On is after the DB export is labelled CHANGED AFTER DB EXPORT. Paths cannot be compared
    with the UI (the app only has an internal URL), so they are checked for the expected pattern and
    the DB path file name is compared with the attachment name.
    """
    ui_att = pd.read_excel(att_path, sheet_name="Attachments", dtype=object)
    ui_att["_key"] = ui_att["Record Number"].astype(str).str.strip()
    ui_by = {k: g for k, g in ui_att.groupby("_key")}
    stale = res["stale_keys"]
    db_count = pd.to_numeric(db["Attachment Count"], errors="coerce").fillna(0).astype(int)

    name_rows, size_rows, path_rows = [], [], []
    path_cat, n_files, n_path_bad, n_tickets = {}, 0, 0, 0
    for k in checked:
        if db_count.loc[k] == 0:
            continue
        n_tickets += 1
        g = ui_by.get(k)
        ui_names = [str(x) for x in g["File Name"]] if g is not None else []
        db_names = [unquote(x) for x in split_list(db.loc[k, DB_NAMES])]

        if sorted(db_names) != sorted(ui_names):
            only_db = list((Counter(db_names) - Counter(ui_names)).elements())
            only_ui = list((Counter(ui_names) - Counter(db_names)).elements())
            name_rows.append({"Record Number": k,
                              "Category": CAT_ATT_STALE if k in stale else CAT_NAME_DIFF,
                              "In DB only": "; ".join(only_db)[:EXCEL_CELL_LIMIT],
                              "In UI only": "; ".join(only_ui)[:EXCEL_CELL_LIMIT]})

        ui_mib = (pd.to_numeric(g["Size (bytes)"], errors="coerce").sum() / 1048576) if g is not None else 0.0
        db_mb = pd.to_numeric(db.loc[k, DB_SIZE], errors="coerce")
        if pd.isna(db_mb) or abs(db_mb - ui_mib) > SIZE_TOLERANCE_MB:
            size_rows.append({"Record Number": k,
                              "Category": CAT_ATT_STALE if k in stale else CAT_SIZE_DIFF,
                              "DB size (MB)": db_mb, "UI size (MiB)": round(ui_mib, 2)})

        paths = split_list(db.loc[k, DB_PATHS])
        raw_names = split_list(db.loc[k, DB_NAMES])
        internal_id = str(int(float(db.loc[k, DB_INTERNAL_ID])))
        if len(paths) != len(raw_names):
            n_path_bad += 1
            path_rows.append({"Record Number": k, "Category": "PATH COUNT != NAME COUNT",
                              "Attachment name": "", "Path file name": "", "DB path": ""})
            continue
        for name_raw, p in zip(raw_names, paths):
            n_files += 1
            m = PATH_RX.match(p.replace(chr(92), "/"))
            if not m:
                n_path_bad += 1
                path_rows.append({"Record Number": k, "Category": "UNEXPECTED PATH PATTERN",
                                  "Attachment name": unquote(name_raw), "Path file name": "", "DB path": p})
                continue
            if m.group(1) != internal_id:
                n_path_bad += 1
                path_rows.append({"Record Number": k, "Category": "FOLDER ID != TICKET INTERNAL ID",
                                  "Attachment name": unquote(name_raw), "Path file name": m.group(2), "DB path": p})
                continue
            cat = path_name_category(unquote(name_raw), m.group(2))
            path_cat[cat] = path_cat.get(cat, 0) + 1
            if cat != CAT_P_SAME:
                path_rows.append({"Record Number": k, "Category": cat, "Attachment name": unquote(name_raw),
                                  "Path file name": m.group(2), "DB path": p})

    empty = lambda cols: pd.DataFrame(columns=cols)
    return {
        "tickets": n_tickets, "files": n_files, "path_violations": n_path_bad, "path_cat": path_cat,
        "names": pd.DataFrame(name_rows) if name_rows else empty(["Record Number", "Category", "In DB only", "In UI only"]),
        "sizes": pd.DataFrame(size_rows) if size_rows else empty(["Record Number", "Category", "DB size (MB)", "UI size (MiB)"]),
        "paths": pd.DataFrame(path_rows) if path_rows else
        empty(["Record Number", "Category", "Attachment name", "Path file name", "DB path"]),
    }


# ---------------------------------------------------------------- possible credentials
# A password-like word followed by a value, e.g. "password: Abc123", "pw is Xy7!q".
CRED_WORD_RX = re.compile(r"(?i)\b(?:passw\w*|pwd|pw|passcode)\b")
CRED_VALUE_RX = re.compile(r"(?i)\b(?:passw(?:ord|d)s?|pwd|pw|passcode)\b\s*(?:(?:is|are|was|=|:|-|->)\s*[:=]?\s*)(\S+)")
CRED_STOP = {"to", "for", "and", "the", "reset", "change", "changed", "is", "has", "will", "please", "you", "your",
             "in", "on", "a", "of", "was", "expired", "incorrect", "protected", "manager", "required", "needed",
             "request", "link", "below", "above", "attached", "same", "it", "this", "that", "or", "not", "with"}
CRED_PREFIX = 40  # characters of context shown before the password word


def looks_like_password(token):
    t = token.strip("\"'.,;:)(]>[<“”‘’")
    low = t.lower()
    if len(t) < 4 or low in CRED_STOP or low.startswith(("http", "www.")) or "@" in t:
        return False
    has_digit = any(c.isdigit() for c in t)
    has_symbol = any(not c.isalnum() for c in t)
    inner_mixed = any(c.isupper() for c in t[1:]) and any(c.islower() for c in t)
    return has_digit or has_symbol or inner_mixed


def check_credentials(res):
    """List tickets flagged 'Possible Credential = Yes' and judge whether the text holds a real password.

    The report is public in the repo, so values are never written: a likely password is replaced by
    [REDACTED] and a plain mention shows only the words before the term.
    """
    db = res["db_live"]
    if DB_CRED not in db.columns:
        print("  (DB file has no 'Possible Credential' column - credential review skipped)")
        return None
    flagged = db[db[DB_CRED].astype(str).str.strip().str.lower() == "yes"]
    rows = []
    for _, r in flagged.iterrows():
        text = re.sub(r"\s+", " ", str(r[DB_HISTORY])) if pd.notna(r.get(DB_HISTORY)) else ""
        hits = [m for m in CRED_VALUE_RX.finditer(text) if looks_like_password(m.group(1))]
        # A meeting passcode is not a password typed into the ticket: "...zoom.us/j/123?pwd=abc" in a URL,
        # or "Meeting ID: 123 | Passcode: abc".
        is_meeting = lambda m: (m.start() > 0 and text[m.start() - 1] in "?&") or             "meeting id" in text[max(0, m.start() - 80):m.start()].lower()
        in_url = [m for m in hits if is_meeting(m)]
        likely = [m for m in hits if m not in in_url]
        any_word = CRED_WORD_RX.search(text)
        if likely:
            m = likely[0]
            ctx = text[max(0, m.start() - CRED_PREFIX):m.start(1)].strip() + " [REDACTED]"
            verdict = "LIKELY REAL PASSWORD - REVIEW"
        elif in_url:
            m = in_url[0]
            ctx = text[max(0, m.start() - CRED_PREFIX):m.start()].strip() + m.group(0)[:len(m.group(0)) - len(m.group(1))] + " [REDACTED]"
            verdict = "MEETING PASSCODE (Zoom link / Meeting ID)"
        elif any_word:
            ctx = text[max(0, any_word.start() - CRED_PREFIX):any_word.end()].strip()
            verdict = "MENTIONS PASSWORD WORD ONLY"
        else:
            ctx, verdict = "", "NO PASSWORD WORD FOUND"
        rows.append({"Record Number": r["_key"], "Possible Credential": "Yes", "Assessment": verdict,
                     "Likely password matches": len(likely), "Masked context": ctx})
    order = {"LIKELY REAL PASSWORD - REVIEW": 0, "MEETING PASSCODE (Zoom link / Meeting ID)": 1,
             "MENTIONS PASSWORD WORD ONLY": 2, "NO PASSWORD WORD FOUND": 3}
    cols = ["Record Number", "Possible Credential", "Assessment", "Likely password matches", "Masked context"]
    df = pd.DataFrame(rows, columns=cols)
    if len(df):
        df = df.sort_values("Assessment", key=lambda s: s.map(order), kind="stable").reset_index(drop=True)
    return {"flagged": len(flagged), "table": df,
            "counts": df["Assessment"].value_counts().to_dict() if len(df) else {}}


# ---------------------------------------------------------------- attachments
CAT_ATT_STALE = "CHANGED AFTER DB EXPORT"
CAT_ATT_MORE = "MORE ATTACHMENTS IN UI THAN DB"
CAT_ATT_FEWER = "FEWER ATTACHMENTS IN UI THAN DB"


def check_attachments(att_path, res):
    """UI attachment counts (from extract_api_attachments.py) vs the DB's Attachment Count.

    Only tickets present in the attachment file are checked. The extract normally holds every ticket the
    DB says has attachments plus a random sample of tickets it says have none.
    """
    print(f"Attachments file: {att_path}")
    ui = pd.read_excel(att_path, sheet_name="Ticket_Counts", dtype=object)
    ui["_key"] = ui["Record Number"].astype(str).str.strip()
    ui["UI Attachment Count"] = pd.to_numeric(ui["UI Attachment Count"], errors="coerce").fillna(0).astype(int)
    ui = ui.drop_duplicates("_key").set_index("_key")

    db = res["db_live"].set_index("_key")
    db_count = pd.to_numeric(db["Attachment Count"], errors="coerce").fillna(0).astype(int)

    checked = sorted(set(ui.index) & set(db.index))
    not_in_db = sorted(set(ui.index) - set(db.index))
    todo = sorted(set(db_count.index[db_count > 0]) - set(ui.index))  # DB says attachments, not fetched yet

    cmp = pd.DataFrame({"DB Attachment Count": db_count.loc[checked],
                        "UI Attachment Count": ui.loc[checked, "UI Attachment Count"]})
    cmp["Difference (UI - DB)"] = cmp["UI Attachment Count"] - cmp["DB Attachment Count"]
    bad = cmp[cmp["Difference (UI - DB)"] != 0].copy()
    bad["Category"] = [CAT_ATT_STALE if k in res["stale_keys"] else
                       (CAT_ATT_MORE if d > 0 else CAT_ATT_FEWER)
                       for k, d in zip(bad.index, bad["Difference (UI - DB)"])]
    bad = bad.reset_index().rename(columns={"_key": "Record Number", "index": "Record Number"})
    bad = bad[["Record Number", "Category", "DB Attachment Count", "UI Attachment Count", "Difference (UI - DB)"]]

    none_sample = cmp[cmp["DB Attachment Count"] == 0]
    with_att = cmp[cmp["DB Attachment Count"] > 0]
    details = None
    if all(c in db.columns for c in (DB_NAMES, DB_PATHS, DB_SIZE)):
        details = check_attachment_details(att_path, res, db, checked)
    else:
        print("  (DB file has no attachment name/path/size columns - name, path and size checks skipped)")
    return {
        "details": details,
        "file": os.path.basename(att_path), "checked": len(checked), "not_in_db": not_in_db,
        "to_fetch": todo, "mismatches": bad,
        "with_att_checked": len(with_att), "with_att_ok": int((with_att["Difference (UI - DB)"] == 0).sum()),
        "none_checked": len(none_sample),
        "none_with_ui_att": int((none_sample["UI Attachment Count"] > 0).sum()),
        "db_total": int(cmp["DB Attachment Count"].sum()), "ui_total": int(cmp["UI Attachment Count"].sum()),
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
    # One fixed file name, overwritten on every run, so the repo only ever holds the latest report.
    # Sample runs use a different name so they never overwrite the full report.
    out = os.path.join(OUTPUT_DIR, "qc_footprints_report_sample.xlsx" if sample else "qc_footprints_report.xlsx")
    mm = res["mismatches"]
    mm_cats = (mm.groupby(["Category", "App column"]).size().rename("Cells").reset_index()
               .sort_values(["Category", "Cells"], ascending=[True, False])) if len(mm) else \
        pd.DataFrame(columns=["Category", "App column", "Cells"])
    missing_new = [k for k in res["missing_in_db"] if ticket_number(k) > res["db_max_number"]]
    missing_old = [k for k in res["missing_in_db"] if ticket_number(k) <= res["db_max_number"]]
    missing_df = pd.DataFrame({
        "Record Number": res["missing_in_db"],
        "Category": ["NEWER THAN DB EXPORT" if ticket_number(k) > res["db_max_number"]
                     else "WITHIN DB RANGE - INVESTIGATE" for k in res["missing_in_db"]]})
    summary = pd.DataFrame([
        ("Mode", "SAMPLE (only tickets in the app file)" if sample else "FULL"),
        ("App file", os.path.basename(app_path)), ("DB file", os.path.basename(db_path)),
        ("DB export cutoff (Eastern)", str(res["export_cutoff"])),
        ("App tickets", res["n_app"]), ("DB live tickets (not deleted)", res["n_db_live"]),
        ("DB tickets flagged deleted (excluded)", res["n_db_deleted"]),
        ("Compared (in both)", res["compared"]),
        ("Tickets changed in app after DB export", res["stale_tickets"]),
        ("MISSING IN DB (in app, not in DB)", len(res["missing_in_db"])),
        ("   of which newer than the DB export", len(missing_new)),
        ("   of which inside the DB's range (investigate)", len(missing_old)),
        ("MISSING IN APP (in DB, not in app)", "n/a (sample mode)" if sample else len(res["missing_in_app"])),
        ("DELETED IN DB but present in app", len(res["deleted_in_app"])),
        ("Value mismatches (cells)", len(res["mismatches"])),
        ("Tickets with a mismatch", res["mismatches"]["Record Number"].nunique() if len(res["mismatches"]) else 0),
    ], columns=["Item", "Value"])
    att = res.get("attach")
    if att:
        mm_att = att["mismatches"]
        summary = pd.concat([summary, pd.DataFrame([
            ("--- ATTACHMENTS ---", ""),
            ("Attachment file", att["file"]),
            ("Tickets checked (UI vs DB count)", att["checked"]),
            ("   DB says attachments: checked / count matches", f"{att['with_att_checked']} / {att['with_att_ok']}"),
            ("   DB says none (sample): checked", att["none_checked"]),
            ("   DB says none but UI has attachments", att["none_with_ui_att"]),
            ("Total attachments, DB vs UI (checked tickets)", f"{att['db_total']} vs {att['ui_total']}"),
            ("Attachment count mismatches (tickets)", len(mm_att)),
            ("   of which changed after DB export", int((mm_att["Category"] == CAT_ATT_STALE).sum())),
            ("DB says attachments but not yet fetched", len(att["to_fetch"])),
            ("In attachment file but not a live DB ticket", len(att["not_in_db"])),
        ], columns=["Item", "Value"])], ignore_index=True)
        d = att.get("details")
        if d:
            pc = d["path_cat"]
            n_same = pc.get(CAT_P_SAME, 0)
            n_diff = sum(v for k, v in pc.items() if k != CAT_P_SAME)
            rows = [
                ("--- ATTACHMENT NAMES / SIZE / D-DRIVE PATHS (DB columns U-W) ---", ""),
                ("Tickets checked (DB says attachments)", d["tickets"]),
                ("Name differences UI vs DB (tickets)", len(d["names"])),
                ("   of which changed after DB export", int((d["names"]["Category"] == CAT_ATT_STALE).sum())),
                ("Total-size differences UI vs DB (tickets)", len(d["sizes"])),
                ("   of which changed after DB export", int((d["sizes"]["Category"] == CAT_ATT_STALE).sum())),
                ("DB paths checked (files)", d["files"]),
                ("Path problems (bad pattern / wrong folder ID / count)", d["path_violations"]),
                ("Path file name identical to attachment name", n_same),
                ("Path file name differs from attachment name", n_diff),
            ]
            rows += [(f"   {k}", v) for k, v in sorted(pc.items(), key=lambda kv: -kv[1]) if k != CAT_P_SAME]
            summary = pd.concat([summary, pd.DataFrame(rows, columns=["Item", "Value"])], ignore_index=True)
    cred = res.get("cred")
    if cred:
        rows = [("--- POSSIBLE CREDENTIAL = YES (masked review list) ---", ""),
                ("Tickets flagged Possible Credential = Yes", cred["flagged"])]
        rows += [(f"   {k}", v) for k, v in cred["counts"].items()]
        summary = pd.concat([summary, pd.DataFrame(rows, columns=["Item", "Value"])], ignore_index=True)
    try:
        with pd.ExcelWriter(out, engine="openpyxl") as xw:
            summary.to_excel(xw, sheet_name="Summary", index=False)
            res["col_stats"].to_excel(xw, sheet_name="Column_Check", index=False)
            mm_cats.to_excel(xw, sheet_name="Mismatch_Categories", index=False)
            cap(res["mismatches"]).to_excel(xw, sheet_name="Value_Mismatches", index=False)
            missing_df.to_excel(xw, sheet_name="Missing_in_DB", index=False)
            pd.DataFrame({"Record Number": res["missing_in_app"]}).to_excel(xw, sheet_name="Missing_in_App", index=False)
            pd.DataFrame({"Record Number": res["deleted_in_app"]}).to_excel(xw, sheet_name="Deleted_in_DB", index=False)
            if att:
                att["mismatches"].to_excel(xw, sheet_name="Attachment_Count_Check", index=False)
                pd.DataFrame({"Record Number": att["to_fetch"]}).to_excel(xw, sheet_name="Attachments_Not_Fetched", index=False)
                d = att.get("details")
                if d:
                    d["names"].to_excel(xw, sheet_name="Attachment_Name_Check", index=False)
                    d["sizes"].to_excel(xw, sheet_name="Attachment_Size_Check", index=False)
                    cap(d["paths"]).to_excel(xw, sheet_name="Attachment_Path_Check", index=False)
            if cred:
                cap(cred["table"]).to_excel(xw, sheet_name="Credential_Review", index=False)
    except PermissionError:
        sys.exit(f"Cannot write {out} - close it in Excel and run again.")
    return out, summary


def print_console(res, sample, summary):
    print("\n==== QC FOOTPRINTS COMPARISON ====")
    for item, value in summary.itertuples(index=False):
        print(f"{item:<42}: {value}")
    print("\n---- per-column check ----")
    print(res["col_stats"].to_string(index=False))
    mm = res["mismatches"]
    if len(mm):
        print("\n---- mismatches by category ----")
        print(mm.groupby("Category").agg(Cells=("Record Number", "size"),
                                         Tickets=("Record Number", "nunique")).to_string())
        print("\n---- sample mismatches (up to 3 per category and column) ----")
        for (cat, col), grp in mm.groupby(["Category", "App column"], sort=False):
            print(f"\n[{cat}] {col}: {len(grp)} cells")
            for r in grp.head(3).itertuples(index=False):
                print(f"  {r[0]}: app={str(r[4])[:60]!r}  db={str(r[5])[:60]!r}")
    if res["missing_in_db"]:
        print(f"\nMISSING IN DB (first 10): {res['missing_in_db'][:10]}")
    if res["missing_in_app"]:
        print(f"MISSING IN APP (first 10): {res['missing_in_app'][:10]}")
    if res["deleted_in_app"]:
        print(f"DELETED IN DB but in app (first 10): {res['deleted_in_app'][:10]}")


def main():
    try:  # Windows consoles (cp1252) cannot print every character found in ticket text
        sys.stdout.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--app", help="app extract Excel (default: newest QCFootprints_<date>.xlsx)")
    p.add_argument("--db", default=DEFAULT_DB_FILE, help="DB export Excel")
    p.add_argument("--sample", action="store_true",
                   help="only compare tickets present in the app file; skip MISSING IN APP")
    p.add_argument("--tolerance", type=float, default=1.0, help="datetime tolerance in seconds (default 1)")
    p.add_argument("--attachments", default=DEFAULT_ATT_FILE,
                   help="attachment extract from extract_api_attachments.py (default: fixtures/QCFootprints_Attachments.xlsx)")
    p.add_argument("--no-attachments", action="store_true", help="skip the attachment count checks")
    args = p.parse_args()

    app_path = args.app or newest_app_file()
    app = load_app(app_path)
    if not args.sample and len(app) < 5000:
        print(f"[warn] app file has only {len(app)} rows. Without --sample every other DB ticket "
              "will be reported as MISSING IN APP.")
    db = clean_db(load_db(args.db))

    res = compare(app, db, args.sample, args.tolerance)
    if not args.no_attachments:
        if os.path.exists(args.attachments):
            res["attach"] = check_attachments(args.attachments, res)
        else:
            print(f"(no attachment file at {args.attachments} - attachment checks skipped)")
    res["cred"] = check_credentials(res)
    out, summary = write_report(res, args.sample, app_path, args.db)
    print_console(res, args.sample, summary)
    print(f"\nReport written to: {out}")


if __name__ == "__main__":
    main()
