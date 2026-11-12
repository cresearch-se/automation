"""
Validate the attachment FILES on the export drive against the UI (app) data and the DB export.

Run from the repo root on a machine that can reach the DB export drive (here: Z: = \\\\25sqlt4cotest\\d$,
the DB paths start with D:\\FPextract\\out\\... and are mapped to <drive>\\FPextract\\out\\...):

    python tests/QCFootprints/check_attachment_files.py                  # full run
    python tests/QCFootprints/check_attachment_files.py --limit 200      # quick trial on 200 tickets
    python tests/QCFootprints/check_attachment_files.py --drive Y:       # other drive letter

Inputs (all from fixtures/, nothing is fetched from the app or the DB again)
    UI   : fixtures/QCFootprints_Attachments.xlsx   (extract_api_attachments.py: name, size in bytes, created)
    DB   : fixtures/QCFootprint_DB_Details.xlsx      (extract_db_footprints.py: names, paths, total size MB)
    Disk : the drive (read-only: folder listings, file sizes, first/last bytes)

Stage 1 - existence and sizes (UI <-> disk <-> DB)
    * every DB path exists on the drive                         -> FILE MISSING ON DRIVE
    * file is not empty                                         -> ZERO-BYTE FILE
    * disk size == UI size (bytes), files paired by name        -> SIZE DIFFERS (UI vs DISK)
    * every UI attachment has a DB path                         -> IN UI BUT NOT IN DB / CREATED AFTER DB EXPORT
    * sum of disk sizes == DB 'Total Attachment Size (MB)'      -> TICKET TOTAL SIZE != DB
    * files in the ticket folder that no DB path points to      -> sheet Attachment_Orphan_Files
Stage 2 - file health (skip with --no-health)
    * header matches the extension (pdf, docx/xlsx/pptx, zip, doc/xls/ppt/msg, jpg, png, gif, bmp, tif, rtf)
    * end markers: pdf without %%EOF, jpg/png/gif without end marker -> NO END MARKER (only a hint; many open fine,
      open_check_files.py gives the verdict); bad header or zip/Office files that do not open -> CORRUPT
    * text-like types (eml, txt, csv, htm, ...) have no signature and are counted as NOT CHECKED

Output: sheets are ADDED to tests/QCFootprints/output/qc_footprints_report.xlsx (same single report file):
    File_Check_Summary, Attachment_File_Check, Attachment_File_Health, Attachment_Orphan_Files
Run compare_footprints.py first (it rewrites the report), then this script. Close the report in Excel first.
"""
import argparse
import os
import re
import sys
import time
import unicodedata
import zipfile
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import unquote

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import compare_footprints as cf  # noqa: E402  (re-use the DB loader and column names)

DB_PATH_RX = re.compile(r"^D:[\\/]FPextract[\\/]out[\\/]item[\\/]attachment[\\/](\d+)[\\/](\d+)[\\/](\d+)[\\/]([^\\/]+)$")
REPORT = os.path.join(cf.OUTPUT_DIR, "qc_footprints_report.xlsx")

C_MISSING = "FILE MISSING ON DRIVE"
C_ZERO = "ZERO-BYTE FILE"
C_SIZE = "SIZE DIFFERS (UI vs DISK)"
C_SIZE_PATTERN = "DISK ~2.6% SMALLER THAN UI"
C_UI_ONLY = "IN UI BUT NOT IN DB"
C_UI_NEW = "CREATED AFTER DB EXPORT"
C_TOTAL = "TICKET TOTAL SIZE != DB"
C_STALE = "CHANGED AFTER DB EXPORT"
C_BADPATH = "UNEXPECTED DB PATH"
C_COUNT = "PATH COUNT != NAME COUNT"

H_MISMATCH = "EXTENSION / CONTENT MISMATCH"
H_CORRUPT = "CORRUPT (BAD HEADER OR ZIP)"
H_NOEND = "NO END MARKER (CHECK - see File_Open_Check)"
H_UNREAD = "CANNOT READ FILE"


def nfc(s):
    return unicodedata.normalize("NFC", str(s)).lower()


def long_path(p):
    """Windows 260-character limit: use the extended-length prefix for long paths."""
    return "\\\\?\\" + p if len(p) > 240 and not p.startswith("\\\\?\\") else p


def map_path(db_path, drive):
    return drive.rstrip("\\/") + "\\" + db_path[3:].replace("/", "\\")  # drop 'D:\'


def scan_folder(folder, cache):
    """name(NFC lower) -> (real name, size). Missing folder -> empty dict."""
    if folder in cache:
        return cache[folder]
    out = {}
    try:
        with os.scandir(long_path(folder)) as it:
            for e in it:
                if e.is_file():
                    out[nfc(e.name)] = (e.name, e.stat().st_size)
    except (FileNotFoundError, NotADirectoryError):
        pass
    cache[folder] = out
    return out


# ---------------------------------------------------------------- stage 2: health
SIGS = [("pdf", b"%PDF"), ("zip", b"PK\x03\x04"), ("zip", b"PK\x05\x06"), ("ole", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"),
        ("jpg", b"\xff\xd8\xff"), ("png", b"\x89PNG\r\n\x1a\n"), ("gif", b"GIF87a"), ("gif", b"GIF89a"),
        ("bmp", b"BM"), ("tif", b"II*\x00"), ("tif", b"MM\x00*"), ("rtf", b"{\\rtf"),
        ("ole", b"\x09\x04\x06\x00"), ("ole", b"\x09\x02\x06\x00"), ("ole", b"\x09\x08\x10\x00")]  # old Excel BIFF4/3/8 files
EXPECT = {  # extension -> acceptable content kinds
    "pdf": {"pdf"}, "jpg": {"jpg"}, "jpeg": {"jpg"}, "png": {"png"}, "gif": {"gif"}, "bmp": {"bmp"},
    "tif": {"tif"}, "tiff": {"tif"}, "rtf": {"rtf"},
    "docx": {"zip"}, "xlsx": {"zip"}, "pptx": {"zip"}, "docm": {"zip"}, "xlsm": {"zip"}, "xlsb": {"zip"},
    "zip": {"zip"}, "odt": {"zip"}, "ods": {"zip"},
    "doc": {"ole", "rtf", "zip"}, "xls": {"ole", "zip"}, "ppt": {"ole", "zip"}, "msg": {"ole"},
}
OFFICE_ZIP = {"docx", "xlsx", "pptx", "docm", "xlsm", "xlsb"}


def health_check(path, deep):
    """Returns (category, detail) for a problem, ('OK', kind) or ('NOT CHECKED', '')."""
    ext = os.path.splitext(path)[1].lower().lstrip(".")
    if ext not in EXPECT:
        return "NOT CHECKED", ""
    try:
        size = os.path.getsize(long_path(path))
        with open(long_path(path), "rb") as f:
            head = f.read(16)
            f.seek(max(0, size - 2048))
            tail = f.read()
    except OSError as e:
        return H_UNREAD, str(e)[:200]
    kind = next((k for k, sig in SIGS if head.startswith(sig)), None)
    if kind is None:
        return H_CORRUPT, f"header not recognised for .{ext} (starts {head[:8].hex()})"
    if kind not in EXPECT[ext]:
        return H_MISMATCH, f".{ext} but content is {kind}"
    t = tail.rstrip(b"\x00\r\n ")
    if kind == "pdf" and b"%%EOF" not in tail:
        return H_NOEND, "no %%EOF marker; many PDFs still open fine - see File_Open_Check"
    if kind == "jpg" and not t.endswith(b"\xff\xd9"):
        return H_NOEND, "missing JPEG end marker - see File_Open_Check"
    if kind == "png" and b"IEND" not in tail:
        return H_NOEND, "missing PNG IEND chunk - see File_Open_Check"
    if kind == "gif" and not t.endswith(b";"):
        return H_NOEND, "missing GIF trailer - see File_Open_Check"
    if kind == "zip":
        try:
            with zipfile.ZipFile(long_path(path)) as z:
                if ext in OFFICE_ZIP and "[Content_Types].xml" not in z.namelist():
                    return H_CORRUPT, "zip opens but is not a valid Office file"
                if deep:
                    bad = z.testzip()
                    if bad:
                        return H_CORRUPT, f"corrupt member: {bad}"
        except (zipfile.BadZipFile, OSError, EOFError, RuntimeError, NotImplementedError) as e:
            return H_CORRUPT, f"zip does not open ({str(e)[:100]})"
    return "OK", kind


# ---------------------------------------------------------------- stage 1
def load_inputs(att_path, db_path):
    ui_att = pd.read_excel(att_path, sheet_name="Attachments", dtype=object)
    ui_att["_key"] = ui_att["Record Number"].astype(str).str.strip()
    ui_att["_size"] = pd.to_numeric(ui_att["Size (bytes)"], errors="coerce")
    ui_att["_created"] = pd.to_datetime(ui_att["Creation Date"], errors="coerce", format="ISO8601")
    ui_tix = pd.read_excel(att_path, sheet_name="Ticket_Counts", dtype=object)
    ui_keys = set(ui_tix["Record Number"].astype(str).str.strip())

    db_all = cf.clean_db(cf.load_db(db_path))
    deleted = db_all[cf.DB_DELETED_FLAG].astype(str).str.strip().str.lower() == "yes"
    db = db_all[~deleted].drop_duplicates("_key").set_index("_key")
    stamps = pd.concat([pd.to_datetime(db_all["Ticket Created Date"], errors="coerce"),
                        pd.to_datetime(db_all["Last Updated Date"], errors="coerce")])
    cutoff = stamps.max().tz_localize("UTC").tz_convert(cf.APP_TZ).tz_localize(None)
    return ui_att, ui_keys, db, cutoff


def stage1(ui_att, ui_keys, db, cutoff, drive, limit):
    ui_by = {k: g for k, g in ui_att.groupby("_key")}
    keys = sorted(k for k in ui_keys if k in db.index)
    if limit:
        keys = keys[:limit]
    dbcount = pd.to_numeric(db["Attachment Count"], errors="coerce").fillna(0).astype(int)

    # folder prefix for tickets with no DB attachments (orphan scan): most common <a>\<b> in DB paths
    ab = Counter()
    for v in db[cf.DB_PATHS].dropna().head(2000):
        for p in cf.split_list(v):
            m = DB_PATH_RX.match(p)
            if m:
                ab[(m.group(1), m.group(2))] += 1
    a, b = ab.most_common(1)[0][0] if ab else ("6257", "6518")

    rows, orphans, health_targets = [], [], []
    cache = {}
    st = Counter()
    t0 = time.time()
    for i, k in enumerate(keys, 1):
        if i % 2000 == 0:
            print(f"  ... {i}/{len(keys)} tickets ({time.time() - t0:.0f}s)")
        st["tickets"] += 1
        row_start = len(rows)
        g = ui_by.get(k)
        ui_items = [] if g is None else [(str(r["File Name"]), r["_size"], r["_created"]) for _, r in g.iterrows()]
        raw_names = cf.split_list(db.loc[k, cf.DB_NAMES])
        paths = cf.split_list(db.loc[k, cf.DB_PATHS])
        internal = str(int(float(db.loc[k, cf.DB_INTERNAL_ID])))
        base = {"Record Number": k}

        if len(raw_names) != len(paths):
            rows.append({**base, "Category": C_COUNT, "Attachment name": "", "UI size (bytes)": "",
                         "Disk size (bytes)": "", "DB path": "", "Mapped path": "", "Note": ""})
            continue

        entries, folders = [], set()
        for nr, p in zip(raw_names, paths):
            name = unquote(nr)
            if not DB_PATH_RX.match(p):
                rows.append({**base, "Category": C_BADPATH, "Attachment name": name, "UI size (bytes)": "",
                             "Disk size (bytes)": "", "DB path": p, "Mapped path": "", "Note": ""})
                continue
            mapped = map_path(p, drive)
            folder, fname = os.path.split(mapped)
            lst = scan_folder(folder, cache)
            hit = lst.get(nfc(fname))
            entries.append({"name": name, "db_path": p, "mapped": mapped, "folder": folder, "fname": fname,
                            "disk": hit[1] if hit else None})
            folders.add(folder)
        st["db_files"] += len(entries)

        # pair UI items and DB entries by attachment name (sorted by size when a name repeats)
        ui_g, db_g = defaultdict(list), defaultdict(list)
        for n, s, c in ui_items:
            ui_g[nfc(n)].append((n, s, c))
        for e in entries:
            db_g[nfc(e["name"])].append(e)
        pairs, db_only, ui_only = [], [], []
        for nk in set(ui_g) | set(db_g):
            u = sorted(ui_g.get(nk, []), key=lambda x: (x[1] if pd.notna(x[1]) else -1))
            d = sorted(db_g.get(nk, []), key=lambda x: (x["disk"] if x["disk"] is not None else -1))
            for uu, dd in zip(u, d):
                pairs.append((uu, dd))
            db_only += d[len(u):]
            ui_only += u[len(d):]

        ticket_stale = any(pd.notna(c) and c > cutoff for _, _, c in ui_only)
        has_missing = False
        for uu, dd in pairs + [(None, e) for e in db_only]:
            ui_size = None if uu is None or pd.isna(uu[1]) else int(uu[1])
            r = {**base, "Attachment name": dd["name"], "UI size (bytes)": ui_size if ui_size is not None else "",
                 "Disk size (bytes)": dd["disk"] if dd["disk"] is not None else "",
                 "DB path": dd["db_path"], "Mapped path": dd["mapped"], "Note": ""}
            if dd["disk"] is None:
                has_missing = True
                st["missing"] += 1
                rows.append({**r, "Category": C_MISSING})
                continue
            health_targets.append((k, dd["name"], dd["mapped"], dd["disk"]))
            if dd["disk"] == 0 and ui_size != 0:
                st["zero"] += 1
                rows.append({**r, "Category": C_ZERO})
            elif ui_size is not None and ui_size != dd["disk"]:
                ratio = dd["disk"] / ui_size if ui_size else 0
                pattern = 0.973 <= ratio <= 0.976  # the usual 'about 2.6% smaller' pattern
                st["size_pattern" if pattern else "size_diff"] += 1
                rows.append({**r, "Category": C_SIZE_PATTERN if pattern else C_SIZE,
                             "Note": f"difference {dd['disk'] - ui_size:+d} bytes (disk/UI = {ratio:.4f})"})
            else:
                st["ok"] += 1
        for n, s, c in ui_only:
            new = pd.notna(c) and c > cutoff
            st["ui_new" if new else "ui_only"] += 1
            rows.append({**base, "Category": C_UI_NEW if new else C_UI_ONLY, "Attachment name": n,
                         "UI size (bytes)": int(s) if pd.notna(s) else "", "Disk size (bytes)": "",
                         "DB path": "", "Mapped path": "", "Note": f"UI created {c}"})

        # total size on disk vs DB total (MiB, rounded to 2 decimals in the DB)
        if entries and not has_missing:
            disk_mib = sum(e["disk"] for e in entries) / 1048576
            db_mb = pd.to_numeric(db.loc[k, cf.DB_SIZE], errors="coerce")
            if pd.isna(db_mb) or abs(db_mb - disk_mib) > cf.SIZE_TOLERANCE_MB:
                rows.append({**base, "Category": C_STALE if ticket_stale else C_TOTAL, "Attachment name": "",
                             "UI size (bytes)": "", "Disk size (bytes)": "", "DB path": "", "Mapped path": "",
                             "Note": f"DB {db_mb} MB, disk {disk_mib:.2f} MiB"})
                st["total_diff"] += 1

        # orphan files: in the ticket folder but not referenced by any DB path
        scan_dirs = folders or ({f"{drive.rstrip(chr(92) + '/')}\\FPextract\\out\\item\\attachment\\{a}\\{b}\\{internal}"}
                                if dbcount.loc[k] == 0 else set())
        ref = {nfc(e["fname"]) for e in entries}
        orphans_k = []
        for folder in scan_dirs:
            for nk, (real, size) in scan_folder(folder, cache).items():
                if nk not in ref:
                    orphans_k.append({"Record Number": k, "Folder": folder, "File name": real, "Size (bytes)": size,
                                      "DB attachment count": int(dbcount.loc[k])})
        st["orphans"] += len(orphans_k)
        orphans += orphans_k
        # hint for missing files: unreferenced file in the folder with the same size
        for r in rows[row_start:]:
            if r["Category"] == C_MISSING and r["UI size (bytes)"] != "":
                same = [o["File name"] for o in orphans_k if o["Size (bytes)"] == r["UI size (bytes)"]]
                if same:
                    r["Note"] = "unreferenced file with same size: " + "; ".join(same[:3])
        cache.clear()  # folder listings are only needed for one ticket
    return rows, orphans, health_targets, st, len(keys)


def stage2(targets, deep, threads):
    print(f"Health check of {len(targets)} files ({threads} threads)...")
    t0 = time.time()
    out, done = [], 0

    def work(t):
        return t, health_check(t[2], deep)

    kinds = Counter()
    with ThreadPoolExecutor(max_workers=threads) as ex:
        for t, (cat, detail) in ex.map(work, targets, chunksize=64):
            done += 1
            if done % 4000 == 0:
                print(f"  ... {done}/{len(targets)} ({time.time() - t0:.0f}s)")
            kinds[cat if cat in ("OK", "NOT CHECKED") else cat] += 1
            if cat not in ("OK", "NOT CHECKED"):
                k, name, mapped, size = t
                out.append({"Record Number": k, "Category": cat, "Attachment name": name,
                            "Extension": os.path.splitext(mapped)[1].lower().lstrip("."),
                            "Disk size (bytes)": size, "Detail": detail, "Mapped path": mapped})
    return out, kinds


def write_report(report, summary, rows, orphans, health):
    def df(data, cols):
        return pd.DataFrame(data) if data else pd.DataFrame(columns=cols)

    sheets = {
        "File_Check_Summary": summary,
        "Attachment_File_Check": df(rows, ["Record Number", "Category", "Attachment name", "UI size (bytes)",
                                           "Disk size (bytes)", "DB path", "Mapped path", "Note"]),
        "Attachment_File_Health": df(health, ["Record Number", "Category", "Attachment name", "Extension",
                                              "Disk size (bytes)", "Detail", "Mapped path"]),
        "Attachment_Orphan_Files": df(orphans, ["Record Number", "Folder", "File name", "Size (bytes)",
                                                "DB attachment count"]),
    }
    mode = "a" if os.path.exists(report) else "w"
    kw = {"if_sheet_exists": "replace"} if mode == "a" else {}
    try:
        with pd.ExcelWriter(report, engine="openpyxl", mode=mode, **kw) as xw:
            for name, d in sheets.items():
                d.iloc[:1_000_000].to_excel(xw, sheet_name=name, index=False)
    except PermissionError:
        sys.exit(f"Cannot write {report} - close it in Excel and run again.")


def main():
    try:
        sys.stdout.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--drive", default="Z:", help="drive that shows the export's D: drive (default Z:)")
    p.add_argument("--att", default=cf.DEFAULT_ATT_FILE, help="attachments extract (UI data)")
    p.add_argument("--db", default=cf.DEFAULT_DB_FILE, help="DB export Excel")
    p.add_argument("--report", default=REPORT, help="report file the sheets are added to")
    p.add_argument("--limit", type=int, default=0, help="only the first N tickets (trial run)")
    p.add_argument("--no-health", action="store_true", help="skip stage 2 (file health)")
    p.add_argument("--deep", action="store_true", help="read every byte of zip/Office files (slow)")
    p.add_argument("--threads", type=int, default=16)
    args = p.parse_args()

    root = args.drive.rstrip("\\/") + "\\FPextract"
    if not os.path.isdir(root):
        sys.exit(f"Cannot reach {root} - is the drive {args.drive} mapped to the export server's D drive?")

    ui_att, ui_keys, db, cutoff = load_inputs(args.att, args.db)
    print(f"UI attachments: {len(ui_att)} files / {len(ui_keys)} tickets;  DB live tickets: {len(db)};  "
          f"DB export cutoff (Eastern): {cutoff}")
    rows, orphans, targets, st, n_tix = stage1(ui_att, ui_keys, db, cutoff, args.drive, args.limit)

    health, kinds = [], Counter()
    if not args.no_health:
        health, kinds = stage2(targets, args.deep, args.threads)

    cat = Counter(r["Category"] for r in rows)
    srows = [("Drive", args.drive), ("Tickets checked", n_tix), ("DB attachment paths", st["db_files"]),
             ("--- STAGE 1: existence and sizes ---", ""),
             ("Files present, size equals UI", st["ok"]),
             (C_MISSING, st["missing"]), (C_ZERO, st["zero"]),
             (C_SIZE_PATTERN, st["size_pattern"]), (C_SIZE + " - other", st["size_diff"]),
             (C_UI_ONLY, st["ui_only"]), (C_UI_NEW + " (expected)", st["ui_new"]),
             (C_TOTAL, cat[C_TOTAL]), (C_STALE + " (total size, expected)", cat[C_STALE]),
             (C_BADPATH, cat[C_BADPATH]), (C_COUNT, cat[C_COUNT]),
             ("Orphan files (in folder, no DB path)", st["orphans"])]
    if not args.no_health:
        srows += [("--- STAGE 2: file health ---", ""), ("Files read OK (header/end marker/zip)", kinds["OK"]),
                  ("Not checked (text-like type, no signature)", kinds["NOT CHECKED"]),
                  (H_MISMATCH, kinds[H_MISMATCH]), (H_NOEND, kinds[H_NOEND]),
                  (H_CORRUPT, kinds[H_CORRUPT]), (H_UNREAD, kinds[H_UNREAD])]
    summary = pd.DataFrame(srows, columns=["Item", "Value"])

    print("\n==== ATTACHMENT FILE CHECK ====")
    for item, v in srows:
        print(f"{item:<46}: {v}")
    if args.limit:
        print(f"\n(--limit {args.limit}: trial run, NOT written to the report)")
        return
    write_report(args.report, summary, rows, orphans, health)
    print(f"\nSheets added to: {args.report}")


if __name__ == "__main__":
    main()
