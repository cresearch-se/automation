"""
Really OPEN attachment files with real parsers (no Word/Acrobat on this server, so libraries stand in):
PDF -> pypdf (read every page), images -> Pillow (decode every pixel), old .xls -> xlrd, .eml -> email parser,
txt/csv/htm/xml/... -> decode + parse. Nothing is modified; files are only read.

Run on the server (needs the Z: drive and a report that already holds the Attachment_File_Health sheet):

    python tests/QCFootprints/open_check_files.py                 # 78 suspicious files + 200 random text-like files
    python tests/QCFootprints/open_check_files.py --sample 500    # bigger spot check of the text-like files
    python tests/QCFootprints/open_check_files.py --include-mismatch   # also the 13 extension/content mismatches
    python tests/QCFootprints/open_check_files.py --limit 10      # trial, not written to the report

Targets
    SUSPICIOUS       files in sheet Attachment_File_Health, categories NO END MARKER and CORRUPT (written by check_attachment_files.py)
    MISMATCH         (optional) the EXTENSION / CONTENT MISMATCH files
    TEXT SPOT CHECK  random sample of the text-like files check_attachment_files.py could not check (eml, txt, csv, htm, ...)

Results (sheet File_Open_Check, plus File_Open_Summary), added to output/qc_footprints_report.xlsx:
    OPENS OK            parsed completely, no problem
    OPENS WITH WARNINGS parsed, but the parser had to repair or skip something (detail says what)
    FAILS TO OPEN       the parser raised an error (detail says which)
    NO PARSER           type with no parser here (not opened)
Run compare_footprints.py -> check_attachment_files.py -> this script (compare rewrites the report).
"""
import argparse
import csv
import email
import email.policy
import io
import logging
import os
import random
import sys
import warnings
from collections import Counter
from concurrent.futures import ThreadPoolExecutor

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import check_attachment_files as caf  # noqa: E402
import compare_footprints as cf  # noqa: E402

OK, WARN, FAIL, NOPARSER = "OPENS OK", "OPENS WITH WARNINGS", "FAILS TO OPEN", "NO PARSER"
TEXT_LIKE = {"eml", "txt", "csv", "htm", "html", "xml", "log", "json", "md", "tsv", "ini", "msg", "url"}
IMAGE_EXT = {"jpg", "jpeg", "png", "gif", "bmp", "tif", "tiff"}


# ---------------------------------------------------------------- openers: return (result, detail)
class _Capture(logging.Handler):
    def __init__(self):
        super().__init__(level=logging.WARNING)
        self.msgs = []

    def emit(self, record):
        self.msgs.append(record.getMessage()[:120])


def open_pdf(path):
    """MuPDF first (tolerant, like Acrobat/Edge: rebuilds a damaged index), pypdf as the strict second opinion.

    pypdf alone said 'cannot open' for 64 PDFs that viewers open fine, so the verdict comes from MuPDF:
    FAILS = no readable pages; WARNINGS = opens but repaired / no %%EOF end marker / some page unreadable.
    """
    try:
        import pymupdf
    except ImportError:
        return open_pdf_pypdf(path)
    try:
        pymupdf.TOOLS.mupdf_display_errors(False)
        with pymupdf.open(caf.long_path(path)) as doc:
            n, repaired = doc.page_count, doc.is_repaired
            if n == 0:
                return FAIL, f"MuPDF found 0 pages (repaired: {repaired})"
            chars, bad = 0, 0
            for pg in doc:
                try:
                    chars += len(pg.get_text())
                except Exception:
                    bad += 1
        with open(caf.long_path(path), "rb") as f:
            f.seek(max(0, os.path.getsize(caf.long_path(path)) - 2048))
            no_eof = b"%%EOF" not in f.read()
        detail = f"{n} pages, {chars} text chars"
        notes = []
        if bad:
            notes.append(f"{bad} page(s) unreadable")
        if repaired:
            notes.append("index repaired by MuPDF")
        if no_eof:
            notes.append("no %%EOF end marker (viewer opens it anyway)")
        if notes:
            return WARN, detail + "; " + "; ".join(notes)
        return OK, detail
    except Exception as e:
        return FAIL, f"{type(e).__name__}: {str(e)[:150]}"


def open_pdf_pypdf(path):
    from pypdf import PdfReader
    cap = _Capture()
    lg = logging.getLogger("pypdf")
    lg.addHandler(cap)
    try:
        with warnings.catch_warnings(record=True):
            with open(caf.long_path(path), "rb") as f:
                r = PdfReader(f, strict=False)
                if r.is_encrypted:
                    try:
                        r.decrypt("")
                    except Exception:
                        return WARN, "encrypted (password needed), cannot read pages"
                n = len(r.pages)
                chars, bad_pages = 0, 0
                for p in r.pages:  # touch every page: content stream + text
                    try:
                        chars += len(p.extract_text() or "")
                    except Exception:
                        bad_pages += 1
        extra = f"{n} pages, {chars} text chars"
        if bad_pages:
            return WARN, f"{extra}; {bad_pages} page(s) could not be read"
        if cap.msgs:
            return WARN, f"{extra}; parser repaired: {cap.msgs[0]}"
        return OK, extra
    except Exception as e:
        return FAIL, f"{type(e).__name__}: {str(e)[:150]}"
    finally:
        lg.removeHandler(cap)


def open_image(path):
    from PIL import Image, ImageFile
    ImageFile.LOAD_TRUNCATED_IMAGES = False
    try:
        with Image.open(caf.long_path(path)) as im:
            fmt, size = im.format, im.size
            im.load()  # decode all pixel data: a truncated file raises here
        return OK, f"{fmt} {size[0]}x{size[1]}"
    except Exception as e:
        return FAIL, f"{type(e).__name__}: {str(e)[:150]}"


def open_xls(path):
    import xlrd
    try:
        wb = xlrd.open_workbook(caf.long_path(path), on_demand=False, logfile=io.StringIO())
        rows = sum(wb.sheet_by_index(i).nrows for i in range(wb.nsheets))
        return OK, f"{wb.nsheets} sheet(s), {rows} rows"
    except Exception as e:
        return FAIL, f"{type(e).__name__}: {str(e)[:150]}"


def read_text(path):
    raw = open(caf.long_path(path), "rb").read()
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw.decode("utf-16"), "utf-16"
    for enc in ("utf-8-sig", "cp1252"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1"), "latin-1"


def open_eml(path):
    try:
        with open(caf.long_path(path), "rb") as f:
            msg = email.message_from_binary_file(f, policy=email.policy.default)
        missing = [h for h in ("From", "Date", "Subject") if not msg[h]]
        parts = sum(1 for _ in msg.walk())
        body = msg.get_body(preferencelist=("plain", "html"))
        text = body.get_content() if body is not None else ""
        detail = f"{parts} part(s), body {len(text)} chars"
        if len(msg.defects):
            return WARN, f"{detail}; parser defects: {msg.defects[0].__class__.__name__}"
        if missing:
            return WARN, f"{detail}; missing header(s): {', '.join(missing)}"
        return OK, detail
    except Exception as e:
        return FAIL, f"{type(e).__name__}: {str(e)[:150]}"


def open_textlike(path, ext):
    try:
        size = os.path.getsize(caf.long_path(path))
        if size == 0:
            return FAIL, "empty file"
        text, enc = read_text(path)
        nul = text.count("\x00")
        if ext in ("csv", "tsv"):
            rows = list(csv.reader(io.StringIO(text), delimiter="\t" if ext == "tsv" else ","))
            widths = Counter(len(r) for r in rows if r)
            detail = f"{len(rows)} rows, {enc}"
            if len(widths) > 3:
                return WARN, f"{detail}; ragged rows ({len(widths)} different widths)"
            return (WARN, f"{detail}; {nul} NUL chars") if nul else (OK, detail)
        if ext in ("htm", "html", "xml"):
            from html.parser import HTMLParser
            tags = Counter()

            class P(HTMLParser):
                def handle_starttag(self, tag, attrs):
                    tags[tag] += 1
            P().feed(text)
            detail = f"{sum(tags.values())} tags, {enc}"
            return (WARN, f"{detail}; no tags found") if not tags else (OK, detail)
        detail = f"{len(text)} chars, {enc}"
        return (WARN, f"{detail}; {nul} NUL chars (binary?)") if nul else (OK, detail)
    except Exception as e:
        return FAIL, f"{type(e).__name__}: {str(e)[:150]}"


def open_file(path):
    ext = os.path.splitext(path)[1].lower().lstrip(".")
    try:
        if not os.path.exists(caf.long_path(path)):
            return FAIL, "file not found"
        if ext == "pdf":
            return open_pdf(path)
        if ext in IMAGE_EXT:
            return open_image(path)
        if ext == "xls":
            return open_xls(path)
        if ext == "eml":
            return open_eml(path)
        if ext in TEXT_LIKE:
            return open_textlike(path, ext)
        return NOPARSER, f".{ext}: no parser here"
    except Exception as e:  # a parser must never stop the run
        return FAIL, f"{type(e).__name__}: {str(e)[:150]}"


# ---------------------------------------------------------------- choosing the files
def text_like_files(db, drive):
    """(ticket, attachment name, mapped path) of live-ticket files whose type check_attachment_files does not verify."""
    out = []
    for k, row in db.iterrows():
        names = cf.split_list(row.get(cf.DB_NAMES))
        paths = cf.split_list(row.get(cf.DB_PATHS))
        if len(names) != len(paths):
            continue
        for n, p in zip(names, paths):
            ext = os.path.splitext(p)[1].lower().lstrip(".")
            if ext not in caf.EXPECT and p.startswith("D:"):
                out.append((k, caf.unquote(n), caf.map_path(p, drive)))
    return out


def main():
    try:
        sys.stdout.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--drive", default="Z:")
    ap.add_argument("--report", default=caf.REPORT)
    ap.add_argument("--db", default=cf.DEFAULT_DB_FILE)
    ap.add_argument("--sample", type=int, default=200, help="random text-like files to open (default 200, 0 = none)")
    ap.add_argument("--seed", type=int, default=1)
    ap.add_argument("--include-mismatch", action="store_true", help="also open the extension/content mismatches")
    ap.add_argument("--limit", type=int, default=0, help="trial: only the first N targets, not written")
    ap.add_argument("--threads", type=int, default=8)
    args = ap.parse_args()

    if not os.path.isdir(args.drive.rstrip("\\/") + "\\FPextract"):
        sys.exit(f"Cannot reach {args.drive}\\FPextract - is the drive mapped?")
    if not os.path.exists(args.report):
        sys.exit(f"No report at {args.report}. Run compare_footprints.py and check_attachment_files.py first.")
    try:
        health = pd.read_excel(args.report, sheet_name="Attachment_File_Health")
    except ValueError:
        sys.exit("Sheet Attachment_File_Health not found - run check_attachment_files.py first.")

    targets = []
    for _, r in health.iterrows():
        if r["Category"] in (caf.H_CORRUPT, caf.H_NOEND):
            targets.append(("SUSPICIOUS", r["Record Number"], r["Attachment name"], r["Mapped path"], r["Detail"]))
        elif r["Category"] == caf.H_MISMATCH and args.include_mismatch:
            targets.append(("MISMATCH", r["Record Number"], r["Attachment name"], r["Mapped path"], r["Detail"]))
    n_fixed = len(targets)
    if args.sample:
        raw = cf.clean_db(cf.load_db(args.db))
        live = raw[raw[cf.DB_DELETED_FLAG].astype(str).str.strip().str.lower() != "yes"].drop_duplicates("_key")
        pool = text_like_files(live.set_index("_key"), args.drive)
        random.Random(args.seed).shuffle(pool)
        targets += [("TEXT SPOT CHECK", k, n, p, "") for k, n, p in pool[:args.sample]]
        print(f"Text-like files available: {len(pool)}; sampling {min(args.sample, len(pool))} (seed {args.seed})")
    if args.limit:
        targets = targets[:args.limit]
    print(f"Opening {len(targets)} files ({n_fixed} from the health sheet) with {args.threads} threads...")

    def work(t):
        res, detail = open_file(t[3])
        size = os.path.getsize(caf.long_path(t[3])) if os.path.exists(caf.long_path(t[3])) else ""
        return {"Record Number": t[1], "Source": t[0], "Attachment name": t[2],
                "Extension": os.path.splitext(t[3])[1].lower().lstrip("."), "Result": res, "Detail": detail,
                "Earlier health finding": t[4], "Disk size (bytes)": size, "Mapped path": t[3]}

    with ThreadPoolExecutor(max_workers=args.threads) as ex:
        rows = list(ex.map(work, targets))
    df = pd.DataFrame(rows)

    summary = []
    for src, g in df.groupby("Source", sort=False):
        summary.append((f"--- {src} ({len(g)} files) ---", ""))
        for res in (OK, WARN, FAIL, NOPARSER):
            summary.append((f"   {res}", int((g["Result"] == res).sum())))
        for ext, ge in g.groupby("Extension"):
            parts = [f"{int((ge['Result'] == r).sum())} {r.lower()}" for r in (OK, WARN, FAIL, NOPARSER)
                     if (ge["Result"] == r).any()]
            summary.append((f"      .{ext}: {len(ge)} files = " + ", ".join(parts), ""))
    summary = pd.DataFrame(summary, columns=["Item", "Value"])
    print("\n==== OPEN CHECK ====")
    for item, v in summary.itertuples(index=False):
        print(f"{item:<46}: {v}")
    bad = df[df["Result"].isin([FAIL, WARN])]
    if len(bad):
        print("\nFirst problems:")
        print(bad[["Record Number", "Attachment name", "Result", "Detail"]].head(15).to_string(index=False))
    if args.limit:
        print(f"\n(--limit {args.limit}: trial run, NOT written to the report)")
        return
    try:
        with pd.ExcelWriter(args.report, engine="openpyxl", mode="a", if_sheet_exists="replace") as xw:
            summary.to_excel(xw, sheet_name="File_Open_Summary", index=False)
            rank = df["Result"].map({FAIL: 0, WARN: 1, NOPARSER: 2, OK: 3})
            df.assign(_r=rank).sort_values(["_r", "Source"], kind="stable").drop(columns="_r") \
                .to_excel(xw, sheet_name="File_Open_Check", index=False)
    except PermissionError:
        sys.exit(f"Cannot write {args.report} - close it in Excel and run again.")
    print(f"\nSheets File_Open_Summary and File_Open_Check added to: {args.report}")


if __name__ == "__main__":
    main()
