"""
Extract all rows from the Footprints advanced-search API into one Excel file.

Run on the LOCAL machine (the app is not reachable from the dev server), from the repo root:

    # 1) Probe: fetch 5 rows, save raw JSON, print the response structure. Send the output back.
    python tests/QCFootprints/extract_api_tickets.py --probe

    # 2) Full extract: all pages (limit 500) -> fixtures/QCFootprints_<YYYYMMDD>.xlsx
    python tests/QCFootprints/extract_api_tickets.py

Inputs (all in tests/QCFootprints/):
    request_body.txt   raw POST body copied from the browser (DevTools > Network > the "run" request > Payload)
    column_names.json  OPTIONAL override list of column names, in rowValues order.
                       Without it, names come from the response's own 'columns' list.

Auth (--auth):
    ntlm   (default) FP_USERNAME / FP_PASSWORD from config/footprints.env (or prompted if not set)
    cookie FP_COOKIE from config/footprints.env: the Cookie header value (DevTools > Network > Request Headers)
Extra headers some apps need (e.g. CSRF / X-Requested-With):  --header "Name: value"  (repeatable)

Raw pages are cached in output/raw/page_NNN.json, so a re-run after a failure resumes where it stopped.
Use --refresh to ignore the cache.
"""
import argparse
import getpass
import glob
import json
import math
import os
import re
import sys
import time

import pandas as pd
import requests
from dotenv import load_dotenv

HERE = os.path.dirname(os.path.abspath(__file__))
# Credentials live in config/footprints.env (gitignored via *.env) - see footprints.env.example
load_dotenv(os.path.join(HERE, "..", "..", "config", "footprints.env"))
RAW_DIR = os.path.join(HERE, "output", "raw")
FIXTURES_DIR = os.path.join(HERE, "fixtures")
BODY_FILE = os.path.join(HERE, "request_body.txt")
COLUMNS_FILE = os.path.join(HERE, "column_names.json")

DEFAULT_URL = "https://crcofpsc.cornerstone.com/footprints/servicedesk/api/search/advanced/run"
EXCEL_CELL_LIMIT = 32767
ILLEGAL_XML_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


# ---------------------------------------------------------------- HTTP
class NotJsonError(RuntimeError):
    """Server returned HTML instead of JSON (login page, error page, ...)."""


def describe_html(r):
    """Save the HTML response and summarise it so the cause is visible."""
    os.makedirs(os.path.join(HERE, "output"), exist_ok=True)
    path = os.path.join(HERE, "output", "last_response.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(r.text)
    title = re.search(r"<title[^>]*>(.*?)</title>", r.text, re.I | re.S)
    text = re.sub(r"<(script|style).*?</\1>", " ", r.text, flags=re.I | re.S)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    redirects = " -> ".join(f"{h.status_code} {h.url}" for h in r.history) or "none"
    return (
        "Server returned HTML instead of JSON.\n"
        f"  HTTP status     : {r.status_code}\n"
        f"  final URL       : {r.url}\n"
        f"  redirects       : {redirects}\n"
        f"  page title      : {title.group(1).strip() if title else '(none)'}\n"
        f"  visible text    : {text[:300]!r}\n"
        f"  full page saved : {path}\n"
        "  Likely cause: not logged in (missing/expired cookie), or the request body/headers are wrong."
    )


def build_session(args):
    s = requests.Session()
    s.verify = not args.insecure
    if args.insecure:
        requests.packages.urllib3.disable_warnings()
    if args.auth == "ntlm":
        from requests_ntlm import HttpNtlmAuth
        user = os.getenv("FP_USERNAME") or input("Username (e.g. DOMAIN\\user): ").strip()
        pwd = os.getenv("FP_PASSWORD") or getpass.getpass("Password: ")
        print(f"Authenticating as {user}")
        s.auth = HttpNtlmAuth(user, pwd)
    else:
        cookie = os.getenv("FP_COOKIE") or getpass.getpass("Cookie header value: ")
        cookie = cookie.strip()
        if cookie.lower().startswith("cookie:"):
            cookie = cookie[7:].strip()
        print(f"Cookie received: {len(cookie)} characters" + ("  <-- EMPTY!" if not cookie else ""))
        s.headers["Cookie"] = cookie
    for h in args.header or []:
        name, _, value = h.partition(":")
        s.headers[name.strip()] = value.strip()
    return s


def read_body():
    if not os.path.exists(BODY_FILE):
        print(f"[warn] {BODY_FILE} not found - sending an EMPTY body. "
              "The server will probably return nothing useful.")
        return ""
    with open(BODY_FILE, "r", encoding="utf-8") as f:
        body = f.read().strip()
    print(f"Request body loaded: {len(body)} characters" + ("  <-- EMPTY!" if not body else ""))
    return body


def fetch_page(session, args, body, page, limit):
    params = {
        "_dc": int(time.time() * 1000),
        "limit": limit,
        "page": page,
        "start": (page - 1) * limit,
    }
    headers = {"Content-Type": args.content_type, "Accept": "application/json"}
    last_err = None
    for attempt in range(1, args.retries + 1):
        try:
            r = session.post(args.url, params=params, data=body.encode("utf-8"),
                             headers=headers, timeout=args.timeout)
            r.raise_for_status()
            try:
                return r.json()
            except ValueError:
                raise NotJsonError(describe_html(r))
        except NotJsonError:
            raise  # login/error page: retrying will not help
        except requests.exceptions.SSLError as e:
            raise RuntimeError(f"SSL certificate problem - re-run with --insecure. ({e})")
        except Exception as e:  # noqa: BLE001 - retry on anything, report the last one
            last_err = e
            wait = 2 ** attempt
            print(f"  page {page} attempt {attempt}/{args.retries} failed: {e}  (retry in {wait}s)")
            time.sleep(wait)
    raise RuntimeError(f"page {page} failed after {args.retries} attempts: {last_err}")


# ---------------------------------------------------------------- parsing
def get_blocks(payload):
    # Server uses "pageData" (older/other responses used "data")
    data = payload.get("pageData") or payload.get("data")
    return data if isinstance(data, list) else []


NAME_KEYS = ("name", "label", "displayName", "header", "title", "text", "fieldName", "dataIndex", "caption")


def get_columns(payload):
    """Column names from the response's 'columns' list, or None if they can't be worked out."""
    blocks = get_blocks(payload)
    cols = blocks[0].get("columns") if blocks else None
    if not cols:
        return None
    names = []
    for i, c in enumerate(cols):
        name = None
        if isinstance(c, dict):
            for k in NAME_KEYS:
                if c.get(k):
                    name = str(c[k]).strip()
                    break
        names.append(name or f"Col_{i + 1:02d}")
    # make names unique (Excel/pandas get confused by duplicates)
    seen = {}
    for i, n in enumerate(names):
        seen[n] = seen.get(n, 0) + 1
        if seen[n] > 1:
            names[i] = f"{n}_{seen[n]}"
    return names


def get_rows(payload):
    rows = []
    for block in get_blocks(payload):
        rows.extend(block.get("rowsData", []))
    return rows


def get_total(payload):
    for key in ("totalCount", "total"):
        if payload.get(key) is not None:
            return int(payload[key])
    blocks = get_blocks(payload)
    return int(blocks[0]["total"]) if blocks and "total" in blocks[0] else None


def outline(obj, depth=0, max_depth=4, key="root"):
    """Print structure of a JSON value without dumping row data."""
    pad = "  " * depth
    if isinstance(obj, dict):
        print(f"{pad}{key}: dict({len(obj)} keys)")
        if depth < max_depth:
            for k, v in obj.items():
                outline(v, depth + 1, max_depth, k)
    elif isinstance(obj, list):
        print(f"{pad}{key}: list(len={len(obj)})")
        if obj and depth < max_depth and key != "rowValues":
            outline(obj[0], depth + 1, max_depth, "[0]")
    else:
        text = repr(obj)
        print(f"{pad}{key}: {type(obj).__name__} = {text[:60]}{'...' if len(text) > 60 else ''}")


def clean_cell(v):
    if isinstance(v, list):
        if all(isinstance(i, dict) and "userName" in i for i in v):
            return "; ".join(i["userName"] for i in v)
        return json.dumps(v, ensure_ascii=False)
    if isinstance(v, dict):
        return json.dumps(v, ensure_ascii=False)
    return v


def load_column_names():
    if os.path.exists(COLUMNS_FILE):
        with open(COLUMNS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


# ---------------------------------------------------------------- modes
def probe(session, args, body):
    payload = fetch_page(session, args, body, page=1, limit=5)
    os.makedirs(os.path.join(HERE, "output"), exist_ok=True)
    out = os.path.join(HERE, "output", "probe_page1.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    rows = get_rows(payload)
    print(f"\nSaved raw probe response to {out}")
    print(f"total reported by server : {get_total(payload)}")
    print(f"rows in this response    : {len(rows)}")
    if rows:
        print(f"values per row           : {len(rows[0].get('rowValues', []))}")
    print("\n--- response structure (row data omitted) ---")
    outline(payload)
    blocks = get_blocks(payload)
    cols = blocks[0].get("columns") if blocks else None
    if cols:
        print("\n--- first 2 column definitions (full) ---")
        for c in cols[:2]:
            print(json.dumps(c, ensure_ascii=False))
        print("\n--- column names the script will use (rowValues order) ---")
        for i, n in enumerate(get_columns(payload) or []):
            print(f"  {i + 1:>2}. {n}")
        print("\n(Names like Col_NN mean the name key was not recognised - send me the 2 definitions above.)")


def extract(session, args, body):
    os.makedirs(RAW_DIR, exist_ok=True)
    limit = args.limit

    first_path = os.path.join(RAW_DIR, "page_001.json")
    if os.path.exists(first_path) and not args.refresh:
        with open(first_path, "r", encoding="utf-8") as f:
            first = json.load(f)
        print("page 1 loaded from cache")
    else:
        first = fetch_page(session, args, body, 1, limit)
        with open(first_path, "w", encoding="utf-8") as f:
            json.dump(first, f, ensure_ascii=False)
        print(f"page 1 fetched ({len(get_rows(first))} rows)")

    total = get_total(first)
    if total is None:
        sys.exit("Could not find 'total' in the response - run --probe and send me the structure.")
    pages = args.pages or math.ceil(total / limit)
    print(f"server total = {total}; limit = {limit}; pages to fetch = {pages}")

    for page in range(2, pages + 1):
        path = os.path.join(RAW_DIR, f"page_{page:03d}.json")
        if os.path.exists(path) and not args.refresh:
            print(f"page {page}/{pages} cached")
            continue
        payload = fetch_page(session, args, body, page, limit)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
        print(f"page {page}/{pages} fetched ({len(get_rows(payload))} rows)")
        time.sleep(args.delay)

    build_excel(total, pages, limit)


def build_excel(total, pages, limit):
    records, page_counts = [], {}
    for page in range(1, pages + 1):
        path = os.path.join(RAW_DIR, f"page_{page:03d}.json")
        if not os.path.exists(path):
            print(f"[warn] missing raw file for page {page}")
            continue
        with open(path, "r", encoding="utf-8") as f:
            rows = get_rows(json.load(f))
        page_counts[page] = len(rows)
        for r in rows:
            rec = {"_page": page, "_definitionId": r.get("definitionId"), "_metaId": r.get("metaId")}
            for i, v in enumerate(r.get("rowValues", [])):
                rec[i] = clean_cell(v)
            records.append(rec)

    df = pd.DataFrame(records)
    value_cols = [c for c in df.columns if isinstance(c, int)]
    names = load_column_names()  # manual override file wins
    if not names:
        with open(os.path.join(RAW_DIR, "page_001.json"), "r", encoding="utf-8") as f:
            names = get_columns(json.load(f))
    rename = {}
    for c in value_cols:
        rename[c] = names[c] if names and c < len(names) else f"Col_{c + 1:02d}"
    if names and len(names) != len(value_cols):
        print(f"[warn] column_names.json has {len(names)} names but rows have {len(value_cols)} values")
    df = df.rename(columns=rename)

    # Excel hygiene: strip illegal chars, cap cell length (Excel max 32767)
    truncated = 0
    for col in df.columns:
        if df[col].dtype == object:
            def fix(v):
                nonlocal truncated
                if isinstance(v, str):
                    v = ILLEGAL_XML_CHARS.sub("", v)
                    if len(v) > EXCEL_CELL_LIMIT:
                        truncated += 1
                        v = v[:EXCEL_CELL_LIMIT]
                return v
            df[col] = df[col].map(fix)

    os.makedirs(FIXTURES_DIR, exist_ok=True)
    out = os.path.join(FIXTURES_DIR, f"QCFootprints_{time.strftime('%Y%m%d')}.xlsx")
    df.to_excel(out, index=False, sheet_name="Footprints", engine="openpyxl")

    # ---- sanity checks
    first_val_col = rename[value_cols[0]] if value_cols else None
    print("\n==== EXTRACT SUMMARY ====")
    print(f"output file          : {out}")
    print(f"rows extracted       : {len(df)}   (server total: {total})  "
          f"{'OK' if len(df) == total else '<-- MISMATCH'}")
    short = {p: n for p, n in page_counts.items() if n != limit and p != pages}
    print(f"short pages (not last): {short if short else 'none'}")
    if first_val_col:
        dups = df[first_val_col].duplicated().sum()
        print(f"duplicate '{first_val_col}'  : {dups}  {'OK' if dups == 0 else '<-- paging overlap? re-run'}")
    print(f"cells truncated >32767: {truncated}")


# ---------------------------------------------------------------- main
def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--probe", action="store_true", help="fetch 5 rows and print response structure")
    p.add_argument("--url", default=DEFAULT_URL)
    p.add_argument("--auth", choices=["ntlm", "cookie"], default="ntlm")
    p.add_argument("--limit", type=int, default=500)
    p.add_argument("--pages", type=int, help="override page count (default: from server total)")
    p.add_argument("--content-type", default="application/json")
    p.add_argument("--header", action="append", help='extra header, e.g. "X-Requested-With: XMLHttpRequest"')
    p.add_argument("--timeout", type=int, default=120)
    p.add_argument("--retries", type=int, default=3)
    p.add_argument("--delay", type=float, default=0.5, help="seconds between pages")
    p.add_argument("--refresh", action="store_true", help="ignore cached raw pages")
    p.add_argument("--insecure", action="store_true", help="skip TLS verification")
    args = p.parse_args()

    body = read_body()
    session = build_session(args)
    if args.probe:
        probe(session, args, body)
    else:
        extract(session, args, body)


if __name__ == "__main__":
    main()
