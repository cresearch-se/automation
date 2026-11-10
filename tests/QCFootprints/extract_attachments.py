"""
Extract attachment info (file name, size, extension, date, path) for every Footprints ticket.

The app extract has no attachment columns, so this calls the per-ticket attachment API:
    GET https://crcofpsc.cornerstone.com/footprints/servicedesk/api/item/attachment/<Item ID>/<ID>?limit=25&page=1&start=0
where <Item ID> (21103, same for all IR tickets) and <ID> (e.g. 7399 for IR-57299) both come from the
app extract (tests/QCFootprints/fixtures/QCFootprints_<date>.xlsx). So fetch that extract first.

Run on the LOCAL machine (the app is not reachable from the dev server), from the repo root:

    # 1) Probe: one ticket, print the response structure
    python tests/QCFootprints/extract_attachments.py --auth cookie --insecure --probe IR-57299

    # 2) Default run: every ticket the DB export says has attachments (12.5k) + a random sample of
    #    500 tickets the DB says have none (to confirm the UI also shows none)
    python tests/QCFootprints/extract_attachments.py --auth cookie --insecure

    # Other scopes
    python tests/QCFootprints/extract_attachments.py --auth cookie --insecure --tickets IR-57299,IR-57327
    python tests/QCFootprints/extract_attachments.py --auth cookie --insecure --first 200      # first N tickets
    python tests/QCFootprints/extract_attachments.py --auth cookie --insecure --all            # all ~30k tickets

Each ticket's response is cached in output/raw_att/<ID>.json, so a re-run after a cookie expiry resumes
where it stopped (use --refresh to ignore the cache). Output (the signed token after '?' in the file
path is removed):  fixtures/QCFootprints_Attachments_<YYYYMMDD>.xlsx
"""
import argparse
import glob
import json
import os
import random
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests

from extract_footprints import FIXTURES_DIR, HERE, NotJsonError, build_session, describe_html

RAW_ATT_DIR = os.path.join(HERE, "output", "raw_att")
DB_FILE = os.path.join(FIXTURES_DIR, "QCFootprint_DB_Details.xlsx")
BASE_URL = "https://crcofpsc.cornerstone.com/footprints/servicedesk/api/item/attachment"
PAGE_SIZE = 25
_local = threading.local()


# ---------------------------------------------------------------- inputs
def newest_app_extract():
    files = sorted(glob.glob(os.path.join(FIXTURES_DIR, "QCFootprints_2*.xlsx")))
    if not files:
        sys.exit(f"No app extract found in {FIXTURES_DIR} - run extract_footprints.py first.")
    return files[-1]


def load_tickets(app_file):
    df = pd.read_excel(app_file, usecols=["Record Number", "ID", "Item ID"])
    df = df.dropna(subset=["Record Number", "ID", "Item ID"]).copy()
    df["ID"] = df["ID"].astype("int64")
    df["Item ID"] = df["Item ID"].astype("int64")
    return df.drop_duplicates("Record Number").reset_index(drop=True)


def db_attachment_flags():
    """Ticket Number -> (Has Attachments, Attachment Count) from the DB export."""
    print(f"Reading {os.path.basename(DB_FILE)} (takes a minute)...")
    db = pd.read_excel(DB_FILE, sheet_name="Sheet1", usecols=range(20))
    db = db[db["Ticket Internal ID"].notna()]
    db = db[db["Flagged Deleted in FootPrints"].astype(str).str.strip().str.lower() != "yes"]
    return db.set_index("Ticket Number")[["Has Attachments", "Attachment Count"]]


def choose_tickets(args, tickets):
    if args.tickets:
        wanted = [t.strip().upper() for t in args.tickets.split(",") if t.strip()]
        sel = tickets[tickets["Record Number"].isin(wanted)]
        missing = set(wanted) - set(sel["Record Number"])
        if missing:
            print(f"[warn] not in the app extract: {sorted(missing)}")
        return sel
    if args.first:
        return tickets.head(args.first)
    if args.all:
        return tickets
    flags = db_attachment_flags()
    has = flags["Has Attachments"].astype(str).str.strip().str.lower() == "yes"
    with_att = set(flags.index[has])
    without = sorted(set(flags.index[~has]) & set(tickets["Record Number"]))
    random.seed(args.seed)
    sample = set(random.sample(without, min(args.sample_none, len(without))))
    print(f"DB says {len(with_att)} tickets have attachments; adding a random sample of "
          f"{len(sample)} of {len(without)} tickets with none.")
    return tickets[tickets["Record Number"].isin(with_att | sample)]


# ---------------------------------------------------------------- HTTP
def thread_session(base):
    """One requests.Session per worker thread, copying auth/verify/headers from the main session."""
    s = getattr(_local, "session", None)
    if s is None:
        s = requests.Session()
        s.verify = base.verify
        s.auth = base.auth
        s.headers.update(base.headers)
        _local.session = s
    return s


def fetch_json(base, url, params, args):
    session = thread_session(base)
    last = None
    for attempt in range(1, args.retries + 1):
        try:
            r = session.get(url, params=params, headers={"Accept": "application/json"}, timeout=args.timeout)
            r.raise_for_status()
            try:
                return r.json()
            except ValueError:
                raise NotJsonError(describe_html(r))
        except NotJsonError:
            raise
        except requests.exceptions.SSLError as e:
            raise RuntimeError(f"SSL certificate problem - re-run with --insecure. ({e})")
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(2 ** attempt)
    raise RuntimeError(f"{url} failed after {args.retries} attempts: {last}")


def fetch_ticket(base, args, item_id, ticket_id):
    """All attachments for one ticket (pages through if there are more than PAGE_SIZE)."""
    url = f"{args.url}/{item_id}/{ticket_id}"
    items, page, payloads = [], 1, []
    while True:
        params = {"_dc": int(time.time() * 1000), "page": page, "start": (page - 1) * PAGE_SIZE, "limit": PAGE_SIZE}
        payload = fetch_json(base, url, params, args)
        payloads.append(payload)
        data = payload.get("data") or []
        items.extend(data)
        total = payload.get("totalCount", payload.get("total"))
        if len(data) < PAGE_SIZE or (total is not None and len(items) >= int(total)):
            break
        page += 1
    if not payloads[0].get("success", True):
        raise RuntimeError(f"server said success=false for ticket {ticket_id}: {str(payloads[0])[:200]}")
    return items


def get_ticket(base, args, row):
    path = os.path.join(RAW_ATT_DIR, f"{row['ID']}.json")
    if os.path.exists(path) and not args.refresh:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    items = fetch_ticket(base, args, row["Item ID"], row["ID"])
    with open(path, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False)
    time.sleep(args.delay)
    return items


# ---------------------------------------------------------------- modes
def probe(base, args, tickets):
    t = args.probe.strip().upper()
    sel = tickets[tickets["Record Number"] == t]
    if sel.empty:
        sys.exit(f"{t} not found in the app extract.")
    row = sel.iloc[0]
    url = f"{args.url}/{row['Item ID']}/{row['ID']}"
    print(f"GET {url}")
    payload = fetch_json(base, url, {"_dc": int(time.time() * 1000), "page": 1, "start": 0, "limit": PAGE_SIZE}, args)
    keys = [k for k in payload if k != "data"]
    print(f"top-level keys: {list(payload)}   other values: { {k: payload[k] for k in keys} }")
    data = payload.get("data") or []
    print(f"attachments returned: {len(data)}")
    for a in data:
        print(f"  {a.get('name')}.{a.get('extension')}  size={a.get('size')}  created={a.get('creationDate')}")


def build_excel(selected, args):
    att_rows, count_rows, failed = [], [], []
    for _, row in selected.iterrows():
        path = os.path.join(RAW_ATT_DIR, f"{row['ID']}.json")
        if not os.path.exists(path):
            failed.append(row["Record Number"])
            continue
        with open(path, "r", encoding="utf-8") as f:
            items = json.load(f)
        count_rows.append({"Record Number": row["Record Number"], "ID": row["ID"], "UI Attachment Count": len(items)})
        for a in items:
            att_rows.append({
                "Record Number": row["Record Number"],
                "ID": row["ID"],
                "Name": a.get("name"),
                "Extension": a.get("extension"),
                "File Name": f"{a.get('name')}.{a.get('extension')}" if a.get("extension") else a.get("name"),
                "Size (bytes)": a.get("size"),
                "Creation Date": a.get("creationDate"),
                "Path": str(a.get("uri") or "").split("?")[0],  # drop the signed token
                "Mime Type": a.get("mimeType"),
            })
    att = pd.DataFrame(att_rows, columns=["Record Number", "ID", "Name", "Extension", "File Name",
                                          "Size (bytes)", "Creation Date", "Path", "Mime Type"])
    cnt = pd.DataFrame(count_rows, columns=["Record Number", "ID", "UI Attachment Count"])
    out = os.path.join(FIXTURES_DIR, f"QCFootprints_Attachments_{time.strftime('%Y%m%d')}.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        att.to_excel(xw, sheet_name="Attachments", index=False)
        cnt.to_excel(xw, sheet_name="Ticket_Counts", index=False)
    print("\n==== ATTACHMENT EXTRACT SUMMARY ====")
    print(f"output file            : {out}")
    print(f"tickets in output      : {len(cnt)}  (selected {len(selected)}, no data for {len(failed)})")
    print(f"tickets with attachments: {(cnt['UI Attachment Count'] > 0).sum()}")
    print(f"attachments total       : {len(att)}")
    if failed:
        print(f"[warn] tickets with no response (re-run to retry): {failed[:10]}{' ...' if len(failed) > 10 else ''}")


def run(base, args, tickets):
    os.makedirs(RAW_ATT_DIR, exist_ok=True)
    selected = choose_tickets(args, tickets)
    total = len(selected)
    todo = [r for _, r in selected.iterrows()
            if args.refresh or not os.path.exists(os.path.join(RAW_ATT_DIR, f"{r['ID']}.json"))]
    print(f"tickets selected: {total}; already cached: {total - len(todo)}; to fetch: {len(todo)} "
          f"({args.workers} parallel)")
    done, errors, t0 = 0, [], time.time()
    stop = False
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures = {pool.submit(get_ticket, base, args, r): r for r in todo}
        try:
            for fut in as_completed(futures):
                r = futures[fut]
                try:
                    fut.result()
                    done += 1
                except NotJsonError as e:
                    print(f"\n{e}\nCookie probably expired - get a fresh one and re-run (cached tickets are kept).")
                    stop = True
                    break
                except Exception as e:  # noqa: BLE001
                    errors.append((r["Record Number"], str(e)[:150]))
                if done and done % 200 == 0:
                    rate = done / (time.time() - t0)
                    print(f"  {done}/{len(todo)} fetched  ({rate:.1f}/s, ~{(len(todo) - done) / rate / 60:.0f} min left)")
        finally:
            if stop:
                for f in futures:
                    f.cancel()
    if errors:
        print(f"\n{len(errors)} tickets failed (first 5): {errors[:5]}")
    if stop:
        sys.exit(1)
    build_excel(selected, args)


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--probe", metavar="IR-NNNNN", help="fetch one ticket and print the response")
    p.add_argument("--tickets", help="comma-separated IR numbers")
    p.add_argument("--first", type=int, help="only the first N tickets of the app extract")
    p.add_argument("--all", action="store_true", help="every ticket (default: DB-flagged + a sample of the rest)")
    p.add_argument("--sample-none", type=int, default=500, help="random tickets the DB says have no attachments")
    p.add_argument("--seed", type=int, default=1)
    p.add_argument("--app", help="app extract file (default: newest fixtures/QCFootprints_2*.xlsx)")
    p.add_argument("--url", default=BASE_URL)
    p.add_argument("--auth", choices=["ntlm", "cookie"], default="ntlm")
    p.add_argument("--header", action="append")
    p.add_argument("--workers", type=int, default=4, help="parallel requests")
    p.add_argument("--timeout", type=int, default=60)
    p.add_argument("--retries", type=int, default=3)
    p.add_argument("--delay", type=float, default=0.1, help="seconds each worker waits after a ticket")
    p.add_argument("--refresh", action="store_true", help="ignore cached responses")
    p.add_argument("--insecure", action="store_true")
    args = p.parse_args()

    app_file = args.app or newest_app_extract()
    print(f"App extract: {os.path.basename(app_file)}")
    tickets = load_tickets(app_file)
    print(f"{len(tickets)} tickets (Item IDs: {sorted(tickets['Item ID'].unique())[:3]})")
    base = build_session(args)
    if args.probe:
        probe(base, args, tickets)
    else:
        run(base, args, tickets)


if __name__ == "__main__":
    main()
