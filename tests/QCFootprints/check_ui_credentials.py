"""
Check in the Footprints UI (API) whether the tickets we flagged as LIKELY REAL PASSWORD still hold a password.

Why: the devs cleaned the database by replacing password values with [REDACTED]. The DB export tells us what the
DB holds, but only the app shows what users see. This script reads the full ticket text from the app for the
tickets on the report's Credential_Review sheet with Assessment = LIKELY REAL PASSWORD - REVIEW and judges each:

    PASSWORD STILL PRESENT - REVIEW   a value after a password label that is not a redaction placeholder
    REDACTED (NO PASSWORD LEFT)       password labels exist but every value is [REDACTED] / removed / xxx
    MENTIONS PASSWORD WORD ONLY       the word appears but no value follows
    NO PASSWORD WORD FOUND            nothing left
    NO TEXT RETURNED                  the endpoint returned no text (wrong endpoint? use --probe)

Values are NEVER written (the report is tracked in git): only counts and a masked context ("... Password: [REDACTED]").
Raw responses are cached in output/raw_ui/ (git-ignored, they may contain real passwords - delete when done).

Run on the LOCAL machine (the app is not reachable from the dev server without a cookie), from the repo root:

    # 1) Find the endpoint that returns a ticket's full description history. Prints status / structure only.
    python tests/QCFootprints/check_ui_credentials.py --auth cookie --insecure --probe IR-68605
    #    If none of the built-in guesses returns the history: open a ticket in the browser, DevTools > Network,
    #    find the request that loads the description entries, and pass it with
    #    --url-template "https://.../api/...{item}/{id}..."   ({item} = Item ID, {id} = ticket ID)

    # 2) Check the 66 tickets (writes sheets UI_Credential_Summary and UI_Credential_Check into the report)
    python tests/QCFootprints/check_ui_credentials.py --auth cookie --insecure [--url-template ...]

Other options: --tickets IR-a,IR-b (instead of the report list), --report PATH, --app PATH (app extract with the
ID / Item ID columns), --refresh (ignore cached responses), --workers N.
Auth is the same as extract_api_tickets.py: a fresh browser Cookie header in config/footprints.env (FP_COOKIE)
or typed at the prompt. Close the report in Excel first.
"""
import argparse
import html
import json
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests

import compare_footprints as cf
from extract_api_attachments import load_tickets, newest_app_extract
from extract_api_tickets import HERE, build_session

API = "https://crcofpsc.cornerstone.com/footprints/servicedesk/api"
REPORT = os.path.join(HERE, "output", "qc_footprints_report.xlsx")
RAW_DIR = os.path.join(HERE, "output", "raw_ui")

# Guesses for the endpoint that returns a ticket's description entries (probe tries them all).
CANDIDATES = [
    API + "/item/{item}/{id}",
    API + "/item/description/{item}/{id}",
    API + "/item/descriptions/{item}/{id}",
    API + "/item/history/{item}/{id}",
    API + "/item/entries/{item}/{id}",
    API + "/item/comment/{item}/{id}",
    API + "/item/note/{item}/{id}",
]

R_PRESENT = "PASSWORD STILL PRESENT - REVIEW"
R_REDACTED = "REDACTED (NO PASSWORD LEFT)"
R_WORD = "MENTIONS PASSWORD WORD ONLY"
R_NONE = "NO PASSWORD WORD FOUND"
R_NOTEXT = "NO TEXT RETURNED"
R_ERROR = "FETCH ERROR"
R_ORDER = {R_PRESENT: 0, R_ERROR: 1, R_NOTEXT: 2, R_WORD: 3, R_REDACTED: 4, R_NONE: 5}

PLACEHOLDER_RX = re.compile(r"(?i)redact|remov|^\W*x{3,}\W*$|^[\W_]*\*{3,}[\W_]*$")


# ---------------------------------------------------------------- text from the response
def strings_in(obj, out):
    """Every string value anywhere in the JSON (keys are ignored)."""
    if isinstance(obj, str):
        out.append(obj)
    elif isinstance(obj, dict):
        for v in obj.values():
            strings_in(v, out)
    elif isinstance(obj, list):
        for v in obj:
            strings_in(v, out)


def response_text(payload):
    parts = []
    strings_in(payload, parts)
    text = "\n".join(parts)
    text = re.sub(r"(?i)<br\s*/?>|</p>|</div>", " ", text)
    text = re.sub(r"<[^>]+>", " ", text)
    return re.sub(r"\s+", " ", html.unescape(text)).strip()


# ---------------------------------------------------------------- judging
def judge(text):
    """(result, real_count, redacted_count, masked context). Never returns a password value."""
    if not text:
        return R_NOTEXT, 0, 0, ""
    hits = [m for m in cf.CRED_VALUE_RX.finditer(text)
            if cf.plausible_value(m) and not cf.is_project_code_or_slug(text, m)]
    is_meeting = lambda m: (m.start() > 0 and text[m.start() - 1] in "?&") or \
        "meeting id" in text[max(0, m.start() - 80):m.start()].lower()
    all_hits = [m for m in cf.CRED_VALUE_RX.finditer(text) if not cf.is_project_code_or_slug(text, m)]
    placeholders = [m for m in all_hits if PLACEHOLDER_RX.search(m.group(1).strip("\"'.,;:)(]>[<"))]
    real = [m for m in hits if not PLACEHOLDER_RX.search(m.group(1)) and not is_meeting(m)]
    if real:
        m = real[0]
        ctx = text[max(0, m.start() - cf.CRED_PREFIX):m.start(1)].strip() + " [REDACTED]"
        return R_PRESENT, len(real), len(placeholders), ctx
    if placeholders:
        return R_REDACTED, 0, len(placeholders), ""
    w = cf.CRED_WORD_RX.search(text)
    if w:
        return R_WORD, 0, 0, text[max(0, w.start() - cf.CRED_PREFIX):w.end()].strip()
    return R_NONE, 0, 0, ""


# ---------------------------------------------------------------- HTTP
def get(session, url, args):
    return session.get(url, params={"_dc": int(time.time() * 1000)}, headers={"Accept": "application/json"},
                       timeout=args.timeout)


def shape(obj, depth=0):
    """A short description of a JSON structure (keys and sizes, no values)."""
    if isinstance(obj, dict):
        return "{" + ", ".join(f"{k}:{shape(v, depth + 1) if depth < 2 else '...'}" for k, v in list(obj.items())[:12]) + "}"
    if isinstance(obj, list):
        return f"[{len(obj)} x {shape(obj[0], depth + 1) if obj and depth < 2 else ''}]"
    return type(obj).__name__


def probe(session, args, tickets, targets):
    ir = args.probe.upper()
    row = tickets[tickets["Record Number"] == ir]
    if row.empty:
        sys.exit(f"{ir} is not in the app extract.")
    item, tid = int(row.iloc[0]["Item ID"]), int(row.iloc[0]["ID"])
    urls = [args.url_template] if args.url_template else CANDIDATES
    os.makedirs(RAW_DIR, exist_ok=True)
    print(f"Probing {ir} (Item ID {item}, ID {tid}). Only status and structure are printed.\n")
    for i, tpl in enumerate(urls):
        url = tpl.format(item=item, id=tid)
        try:
            r = get(session, url, args)
        except Exception as e:  # noqa: BLE001
            print(f"[{i}] {url}\n     ERROR {e}\n")
            continue
        ctype = r.headers.get("Content-Type", "")
        line = f"[{i}] {url}\n     HTTP {r.status_code}, {ctype.split(';')[0]}, {len(r.content)} bytes"
        try:
            payload = r.json()
            text = response_text(payload)
            res = judge(text)[0]
            line += f"\n     structure: {shape(payload)}\n     text length {len(text)}, judged: {res}"
            with open(os.path.join(RAW_DIR, f"probe_{i}.json"), "w", encoding="utf-8") as f:
                json.dump(payload, f)
        except ValueError:
            line += "\n     (not JSON - probably the login page: the cookie is missing or expired)" \
                if "html" in ctype else "\n     (not JSON)"
        print(line + "\n")
    print("Send me this output. The right endpoint returns a long text length (all description entries).")


def fetch(session, args, item, tid):
    cache = os.path.join(RAW_DIR, f"{tid}.json")
    if os.path.exists(cache) and not args.refresh:
        with open(cache, encoding="utf-8") as f:
            return json.load(f)
    url = args.url_template.format(item=item, id=tid) if args.url_template else CANDIDATES[0].format(item=item, id=tid)
    last = None
    for attempt in range(1, args.retries + 1):
        try:
            r = get(session, url, args)
            r.raise_for_status()
            payload = r.json()
            os.makedirs(RAW_DIR, exist_ok=True)
            with open(cache, "w", encoding="utf-8") as f:
                json.dump(payload, f)
            return payload
        except requests.exceptions.SSLError as e:
            raise RuntimeError(f"SSL certificate problem - re-run with --insecure. ({e})")
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(2 ** attempt)
    raise RuntimeError(f"gave up: {last}")


# ---------------------------------------------------------------- run
def load_targets(args):
    if args.tickets:
        return [t.strip().upper() for t in args.tickets.split(",") if t.strip()], {}
    if not os.path.exists(args.report):
        sys.exit(f"{args.report} not found - pull the report or pass --tickets.")
    cr = pd.read_excel(args.report, sheet_name="Credential_Review", dtype=object)
    cr = cr[cr["Assessment"] == cf.V_LIKELY]
    print(f"{len(cr)} tickets with Assessment = {cf.V_LIKELY} on Credential_Review")
    return list(cr["Record Number"]), dict(zip(cr["Record Number"], cr["Assessment"]))


def run(session, args, tickets, targets, db_verdict):
    by_ir = tickets.set_index("Record Number")
    rows, done = [], 0

    def work(ir):
        if ir not in by_ir.index:
            return ir, None, None, "not in the app extract"
        item, tid = int(by_ir.loc[ir, "Item ID"]), int(by_ir.loc[ir, "ID"])
        try:
            payload = fetch(session, args, item, tid)
            time.sleep(args.delay)
            return ir, tid, response_text(payload), None
        except Exception as e:  # noqa: BLE001
            return ir, tid, None, str(e)

    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futures = [ex.submit(work, ir) for ir in targets]
        for fu in as_completed(futures):
            ir, tid, text, err = fu.result()
            done += 1
            if err:
                result, real, red, ctx, n = R_ERROR, 0, 0, err[:200], 0
            else:
                result, real, red, ctx = judge(text)
                n = len(text)
            rows.append({"Record Number": ir, "ID": tid, "Report assessment (DB export)": db_verdict.get(ir, ""),
                         "UI result": result, "Password-like values in UI": real, "[REDACTED] values in UI": red,
                         "UI text length": n, "Masked context": ctx})
            if done % 10 == 0 or done == len(targets):
                print(f"  {done}/{len(targets)} done")
            if err and "login" in err.lower():
                print("Looks like the cookie expired - stopping."); break
    df = pd.DataFrame(rows)
    df["_o"] = df["UI result"].map(R_ORDER)
    df = df.sort_values(["_o", "Record Number"]).drop(columns="_o").reset_index(drop=True)
    return df


def write(df, args):
    summary = df["UI result"].value_counts().rename_axis("UI result").reset_index(name="Tickets")
    summary["_o"] = summary["UI result"].map(R_ORDER)
    summary = summary.sort_values("_o").drop(columns="_o")
    summary.loc[len(summary)] = ["TOTAL", len(df)]
    kwargs = dict(mode="a", if_sheet_exists="replace") if os.path.exists(args.report) else dict(mode="w")
    try:
        with pd.ExcelWriter(args.report, engine="openpyxl", **kwargs) as xw:
            summary.to_excel(xw, sheet_name="UI_Credential_Summary", index=False)
            df.to_excel(xw, sheet_name="UI_Credential_Check", index=False)
    except PermissionError:
        sys.exit(f"Cannot write {args.report} - close it in Excel and run again.")
    print(f"\nAdded sheets UI_Credential_Summary and UI_Credential_Check to {args.report}")
    print(summary.to_string(index=False))


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--probe", metavar="IR-NNNNN", help="try the candidate endpoints on one ticket; print status/structure")
    p.add_argument("--url-template", help="endpoint with {item} and {id} placeholders (default: first candidate)")
    p.add_argument("--tickets", help="comma-separated IR numbers instead of the report's LIKELY list")
    p.add_argument("--report", default=REPORT)
    p.add_argument("--app", help="app extract file (default: newest fixtures/QCFootprints_2*.xlsx)")
    p.add_argument("--auth", choices=["ntlm", "cookie"], default="cookie")
    p.add_argument("--header", action="append")
    p.add_argument("--workers", type=int, default=4)
    p.add_argument("--timeout", type=int, default=60)
    p.add_argument("--retries", type=int, default=3)
    p.add_argument("--delay", type=float, default=0.2)
    p.add_argument("--refresh", action="store_true", help="ignore cached responses")
    p.add_argument("--insecure", action="store_true")
    args = p.parse_args()

    tickets = load_tickets(args.app or newest_app_extract())
    session = build_session(args)
    if args.probe:
        probe(session, args, tickets, [])
        return
    targets, db_verdict = load_targets(args)
    df = run(session, args, tickets, targets, db_verdict)
    write(df, args)


if __name__ == "__main__":
    main()
