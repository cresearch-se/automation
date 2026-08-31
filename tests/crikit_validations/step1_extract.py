"""
STEP 1 — Content Extractor (Selenium + Edge, with resume support)
=================================================================
Run TWICE:
  py step1_extract.py --env legacy --urls urls.csv --output legacy_content.json
  py step1_extract.py --env gcc    --urls urls.csv --output gcc_content.json

Resume support:
  If the script was stopped midway, just re-run the same command.
  It will read the existing JSON, skip pages already extracted,
  and continue from where it left off.

  To start completely fresh, delete the output JSON file first.

Checkpointing:
  Saves progress every 100 pages so nothing is lost if interrupted.
  Pauses 10 seconds every 100 pages to avoid SSO timeout.
"""

import argparse, csv, json, time, os, re
from datetime import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.common.exceptions import WebDriverException, TimeoutException


PAGE_LOAD_TIMEOUT = 30
WAIT_AFTER_LOAD   = 3
BATCH_SIZE        = 100
BATCH_PAUSE       = 10


# ── Extractors ────────────────────────────────────────────────────────────────

def extract_legacy(html, url):
    soup = BeautifulSoup(html, "lxml")
    pt_el = soup.find(id="DeltaPlaceHolderPageTitleInTitleArea")
    page_title = pt_el.get_text(strip=True) if pt_el else ""
    rte = soup.find(class_="ms-rtestate-field") or soup.find(id="DeltaPlaceHolderMain")
    body_text, paragraphs, headings, links = "", [], [], []
    if rte:
        for tag in rte(["script", "style", "nav"]): tag.decompose()
        body_text  = rte.get_text(separator=" ", strip=True)
        paragraphs = [p.get_text(strip=True) for p in rte.find_all("p") if p.get_text(strip=True)]
        headings   = [h.get_text(strip=True) for h in rte.find_all(["h1","h2","h3","h4","strong"])
                      if h.get_text(strip=True) and len(h.get_text(strip=True)) < 120]
        links      = [(a.get_text(strip=True), a.get("href",""))
                      for a in rte.find_all("a", href=True) if a.get_text(strip=True)]
    nav_el   = soup.find(id="DeltaPlaceHolderLeftNavBar")
    left_nav = [(a.get_text(strip=True), a.get("href",""))
                for a in nav_el.find_all("a") if a.get_text(strip=True)] if nav_el else []
    return {
        "url": url, "env": "legacy", "page_title": page_title,
        "body_text": body_text, "paragraphs": paragraphs,
        "headings": headings, "links": links, "left_nav": left_nav,
        "word_count": len(body_text.split()),
        "para_count": len(paragraphs), "link_count": len(links),
    }


def extract_gcc(html, url):
    soup = BeautifulSoup(html, "lxml")
    title_el   = soup.find(attrs={"data-automationid": "SiteHeaderTitle"})
    page_title = title_el.get_text(strip=True) if title_el else ""
    content_el = soup.find(class_="contentView_91b335df")
    body_text, paragraphs, headings, links = "", [], [], []
    if content_el:
        for tag in content_el(["script", "style"]): tag.decompose()
        body_text  = content_el.get_text(separator=" ", strip=True)
        paragraphs = [p.get_text(strip=True) for p in content_el.find_all("p") if p.get_text(strip=True)]
        headings   = [h.get_text(strip=True) for h in content_el.find_all(["h1","h2","h3","h4","strong"])
                      if h.get_text(strip=True) and len(h.get_text(strip=True)) < 120]
        links      = [(a.get_text(strip=True), a.get("href",""))
                      for a in content_el.find_all("a", href=True) if a.get_text(strip=True)]
    nav_el   = soup.find("nav", attrs={"aria-label": "Site"})
    left_nav = [(a.get_text(strip=True), a.get("href",""))
                for a in nav_el.find_all("a") if a.get_text(strip=True)] if nav_el else []
    hub_el    = soup.find(id="HubNavTitle")
    hub_title = hub_el.get_text(strip=True) if hub_el else ""
    return {
        "url": url, "env": "gcc", "page_title": page_title,
        "body_text": body_text, "paragraphs": paragraphs,
        "headings": headings, "links": links, "left_nav": left_nav,
        "hub_title": hub_title,
        "word_count": len(body_text.split()),
        "para_count": len(paragraphs), "link_count": len(links),
    }


# ── URL loader ────────────────────────────────────────────────────────────────

def load_urls(csv_path, env):
    rows = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fieldnames = [h.strip() for h in reader.fieldnames]

        col_map = {}
        for h in fieldnames:
            hl = h.lower().replace(' ', '')
            if 'pageid' in hl:
                col_map['page_id'] = h
            elif 'source' in hl or 'legacy' in hl:
                col_map['legacy_url'] = h
            elif 'gcc' in hl:
                col_map['gcc_url'] = h
            elif 'commertial' in hl or 'commercial' in hl or 'siteurl' in hl:
                col_map['page_id'] = h

        url_col = col_map.get('legacy_url') if env == 'legacy' else col_map.get('gcc_url')

        seen_ids = {}
        for row in reader:
            rel_path = row.get(col_map.get('page_id', ''), '').strip()
            url      = row.get(url_col, '').strip() if url_col else ''
            if not url or not url.startswith('http'):
                continue
            name = os.path.basename(rel_path).replace('.aspx','').replace('\u200b','').strip()
            name = re.sub(r'[^\w\s-]', '', name)
            pid  = re.sub(r'\s+', '_', name)[:80] or f"page_{len(rows)+1}"
            if pid in seen_ids:
                seen_ids[pid] += 1
                pid = f"{pid}_{seen_ids[pid]}"
            else:
                seen_ids[pid] = 1
            rows.append({"page_id": pid, "url": url})
    return rows


# ── Save helper ───────────────────────────────────────────────────────────────

def save_output(output_path, env, total_pages, results, errors):
    data = {
        "env":          env,
        "extracted_at": datetime.utcnow().isoformat(),
        "total_pages":  total_pages,
        "progress":     len(results),
        "success":      sum(1 for v in results.values() if v.get("status") == "ok"),
        "error_count":  len(errors),
        "errors":       errors,
        "pages":        results,
    }
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ── Main ──────────────────────────────────────────────────────────────────────

def run(env, urls_csv, output_path):
    all_pages = load_urls(urls_csv, env)
    total     = len(all_pages)
    extractor = extract_legacy if env == "legacy" else extract_gcc

    print(f"\n{'='*60}")
    print(f"  Extracting {env.upper()} — {total} pages total")
    print(f"{'='*60}\n")

    # ── Resume: load existing output if present ───────────────────
    results, errors = {}, []
    if os.path.exists(output_path):
        try:
            with open(output_path, encoding="utf-8") as f:
                existing = json.load(f)
            results = existing.get("pages", {})
            errors  = existing.get("errors", [])
            done    = sum(1 for v in results.values() if v.get("status") == "ok")
            print(f"  Found existing file: {output_path}")
            print(f"  Already extracted  : {done} pages — these will be skipped")
        except Exception:
            print("  Could not read existing file — starting fresh.")
            results, errors = {}, []

    # Skip pages already successfully extracted
    remaining = [e for e in all_pages
                 if e["page_id"] not in results
                 or results[e["page_id"]].get("status") != "ok"]
    skipped   = total - len(remaining)

    if skipped:
        print(f"  Skipping           : {skipped} already-done pages")
    print(f"  Still to extract   : {len(remaining)} pages")

    if not remaining:
        print(f"\n  All {total} pages already extracted — nothing to do.")
        print(f"  Delete {output_path} and re-run to start fresh.\n")
        return

    # ── Launch Edge ───────────────────────────────────────────────
    opts = Options()
    opts.binary_location = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    opts.add_argument("--start-maximized")
    opts.add_experimental_option("detach", True)

    script_dir  = os.path.dirname(os.path.abspath(__file__))
    driver_path = os.path.join(script_dir, "msedgedriver.exe")
    service     = Service(executable_path=driver_path)

    print("\n  Starting Microsoft Edge...")
    try:
        driver = webdriver.Edge(service=service, options=opts)
    except Exception as e:
        print(f"\n  ERROR starting Edge: {e}")
        print("  1. Make sure msedgedriver.exe is in this folder")
        print("  2. Make sure Edge version matches driver version")
        return

    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)

    # ── Manual login ──────────────────────────────────────────────
    print("\n" + "="*60)
    print(f"  ACTION REQUIRED — {env.upper()} LOGIN")
    print("="*60)
    print(f"  Log in with your {env.upper()} account and complete any MFA.")
    print(f"  Once you can see a SharePoint page, press ENTER here.")
    print("="*60)
    try:
        driver.get(remaining[0]["url"])
    except Exception:
        pass
    input("\n  >> Press ENTER when logged in and ready...\n")
    print(f"  Starting extraction of {len(remaining)} remaining pages...\n")

    # ── Extract ───────────────────────────────────────────────────
    for i, entry in enumerate(remaining):
        pid, url = entry["page_id"], entry["url"]
        overall  = skipped + i + 1
        print(f"  [{overall:04d}/{total}] {pid}  {url[:65]}...")

        try:
            driver.get(url)
            time.sleep(WAIT_AFTER_LOAD)
            html = driver.page_source
            data = extractor(html, url)
            results[pid] = {
                "status": "ok",
                "extracted_at": datetime.utcnow().isoformat(),
                **data
            }
            print(f"           OK — {data['word_count']} words, "
                  f"{data['para_count']} paras, {data['link_count']} links")

        except TimeoutException:
            print(f"           TIMEOUT — skipped")
            errors.append({"page_id": pid, "url": url, "error": "timeout"})
            results[pid] = {"status": "error", "error": "timeout", "url": url}

        except WebDriverException as e:
            msg = str(e).splitlines()[0]
            print(f"           ERROR — {msg}")
            errors.append({"page_id": pid, "url": url, "error": msg})
            results[pid] = {"status": "error", "error": msg, "url": url}

        # ── Checkpoint every BATCH_SIZE pages ─────────────────────
        if (i + 1) % BATCH_SIZE == 0 and (i + 1) < len(remaining):
            save_output(output_path, env, total, results, errors)
            pct = overall * 100 // total
            print(f"\n  ── Checkpoint saved — {overall}/{total} pages ({pct}%) ──")
            print(f"     Pausing {BATCH_PAUSE}s...\n")
            time.sleep(BATCH_PAUSE)

    driver.quit()

    save_output(output_path, env, total, results, errors)
    success = sum(1 for v in results.values() if v.get("status") == "ok")

    print(f"\n{'='*60}")
    print(f"  Done! Saved to: {output_path}")
    print(f"  Success: {success}  |  Errors: {len(errors)}  |  Total: {total}")
    print(f"{'='*60}\n")

    if errors:
        print("  Pages with errors:")
        for e in errors:
            print(f"    [{e['error']}] {e['page_id']}")
            print(f"    {e['url']}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--env",    required=True, choices=["legacy", "gcc"])
    parser.add_argument("--urls",   required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    run(args.env, args.urls, args.output)
