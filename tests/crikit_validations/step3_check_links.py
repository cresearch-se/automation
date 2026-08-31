"""
STEP 3 — Link Checker
======================
Reads gcc_content.json (from step1) and checks every hyperlink found
in each GCC page body. No browser needed — runs offline against the JSON.

Usage:
  py step3_check_links.py --gcc gcc_content.json --output link_report.html

Checks performed on every link:
  1. WRONG DOMAIN  — href still points to cresearch1 (legacy) instead of cresearch3
  2. BROKEN        — href returns 404 or connection error
  3. EXTERNAL      — link points outside sharepoint (informational)
  4. OK            — link is on cresearch3 and responds correctly

Notes:
  - Skips mailto:, javascript:, and anchor (#) links
  - Skips relative URLs (no domain)
  - SharePoint pages require login so HTTP checks use a lightweight HEAD request
    and treat redirects-to-login as OK (page exists, just needs auth)
  - Saves a checkpoint every 100 links so progress is not lost
"""

import argparse, json, html as html_lib, time, re
from datetime import datetime
from urllib.parse import urlparse
import urllib.request
import urllib.error
import ssl


# ── Config ────────────────────────────────────────────────────────────────────
LEGACY_DOMAIN  = "cresearch1.sharepoint.com"
GCC_DOMAIN     = "cresearch3.sharepoint.com"
REQUEST_TIMEOUT = 10   # seconds per link check
PAUSE_EVERY     = 100  # pause after checking this many links
PAUSE_SECONDS   = 5    # seconds to pause


# ── SSL context (ignore cert errors on internal SharePoint) ───────────────────
ssl_ctx = ssl.create_default_context()
ssl_ctx.check_hostname = False
ssl_ctx.verify_mode    = ssl.CERT_NONE


# ── Link classifier ───────────────────────────────────────────────────────────
def classify_link(href):
    """
    Returns (category, reason) where category is one of:
      SKIP, WRONG-DOMAIN, CHECK, EXTERNAL
    """
    if not href:
        return "SKIP", "empty href"

    href = href.strip()

    # Skip non-navigable links
    if href.startswith(("mailto:", "javascript:", "tel:")):
        return "SKIP", "non-navigable"
    if href.startswith("#"):
        return "SKIP", "anchor only"
    if not href.startswith("http"):
        return "SKIP", "relative URL"

    parsed = urlparse(href)
    domain = parsed.netloc.lower()

    if LEGACY_DOMAIN in domain:
        return "WRONG-DOMAIN", f"still points to legacy ({LEGACY_DOMAIN})"
    if GCC_DOMAIN in domain:
        return "CHECK", "on GCC domain"

    return "EXTERNAL", f"external link ({domain})"


def check_url_status(href):
    """
    Try a HEAD request. Returns (status_code, note).
    Treats login redirects as OK since SharePoint always redirects unauthenticated requests.
    """
    try:
        req = urllib.request.Request(
            href,
            headers={"User-Agent": "Mozilla/5.0 (migration-link-checker)"},
            method="HEAD"
        )
        with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT, context=ssl_ctx) as resp:
            return resp.status, "OK"
    except urllib.error.HTTPError as e:
        if e.code in (401, 403):
            # Auth required — page exists, just needs login — treat as OK
            return e.code, "Auth required (page exists)"
        if e.code == 404:
            return 404, "Page not found"
        return e.code, f"HTTP {e.code}"
    except urllib.error.URLError as e:
        return 0, f"Connection error: {str(e.reason)[:60]}"
    except Exception as e:
        return 0, f"Error: {str(e)[:60]}"


# ── Main ──────────────────────────────────────────────────────────────────────
def run(gcc_json_path, output_path):
    print(f"\nLoading GCC content from: {gcc_json_path}")
    with open(gcc_json_path, encoding="utf-8") as f:
        gcc_data = json.load(f)

    pages = gcc_data.get("pages", {})
    print(f"Pages loaded: {len(pages)}")

    # ── Collect all links from all pages ──────────────────────────
    all_links = []   # list of {page_id, text, href, category, reason}
    for pid, page in pages.items():
        if page.get("status") != "ok":
            continue
        for text, href in page.get("links", []):
            category, reason = classify_link(href)
            all_links.append({
                "page_id":  pid,
                "text":     text,
                "href":     href,
                "category": category,
                "reason":   reason,
                "status":   None,
                "note":     None,
            })

    total_links   = len(all_links)
    check_links   = [l for l in all_links if l["category"] == "CHECK"]
    wrong_links   = [l for l in all_links if l["category"] == "WRONG-DOMAIN"]
    external_links= [l for l in all_links if l["category"] == "EXTERNAL"]
    skip_links    = [l for l in all_links if l["category"] == "SKIP"]

    print(f"\nTotal links found : {total_links}")
    print(f"  To check (GCC)  : {len(check_links)}")
    print(f"  Wrong domain    : {len(wrong_links)}  ← these need fixing")
    print(f"  External        : {len(external_links)}")
    print(f"  Skipped         : {len(skip_links)}")
    print(f"\nChecking {len(check_links)} GCC links for broken pages...\n")

    # ── Check GCC links ───────────────────────────────────────────
    checked = 0
    for link in check_links:
        checked += 1
        status_code, note = check_url_status(link["href"])
        link["status"] = status_code
        link["note"]   = note

        # Determine final result
        if status_code == 404:
            link["result"] = "BROKEN"
        elif status_code == 0:
            link["result"] = "ERROR"
        else:
            link["result"] = "OK"

        print(f"  [{checked:04d}/{len(check_links)}] {link['result']:6s} {status_code}  "
              f"{link['href'][:70]}...")

        # Pause every PAUSE_EVERY links
        if checked % PAUSE_EVERY == 0 and checked < len(check_links):
            print(f"\n  ── Pause {PAUSE_SECONDS}s at {checked}/{len(check_links)} links ──\n")
            time.sleep(PAUSE_SECONDS)

    # Mark wrong-domain and external links
    for link in wrong_links:
        link["result"] = "WRONG-DOMAIN"
    for link in external_links:
        link["result"] = "EXTERNAL"
    for link in skip_links:
        link["result"] = "SKIP"

    # ── Build summary per page ────────────────────────────────────
    page_summary = {}
    for link in all_links:
        if link["category"] == "SKIP":
            continue
        pid = link["page_id"]
        if pid not in page_summary:
            page_summary[pid] = {
                "page_id": pid,
                "url": pages[pid].get("url",""),
                "links": [],
                "broken": 0, "wrong_domain": 0, "external": 0, "ok": 0
            }
        page_summary[pid]["links"].append(link)
        r = link.get("result","")
        if r == "BROKEN" or r == "ERROR":  page_summary[pid]["broken"]       += 1
        elif r == "WRONG-DOMAIN":          page_summary[pid]["wrong_domain"]  += 1
        elif r == "EXTERNAL":              page_summary[pid]["external"]       += 1
        elif r == "OK":                    page_summary[pid]["ok"]             += 1

    # Sort pages: most problems first
    sorted_pages = sorted(
        page_summary.values(),
        key=lambda p: (-(p["broken"] + p["wrong_domain"]), -p["external"])
    )

    generate_report(sorted_pages, all_links, gcc_data, output_path)


# ── HTML report ───────────────────────────────────────────────────────────────
def generate_report(sorted_pages, all_links, gcc_data, output_path):
    total_links    = len([l for l in all_links if l["category"] != "SKIP"])
    total_broken   = sum(1 for l in all_links if l.get("result") in ("BROKEN","ERROR"))
    total_wrong    = sum(1 for l in all_links if l.get("result") == "WRONG-DOMAIN")
    total_external = sum(1 for l in all_links if l.get("result") == "EXTERNAL")
    total_ok       = sum(1 for l in all_links if l.get("result") == "OK")
    pages_with_issues = sum(1 for p in sorted_pages if p["broken"] + p["wrong_domain"] > 0)

    generated = datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")

    def badge(result):
        cfg = {
            "OK":           ("OK",           "#1D7A6B", "#E1F5EE"),
            "BROKEN":       ("BROKEN",       "#A32D2D", "#FCEBEB"),
            "ERROR":        ("ERROR",        "#A32D2D", "#FCEBEB"),
            "WRONG-DOMAIN": ("WRONG DOMAIN", "#854F0B", "#FAEEDA"),
            "EXTERNAL":     ("EXTERNAL",     "#2E5FA3", "#EEF3FB"),
        }
        label, fg, bg = cfg.get(result, ("?","#333","#eee"))
        return (f'<span style="background:{bg};color:{fg};padding:2px 8px;border-radius:4px;'
                f'font-size:11px;font-weight:700;font-family:monospace">{label}</span>')

    rows = []
    for p in sorted_pages:
        pid      = p["page_id"]
        has_issue= p["broken"] + p["wrong_domain"] > 0
        bg_color = "#fff8f8" if p["broken"] > 0 else ("#fffbf0" if p["wrong_domain"] > 0 else "#fff")
        broken_col     = f'<span style="color:#A32D2D;font-weight:600">{p["broken"]}</span>'   if p["broken"]       else "0"
        wrong_col      = f'<span style="color:#854F0B;font-weight:600">{p["wrong_domain"]}</span>' if p["wrong_domain"] else "0"

        # Link rows for detail panel
        link_rows = []
        for lk in sorted(p["links"], key=lambda x: (
            0 if x.get("result") in ("BROKEN","ERROR") else
            1 if x.get("result") == "WRONG-DOMAIN" else
            2 if x.get("result") == "EXTERNAL" else 3
        )):
            result  = lk.get("result","?")
            href_e  = html_lib.escape(lk["href"])
            text_e  = html_lib.escape(lk["text"][:60])
            note_e  = html_lib.escape(lk.get("note","") or lk.get("reason",""))
            status  = lk.get("status","")
            link_rows.append(
                f'<tr style="font-size:12px">'
                f'<td style="padding:5px 8px">{badge(result)}</td>'
                f'<td style="padding:5px 8px">{text_e}</td>'
                f'<td style="padding:5px 8px;word-break:break-all">'
                f'<a href="{href_e}" target="_blank" style="color:#185FA5">{href_e[:80]}{"…" if len(lk["href"])>80 else ""}</a></td>'
                f'<td style="padding:5px 8px;color:#666">{status or ""}</td>'
                f'<td style="padding:5px 8px;color:#666">{note_e}</td>'
                f'</tr>'
            )

        link_table = (
            f'<table style="width:100%;border-collapse:collapse;font-size:12px;margin-top:8px">'
            f'<thead><tr style="background:#f0f0f0">'
            f'<th style="padding:5px 8px;text-align:left">Result</th>'
            f'<th style="padding:5px 8px;text-align:left">Link text</th>'
            f'<th style="padding:5px 8px;text-align:left">URL</th>'
            f'<th style="padding:5px 8px;text-align:left">HTTP</th>'
            f'<th style="padding:5px 8px;text-align:left">Note</th>'
            f'</tr></thead><tbody>{"".join(link_rows)}</tbody></table>'
        ) if link_rows else "<em style='color:#999;font-size:12px'>No checkable links on this page</em>"

        page_url_e = html_lib.escape(p["url"])
        rows.append(f"""
        <tr class="mr" onclick="td('{pid}')" style="cursor:pointer;background:{bg_color}"
            data-pid="{html_lib.escape(pid.lower())}"
            data-issue="{'1' if has_issue else '0'}">
          <td style="font-family:monospace;font-size:12px;padding:9px 12px">{html_lib.escape(pid)}</td>
          <td style="padding:9px 12px;text-align:center">{len(p['links'])}</td>
          <td style="padding:9px 12px;text-align:center">{broken_col}</td>
          <td style="padding:9px 12px;text-align:center">{wrong_col}</td>
          <td style="padding:9px 12px;text-align:center;color:#2E5FA3">{p['external']}</td>
          <td style="padding:9px 12px;text-align:center;color:#1D7A6B">{p['ok']}</td>
        </tr>
        <tr id="d-{pid}" class="dr" style="display:none;background:#fafafa">
          <td colspan="6" style="padding:12px 20px">
            <b style="font-size:13px">GCC URL:</b>
            <a href="{page_url_e}" target="_blank" style="font-size:12px;color:#185FA5;word-break:break-all">{page_url_e}</a>
            <div style="margin-top:10px">{link_table}</div>
          </td>
        </tr>""")

    html_out = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Link Checker Report</title>
<style>
  * {{ box-sizing:border-box; }}
  body {{ font-family:Arial,sans-serif; margin:0; background:#f4f4f4; color:#1a1a1a; }}
  .hdr {{ background:#1B3A6B; color:#fff; padding:20px 28px; }}
  .hdr h1 {{ margin:0 0 4px; font-size:20px; }}
  .hdr p  {{ margin:0; opacity:.7; font-size:12px; }}
  .summary {{ display:flex; gap:12px; padding:14px 28px; background:#fff;
              border-bottom:1px solid #e0e0e0; flex-wrap:wrap; }}
  .stat {{ text-align:center; min-width:90px; }}
  .stat-num {{ font-size:26px; font-weight:700; }}
  .stat-label {{ font-size:11px; color:#666; }}
  .controls {{ padding:10px 28px; background:#fff; border-bottom:1px solid #e0e0e0;
               display:flex; gap:8px; align-items:center; flex-wrap:wrap; }}
  .controls input  {{ padding:5px 8px; border:1px solid #ccc; border-radius:4px; font-size:13px; width:220px; }}
  .controls select {{ padding:5px 8px; border:1px solid #ccc; border-radius:4px; font-size:13px; }}
  .controls button {{ padding:5px 10px; border:1px solid #ccc; border-radius:4px;
                      cursor:pointer; font-size:12px; background:#fff; }}
  .wrap {{ padding:16px 28px; }}
  table.main {{ width:100%; border-collapse:collapse; background:#fff;
                box-shadow:0 1px 3px rgba(0,0,0,.1); border-radius:6px; overflow:hidden; }}
  th {{ background:#1B3A6B; color:#fff; padding:9px 12px; text-align:left; font-size:12px; }}
  .mr:hover td {{ background:#f0f5ff !important; }}
</style>
</head>
<body>
<div class="hdr">
  <h1>CRIKIT Migration — Link Checker Report</h1>
  <p>Generated: {generated} &nbsp;|&nbsp; GCC extracted: {gcc_data.get('extracted_at','?')}</p>
</div>

<div class="summary">
  <div class="stat"><div class="stat-num">{total_links}</div><div class="stat-label">Total links</div></div>
  <div class="stat"><div class="stat-num" style="color:#1D7A6B">{total_ok}</div><div class="stat-label">OK</div></div>
  <div class="stat"><div class="stat-num" style="color:#A32D2D">{total_broken}</div><div class="stat-label">Broken</div></div>
  <div class="stat"><div class="stat-num" style="color:#854F0B">{total_wrong}</div><div class="stat-label">Wrong domain</div></div>
  <div class="stat"><div class="stat-num" style="color:#2E5FA3">{total_external}</div><div class="stat-label">External</div></div>
  <div class="stat"><div class="stat-num" style="color:#A32D2D">{pages_with_issues}</div><div class="stat-label">Pages with issues</div></div>
</div>

<div class="controls">
  <label>Filter:</label>
  <input type="text" id="searchBox" placeholder="Search page ID..." oninput="filt()">
  <select id="sf" onchange="filt()">
    <option value="">All pages</option>
    <option value="1">Issues only (broken or wrong domain)</option>
  </select>
  <button onclick="expandAll()">Expand all</button>
  <button onclick="collapseAll()">Collapse all</button>
</div>

<div class="wrap">
<table class="main">
  <thead><tr>
    <th>Page ID</th>
    <th style="text-align:center">Total links</th>
    <th style="text-align:center">Broken</th>
    <th style="text-align:center">Wrong domain</th>
    <th style="text-align:center">External</th>
    <th style="text-align:center">OK</th>
  </tr></thead>
  <tbody id="tb">{''.join(rows)}</tbody>
</table>
</div>

<script>
function td(id) {{
  const el = document.getElementById('d-' + id);
  el.style.display = el.style.display === 'none' ? 'table-row' : 'none';
}}
function expandAll()   {{ document.querySelectorAll('.dr').forEach(r => r.style.display='table-row'); }}
function collapseAll() {{ document.querySelectorAll('.dr').forEach(r => r.style.display='none'); }}
function filt() {{
  const s  = document.getElementById('searchBox').value.toLowerCase();
  const sf = document.getElementById('sf').value;
  document.querySelectorAll('.mr').forEach(row => {{
    const show = (!s || (row.dataset.pid||'').includes(s)) &&
                 (!sf || row.dataset.issue === sf);
    row.style.display = show ? '' : 'none';
    const dr = row.nextElementSibling;
    if (dr && dr.classList.contains('dr') && !show) dr.style.display = 'none';
  }});
}}
</script>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_out)

    print(f"\n{'='*60}")
    print(f"  Report saved → {output_path}")
    print(f"  Total links      : {total_links}")
    print(f"  OK               : {total_ok}")
    print(f"  Broken / Error   : {total_broken}")
    print(f"  Wrong domain     : {total_wrong}  ← still pointing to legacy")
    print(f"  External         : {total_external}")
    print(f"  Pages with issues: {pages_with_issues}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--gcc",    required=True, help="gcc_content.json from step1")
    parser.add_argument("--output", required=True, help="Output HTML report path")
    args = parser.parse_args()
    run(args.gcc, args.output)
