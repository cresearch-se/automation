"""
STEP 2 — Compare & Report
==========================
python step2_compare.py --legacy legacy_content.json --gcc gcc_content.json --output diff_report.html

Validation rules:
  1. BODY CONTENT   — legacy .ms-rtestate-field must match GCC .contentView_91b335df
                      <60% similarity → FAIL, 60-90% → WARN, ≥90% → PASS
  2. WORD COUNT     — must match exactly → FAIL if different
  3. PARAGRAPHS     — any paragraph present in legacy but missing in GCC → FAIL
  4. HEADINGS       — exact match required; any heading in legacy missing from GCC → FAIL
  5. LEFT NAV (GCC) — left nav must be PRESENT on every GCC site → FAIL if missing
                      (legacy has no left nav so we don't compare nav content, just presence)
  6. LINKS          — must match exactly → FAIL if different
"""

import argparse, json, html as html_lib, re
from datetime import datetime
from difflib import ndiff


# ── Normalise text for exact comparison ───────────────────────────────────────
# Collapses all whitespace (spaces, newlines, tabs) to a single space.
# Ignores leading/trailing whitespace and the Unicode zero-width space (​)
# that SharePoint sometimes inserts. Every real word must match exactly.
def normalise(text):
    text = text.replace('\u200b', '')          # remove zero-width spaces
    text = re.sub(r'\s+', ' ', text).strip()   # collapse all whitespace
    return text


# ── Core comparison logic ─────────────────────────────────────────────────────
def compare_page(pid, legacy, gcc):
    issues, warnings, info = [], [], []

    # ── Load errors ───────────────────────────────────────────────────────────
    if legacy.get("status") == "error":
        return {"status": "ERR-LEGACY",
                "issues": [f"Legacy page failed to load: {legacy.get('error','')}"],
                "warnings": [], "info": [], "legacy": legacy, "gcc": gcc,
                "body_sim": 0, "diff_lines": [], "missing_paras": [], "missing_headings": []}
    if gcc.get("status") == "error":
        return {"status": "ERR-GCC",
                "issues": [f"GCC page failed to load: {gcc.get('error','')}"],
                "warnings": [], "info": [], "legacy": legacy, "gcc": gcc,
                "body_sim": 0, "diff_lines": [], "missing_paras": [], "missing_headings": []}

    # ── 1. Body content — exact match after whitespace normalisation ──────────
    legacy_norm = normalise(legacy.get("body_text", ""))
    gcc_norm    = normalise(gcc.get("body_text", ""))
    body_exact  = (legacy_norm == gcc_norm)

    if not body_exact:
        # Find the first point of difference to give a useful clue
        min_len = min(len(legacy_norm), len(gcc_norm))
        diff_pos = next((i for i in range(min_len) if legacy_norm[i] != gcc_norm[i]), min_len)
        snippet_l = legacy_norm[max(0, diff_pos-20):diff_pos+60]
        snippet_g = gcc_norm[max(0, diff_pos-20):diff_pos+60]
        issues.append(
            f"Body text does not match exactly after whitespace normalisation. "
            f"First difference at char {diff_pos}. "
            f"Legacy: '{snippet_l}' | GCC: '{snippet_g}'"
        )
    else:
        info.append("Body text matches exactly (whitespace-normalised) ✓")

    # Store for report display (used in similarity column)
    body_sim = 1.0 if body_exact else 0.0

    # ── 2. Word count ─────────────────────────────────────────────────────────
    lw = legacy.get("word_count", 0)
    gw = gcc.get("word_count", 0)
    if lw != gw:
        issues.append(f"Word count does not match — legacy: {lw}, GCC: {gw}")
    else:
        info.append(f"Word count matches — {lw} words ✓")

    # ── 3. Paragraphs — legacy paragraphs must all be present in GCC ──────────
    legacy_paras = set(p.strip() for p in legacy.get("paragraphs", []) if p.strip())
    gcc_paras    = set(p.strip() for p in gcc.get("paragraphs", []) if p.strip())
    missing_paras = sorted(legacy_paras - gcc_paras)
    extra_paras   = sorted(gcc_paras - legacy_paras)

    if missing_paras:
        issues.append(f"{len(missing_paras)} paragraph(s) from legacy are MISSING in GCC")
    else:
        info.append(f"All {len(legacy_paras)} legacy paragraphs found in GCC ✓")
    if extra_paras:
        info.append(f"{len(extra_paras)} new paragraph(s) in GCC not in legacy (may be intentional)")

    # ── 4. Headings — exact match required ───────────────────────────────────
    # Extract headings from body using <strong> + block-level tags
    # The extractor captures them; we compare sets exactly
    legacy_headings = set(h.strip() for h in legacy.get("headings", []) if h.strip())
    gcc_headings    = set(h.strip() for h in gcc.get("headings", []) if h.strip())
    missing_headings = sorted(legacy_headings - gcc_headings)
    extra_headings   = sorted(gcc_headings - legacy_headings)

    if missing_headings:
        issues.append(f"{len(missing_headings)} heading(s) from legacy are MISSING in GCC")
    else:
        info.append(f"All {len(legacy_headings)} legacy headings found in GCC ✓")
    if extra_headings:
        info.append(f"{len(extra_headings)} new heading(s) in GCC not in legacy (may be intentional)")

    # ── 5. Left nav — GCC presence check only ────────────────────────────────
    gcc_left_nav = gcc.get("left_nav", [])
    if gcc_left_nav:
        info.append(f"Left nav present on GCC — {len(gcc_left_nav)} item(s) ✓")
    else:
        issues.append("Left nav MISSING on GCC site")

    # ── 6. Link count (informational) ────────────────────────────────────────
    ll = legacy.get("link_count", 0)
    gl = gcc.get("link_count", 0)
    if ll != gl:
        issues.append(f"Link count does not match — legacy: {ll}, GCC: {gl}")
    else:
        info.append(f"Link count matches — {ll} links ✓")

    # ── Line-level diff for report display ───────────────────────────────────
    legacy_lines = [l for l in legacy.get("body_text", "").splitlines() if l.strip()]
    gcc_lines    = [l for l in gcc.get("body_text", "").splitlines() if l.strip()]
    diff_lines   = [ln for ln in ndiff(legacy_lines, gcc_lines)
                    if ln.startswith(("+", "-")) and not ln.startswith(("+++", "---"))][:100]

    status = "FAIL" if issues else ("WARN" if warnings else "PASS")

    return {
        "status":           status,
        "body_sim":         round(body_sim, 4),
        "issues":           issues,
        "warnings":         warnings,
        "info":             info,
        "diff_lines":       diff_lines,
        "missing_paras":    missing_paras[:10],
        "missing_headings": missing_headings,
        "extra_paras":      extra_paras[:5],
        "extra_headings":   extra_headings,
        "legacy":           legacy,
        "gcc":              gcc,
    }


# ── HTML report ───────────────────────────────────────────────────────────────
def generate_report(results, legacy_meta, gcc_meta, output_path):
    total  = len(results)
    passed = sum(1 for r in results.values() if r["status"] == "PASS")
    warned = sum(1 for r in results.values() if r["status"] == "WARN")
    failed = sum(1 for r in results.values() if r["status"] == "FAIL")
    errors = sum(1 for r in results.values() if "ERR" in r["status"])

    order = {"FAIL": 0, "WARN": 1, "ERR-LEGACY": 2, "ERR-GCC": 3, "PASS": 4}
    sorted_items = sorted(results.items(), key=lambda x: order.get(x[1]["status"], 9))

    def badge(status):
        cfg = {
            "PASS":       ("PASS",  "#1D7A6B", "#E1F5EE"),
            "WARN":       ("WARN",  "#854F0B", "#FAEEDA"),
            "FAIL":       ("FAIL",  "#A32D2D", "#FCEBEB"),
            "ERR-LEGACY": ("ERR-L", "#5A2D82", "#F3E8FF"),
            "ERR-GCC":    ("ERR-G", "#5A2D82", "#F3E8FF"),
        }
        label, fg, bg = cfg.get(status, ("?", "#333", "#eee"))
        return (f'<span style="background:{bg};color:{fg};padding:2px 9px;border-radius:4px;'
                f'font-size:11px;font-weight:700;font-family:monospace">{label}</span>')

    def block(color, border, prefix, text):
        return (f'<div style="background:{color};border-left:4px solid {border};'
                f'padding:7px 12px;margin:3px 0;font-size:13px">'
                f'<b>{prefix}</b> {html_lib.escape(text)}</div>')

    def diff_html(lines):
        if not lines:
            return "<em style='color:#999;font-size:12px'>No text-level differences found</em>"
        out = ['<pre style="font-size:11px;line-height:1.5;white-space:pre-wrap;margin:0;'
               'border:1px solid #e0e0e0;padding:8px;border-radius:4px;background:#fafafa">']
        for ln in lines:
            e = html_lib.escape(ln)
            if ln.startswith("+"):
                out.append(f'<span style="color:#1D7A6B;background:#E1F5EE;display:block">{e}</span>')
            elif ln.startswith("-"):
                out.append(f'<span style="color:#A32D2D;background:#FCEBEB;display:block">{e}</span>')
        out.append("</pre>")
        return "".join(out)

    def list_items(items, label, color):
        if not items: return ""
        lis = "".join(f"<li style='margin:2px 0;font-size:12px'>{html_lib.escape(str(it)[:200])}"
                      f"{'…' if len(str(it))>200 else ''}</li>" for it in items)
        return (f'<div style="margin:6px 0 10px"><b style="font-size:13px;color:{color}">'
                f'{label}</b><ul style="margin:4px 0 0 18px">{lis}</ul></div>')

    rows = []
    for pid, r in sorted_items:
        l = r.get("legacy", {})
        g = r.get("gcc", {})
        body_match = r.get("body_sim", 0) == 1.0
        sim_pct    = ('<span style="color:#1D7A6B;font-weight:600">MATCH</span>' if body_match
                      else '<span style="color:#A32D2D;font-weight:600">DIFF</span>')
        lw, gw    = l.get("word_count", 0), g.get("word_count", 0)
        lp, gp    = l.get("para_count", 0), g.get("para_count", 0)
        lh        = len(set(l.get("headings", [])))
        gh        = len(set(g.get("headings", [])))
        g_nav = g.get("left_nav", [])
        nav_ok = '<span style="color:#1D7A6B;font-weight:600">PRESENT</span>' if g_nav else '<span style="color:#A32D2D;font-weight:700">MISSING</span>'
        n_issues  = len(r["issues"])
        n_warns   = len(r["warnings"])
        issue_col = f'<span style="color:#A32D2D;font-weight:600">{n_issues} fail</span>' if n_issues else "0 fail"
        warn_col  = f'<span style="color:#854F0B">{n_warns} warn</span>' if n_warns else "0 warn"

        # Detail panel
        issue_blocks = "".join(block("#FCEBEB","#A32D2D","FAIL:", i) for i in r["issues"])
        warn_blocks  = "".join(block("#FAEEDA","#854F0B","WARN:", w) for w in r["warnings"])
        info_blocks  = "".join(block("#F0F4FF","#2E5FA3","INFO:", i) for i in r.get("info", []))

        missing_paras_html    = list_items(r.get("missing_paras",[]),    "Paragraphs in legacy missing from GCC:", "#A32D2D")
        missing_headings_html = list_items(r.get("missing_headings",[]), "Headings in legacy missing from GCC:",   "#A32D2D")
        extra_paras_html      = list_items(r.get("extra_paras",[]),      "New paragraphs in GCC (not in legacy):", "#1D7A6B")

        lurl = html_lib.escape(l.get("url",""))
        gurl = html_lib.escape(g.get("url",""))
        lt   = html_lib.escape(l.get("page_title","—"))
        gt   = html_lib.escape(g.get("page_title","—"))
        lnav_items = l.get("left_nav",[])
        gnav_items = g.get("left_nav",[])
        gnav_html = ""
        if lnav_items or gnav_items:
            l_lis = "".join(f"<li style='font-size:12px'>{html_lib.escape(t)}</li>" for t,_ in lnav_items) or "<li style='color:#888;font-size:12px'>No left nav</li>"
            g_lis = "".join(f"<li style='font-size:12px'>{html_lib.escape(t)}</li>" for t,_ in gnav_items) or "<li style='color:#A32D2D;font-size:12px'>MISSING</li>"
            gnav_html = (f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:10px">' 
                        f'<div><b style="font-size:12px">Legacy left nav:</b><ul style="margin:4px 0 0 16px">{l_lis}</ul></div>'
                        f'<div><b style="font-size:12px">GCC left nav:</b><ul style="margin:4px 0 0 16px">{g_lis}</ul></div>'
                        f'</div>')

        rows.append(f"""
        <tr class="mr" onclick="td('{pid}')" style="cursor:pointer"
            data-status="{r['status']}" data-pid="{html_lib.escape(pid.lower())}">
          <td style="font-family:monospace;font-size:12px">{html_lib.escape(pid)}</td>
          <td>{badge(r['status'])}</td>
          <td style="text-align:center">{sim_pct}</td>
          <td style="text-align:center">{lw} → {gw}</td>
          <td style="text-align:center">{lp} → {gp}</td>
          <td style="text-align:center">{lh} → {gh}</td>
          <td style="text-align:center">{nav_ok}</td>
          <td style="font-size:12px">{issue_col} &nbsp; {warn_col}</td>
        </tr>
        <tr id="d-{pid}" class="dr" style="display:none;background:#fafafa">
          <td colspan="8" style="padding:14px 22px">
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:10px;font-size:12px">
              <div><b>Legacy URL:</b><br><a href="{lurl}" target="_blank" style="color:#185FA5;word-break:break-all">{lurl}</a><br><b>Page title:</b> {lt}</div>
              <div><b>GCC URL:</b><br><a href="{gurl}" target="_blank" style="color:#185FA5;word-break:break-all">{gurl}</a><br><b>Page title:</b> {gt}</div>
            </div>
            {issue_blocks}{warn_blocks}{info_blocks}
            {missing_paras_html}{missing_headings_html}{extra_paras_html}
            {gnav_html}
            {'<div style="margin-top:12px"><b style="font-size:13px">Text diff (red = in legacy only, green = in GCC only):</b><div style="margin-top:6px">' + diff_html(r.get("diff_lines",[])) + '</div></div>' if r.get("diff_lines") else ""}
          </td>
        </tr>""")

    generated = datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")

    html_out = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Migration Diff Report</title>
<style>
  * {{ box-sizing:border-box; }}
  body {{ font-family:Arial,sans-serif; margin:0; background:#f4f4f4; color:#1a1a1a; }}
  .hdr {{ background:#1B3A6B; color:#fff; padding:20px 28px; }}
  .hdr h1 {{ margin:0 0 4px; font-size:20px; }}
  .hdr p  {{ margin:0; opacity:.7; font-size:12px; }}
  .summary {{ display:flex; gap:12px; padding:14px 28px; background:#fff; border-bottom:1px solid #e0e0e0; flex-wrap:wrap; align-items:center; }}
  .stat {{ text-align:center; min-width:75px; }}
  .stat-num {{ font-size:26px; font-weight:700; }}
  .stat-label {{ font-size:11px; color:#666; }}
  .rules {{ padding:10px 28px; background:#EEF3FB; border-bottom:1px solid #c8d8f0; font-size:12px; color:#1B3A6B; }}
  .rules b {{ margin-right:4px; }}
  .controls {{ padding:10px 28px; background:#fff; border-bottom:1px solid #e0e0e0; display:flex; gap:8px; align-items:center; flex-wrap:wrap; }}
  .controls input  {{ padding:5px 8px; border:1px solid #ccc; border-radius:4px; font-size:13px; width:220px; }}
  .controls select {{ padding:5px 8px; border:1px solid #ccc; border-radius:4px; font-size:13px; }}
  .controls button {{ padding:5px 10px; border:1px solid #ccc; border-radius:4px; cursor:pointer; font-size:12px; background:#fff; }}
  .wrap {{ padding:16px 28px; }}
  table {{ width:100%; border-collapse:collapse; background:#fff; box-shadow:0 1px 3px rgba(0,0,0,.1); border-radius:6px; overflow:hidden; }}
  th {{ background:#1B3A6B; color:#fff; padding:9px 12px; text-align:left; font-size:12px; white-space:nowrap; }}
  td {{ padding:9px 12px; border-bottom:1px solid #f0f0f0; font-size:12px; vertical-align:middle; }}
  .mr:hover td {{ background:#f0f5ff; }}
  .pass {{ color:#1D7A6B; }} .fail {{ color:#A32D2D; }} .warn {{ color:#854F0B; }}
</style>
</head>
<body>
<div class="hdr">
  <h1>SharePoint → GCC Migration — Content Diff Report</h1>
  <p>Generated: {generated} &nbsp;|&nbsp; Legacy extracted: {legacy_meta.get('extracted_at','?')} &nbsp;|&nbsp; GCC extracted: {gcc_meta.get('extracted_at','?')}</p>
</div>

<div class="summary">
  <div class="stat"><div class="stat-num">{total}</div><div class="stat-label">Total pages</div></div>
  <div class="stat"><div class="stat-num pass">{passed}</div><div class="stat-label">Pass</div></div>
  <div class="stat"><div class="stat-num warn">{warned}</div><div class="stat-label">Warn</div></div>
  <div class="stat"><div class="stat-num fail">{failed}</div><div class="stat-label">Fail</div></div>
  <div class="stat"><div class="stat-num" style="color:#5A2D82">{errors}</div><div class="stat-label">Errors</div></div>
  <div class="stat"><div class="stat-num">{passed*100//total if total else 0}%</div><div class="stat-label">Pass rate</div></div>
</div>

<div class="rules">
  <b>Validation rules:</b>
  Body text exact match (whitespace-normalised) → PASS, any difference → FAIL &nbsp;|&nbsp;
  
  Any legacy paragraph missing from GCC → FAIL &nbsp;|&nbsp;
  Any legacy heading missing from GCC → FAIL &nbsp;|&nbsp;
  Left nav: presence check only — content comparison skipped &nbsp;|&nbsp;
  
</div>

<div class="controls">
  <label>Filter:</label>
  <input type="text" id="searchBox" placeholder="Search page ID..." oninput="filt()">
  <select id="sf" onchange="filt()">
    <option value="">All statuses</option>
    <option value="FAIL">FAIL only</option>
    <option value="WARN">WARN only</option>
    <option value="PASS">PASS only</option>
    <option value="ERR">ERR only</option>
  </select>
  <button onclick="expandAll()">Expand all</button>
  <button onclick="collapseAll()">Collapse all</button>
</div>

<div class="wrap">
<table>
  <thead><tr>
    <th>Page ID</th>
    <th>Status</th>
    <th style="text-align:center">Body text</th>
    <th style="text-align:center">Words (L→G)</th>
    <th style="text-align:center">Paras (L→G)</th>
    <th style="text-align:center">Headings (L→G)</th>
    <th style="text-align:center">GCC left nav</th>
    <th>Issues / Warnings</th>
  </tr></thead>
  <tbody id="tb">{''.join(rows)}</tbody>
</table>
</div>

<script>
function td(id) {{
  const el = document.getElementById('d-' + id);
  el.style.display = el.style.display === 'none' ? 'table-row' : 'none';
}}
function expandAll()  {{ document.querySelectorAll('.dr').forEach(r => r.style.display='table-row'); }}
function collapseAll() {{ document.querySelectorAll('.dr').forEach(r => r.style.display='none'); }}
function filt() {{
  const s  = document.getElementById('searchBox').value.toLowerCase();
  const sf = document.getElementById('sf').value;
  document.querySelectorAll('.mr').forEach(row => {{
    const show = (!s || (row.dataset.pid||'').includes(s)) &&
                 (!sf || (row.dataset.status||'').includes(sf));
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

    print(f"\n  Report saved → {output_path}")
    print(f"  PASS: {passed}  WARN: {warned}  FAIL: {failed}  ERR: {errors}  Total: {total}")
    if failed:
        print(f"\n  Pages with failures:")
        for pid, r in sorted_items:
            if r["status"] == "FAIL":
                print(f"    {pid}: {r['issues'][0]}")


# ── Entry point ───────────────────────────────────────────────────────────────
def run(legacy_path, gcc_path, output_path):
    print(f"\nLoading legacy: {legacy_path}")
    with open(legacy_path, encoding="utf-8") as f: legacy_data = json.load(f)
    print(f"Loading GCC:    {gcc_path}")
    with open(gcc_path,    encoding="utf-8") as f: gcc_data    = json.load(f)

    legacy_pages = legacy_data.get("pages", {})
    gcc_pages    = gcc_data.get("pages", {})
    all_ids      = sorted(set(legacy_pages) | set(gcc_pages))
    print(f"Comparing {len(all_ids)} page pairs...\n")

    results = {}
    for pid in all_ids:
        l = legacy_pages.get(pid, {"status":"error","error":"not in legacy extract","url":""})
        g = gcc_pages.get(pid,    {"status":"error","error":"not in gcc extract",   "url":""})
        results[pid] = compare_page(pid, l, g)
        status = results[pid]["status"]
        print(f"  {pid[:50]:<50}  {status}")

    generate_report(results, legacy_data, gcc_data, output_path)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--legacy", required=True)
    parser.add_argument("--gcc",    required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    run(args.legacy, args.gcc, args.output)
