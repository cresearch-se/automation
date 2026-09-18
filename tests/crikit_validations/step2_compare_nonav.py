"""
STEP 2 — Compare & Report (No Nav version)
===========================================
python step2_compare_nonav.py --legacy legacy_content.json --gcc gcc_content.json --output diff_report.html

Validation rules:
  1. BODY CONTENT   — must match exactly (whitespace normalised) → FAIL if different
  2. WORD COUNT     — must match exactly → FAIL if different
  3. PARAGRAPHS     — every legacy paragraph must exist in GCC → FAIL if any missing
  4. HEADINGS       — every legacy heading must exist in GCC (exact text) → FAIL if any missing
  5. LEFT NAV       — content compared if both sides have nav (MATCHING/MISMATCH)
                      MISSING if GCC has no nav, PRESENT if only GCC has it
                      NEVER fails overall status — informational only
  6. LINKS          — real links must match; ignored: file://, LAN paths, intranet URLs
                      relative→absolute URL format changes are treated as grey (not a fail)

Report columns:
  Status     — PASS / WARN / FAIL / ERR
  Body text  — MATCH / DIFF
  Words      — legacy → GCC word count
  Paras      — legacy → GCC paragraph count
  Headings   — legacy → GCC heading count
  Left nav   — PRESENT / MISSING (informational only)
  Links      — PASSED / GREY PASSED / FAILED / N/A
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

    # ── 5. Left nav — compare if both present, never fails overall status ───
    legacy_left_nav = legacy.get("left_nav", [])
    gcc_left_nav    = gcc.get("left_nav", [])
    legacy_nav_set  = set(t.strip() for t,_ in legacy_left_nav if t.strip())
    gcc_nav_set     = set(t.strip() for t,_ in gcc_left_nav    if t.strip())

    if legacy_nav_set and gcc_nav_set:
        missing_from_gcc = sorted(legacy_nav_set - gcc_nav_set)
        if missing_from_gcc:
            info.append(f"Left nav MISMATCH — {len(missing_from_gcc)} item(s) in legacy missing from GCC")
        else:
            info.append(f"Left nav MATCHING — {len(gcc_nav_set)} item(s) ✓")
    elif not legacy_nav_set and gcc_nav_set:
        info.append(f"Left nav PRESENT in GCC only — legacy had none")
    elif legacy_nav_set and not gcc_nav_set:
        info.append("Left nav MISSING on GCC site (informational only)")
    else:
        info.append("No left nav on either side (N/A)")

    # ── 6. Links — count + detail comparison ─────────────────────────────────
    legacy_links = legacy.get("links", [])
    gcc_links    = gcc.get("links", [])
    ll = len(legacy_links)
    gl = len(gcc_links)
    def is_ignorable_link(href):
        """Links expected to be missing in GCC — LAN/file/intranet paths."""
        h = href.strip().lower()
        return (h.startswith("file://") or "http://crikit/" in h or "//crikit/" in h
                or h.replace("/","").replace("\\","").startswith("file:"))

    def get_filename(href):
        """Extract and fully decode page filename + query string for comparison.
        Handles absolute URLs, relative paths, all URL encoding variants,
        strips #anchor fragments.
        
        Special cases:
        - DispForm.aspx?ID=N  → includes ID value so different items are distinct
        - xlviewer.aspx?id=.. → uses only the filename at end of the ?id= path
                                 (the path prefix differs between legacy and GCC)
        - Semicolon (;) in URL → encode before parsing (urlparse splits on ; as path sep)
        - All others          → filename only, no query string
        """
        import os as _os, re as _re
        from urllib.parse import urlparse, unquote, parse_qs
        href = href.strip()
        # Strip #anchor fragment first
        if '#' in href:
            href = href.split('#')[0]
        if not href:
            return ''
        # Encode semicolons before parsing — urlparse treats ; as path separator
        # which truncates filenames like "RE; New Pricing..." to just "RE"
        href = href.replace(';', '%3B')
        parsed   = urlparse(href)
        path     = parsed.path if parsed.scheme else href.split('?')[0]
        filename = _os.path.basename(unquote(path)).lower().rstrip('/')
        # Remove .aspx extension
        if filename.endswith('.aspx'):
            filename = filename[:-5]
        query = parsed.query
        if query:
            qs = parse_qs(query)
            if filename == 'dispform':
                # DispForm.aspx?ID=11 — include ID to distinguish different items
                id_val = qs.get('ID', qs.get('id', ['']))[0]
                filename = filename + 'id' + id_val.lower()
            elif filename == 'xlviewer':
                # xlviewer.aspx?id=/path/to/file.xls — use only the file at end of path
                id_path = qs.get('id', qs.get('ID', ['']))[0]
                filename = _os.path.basename(unquote(id_path)).lower()
            # All other query strings ignored — match on filename only
        # Strip CRIKIT_ prefix added during GCC migration
        # e.g. CRIKIT_UKEUPracticeGroup → UKEUPracticeGroup
        filename = _re.sub(r'^crikit', '', filename)
        # Strip CRIKIT_ prefix added in GCC site names during migration
        filename = filename.replace('crikit', '')
        # Normalise: remove all spaces, brackets, punctuation
        filename = _re.sub(r'[^a-z0-9]', '', filename)
        return filename

    # Build comparable sets (exclude ignorable links)
    legacy_real = {(t.strip(), h.strip()) for t,h in legacy_links
                   if t.strip() and not is_ignorable_link(h)}
    gcc_real    = {(t.strip(), h.strip()) for t,h in gcc_links
                   if t.strip() and not is_ignorable_link(h)}

    exact_missing = legacy_real - gcc_real
    exact_extra   = gcc_real    - legacy_real

    # Match by same link text + same filename (relative vs absolute URL pattern)
    url_changed  = []
    still_extra  = set(exact_extra)
    still_missing = set()

    def normalise_text(t):
        """Normalise link text for comparison — strip punctuation, spaces,
        curly quotes, apostrophes so minor formatting differences don't cause mismatches."""
        import re as _re
        t = t.lower().strip()
        # Replace curly quotes, apostrophes, dashes with plain equivalents
        t = t.replace('‘','').replace('’','').replace('“','').replace('”','')
        t = t.replace('–','-').replace('—','-').replace("'","").replace("`","")
        # Remove all non-alphanumeric characters and collapse spaces
        t = _re.sub(r'[^a-z0-9 ]', '', t)
        t = _re.sub(r' +', ' ', t).strip()
        return t

    # Build a lookup of GCC filenames → GCC entries for fast matching
    # Multiple legacy links can point to the same GCC page (e.g. TOP OF PAGE)
    gcc_file_map = {}
    for g_entry in exact_extra:
        g_text, g_href = g_entry
        g_file = get_filename(g_href)
        if g_file:
            if g_file not in gcc_file_map:
                gcc_file_map[g_file] = []
            gcc_file_map[g_file].append(g_entry)

    def normalise_text(t):
        """Normalise link text for fallback matching when filename is empty."""
        import re as _re
        t = t.lower().strip()
        t = t.replace('‘','').replace('’','').replace('“','').replace('”','')
        t = t.replace("'","").replace("`","")
        t = _re.sub(r'[^a-z0-9 ]', '', t)
        t = _re.sub(r' +', ' ', t).strip()
        return t

    # Build text-based lookup for fallback when filename key is empty
    # (handles anchor-only links, root paths, domain-only external URLs)
    gcc_text_map = {}
    for g_entry in exact_extra:
        g_norm = normalise_text(g_entry[0])
        if g_norm:
            gcc_text_map.setdefault(g_norm, []).append(g_entry)

    for l_entry in exact_missing:
        l_text, l_href = l_entry
        l_file = get_filename(l_href)
        l_norm = normalise_text(l_text)
        if l_file and l_file in gcc_file_map and gcc_file_map[l_file]:
            # Match by filename — same page, URL format changed
            g_entry = gcc_file_map[l_file][0]
            url_changed.append((l_entry, g_entry))
        elif not l_file and l_norm and l_norm in gcc_text_map:
            # Filename is empty (anchor, root path, domain-only) — match by link text
            g_entry = gcc_text_map[l_norm][0]
            url_changed.append((l_entry, g_entry))
        else:
            still_missing.add(l_entry)

    # Extra links in GCC that have no legacy equivalent at all
    matched_gcc_files = set(get_filename(g[1]) for _,g in url_changed)
    still_extra = {g_entry for g_entry in exact_extra
                   if get_filename(g_entry[1]) not in matched_gcc_files}

    missing_links  = sorted(still_missing)
    extra_links    = sorted(still_extra)
    legacy_ignored = [(t,h) for t,h in legacy_links if is_ignorable_link(h)]

    if missing_links:
        issues.append(f"Link mismatch — legacy: {ll}, GCC: {gl} "
                      f"({len(missing_links)} links missing from GCC, "
                      f"{len(extra_links)} extra in GCC)")
    else:
        info.append(f"Links match — {len(legacy_real)} comparable links ✓")
    if url_changed:
        info.append(f"{len(url_changed)} link(s) URL format changed "
                    f"(relative→absolute, same page) — shown in grey")
    if legacy_ignored:
        info.append(f"Ignored {len(legacy_ignored)} LAN/file/intranet links "
                    f"(expected to be missing in GCC)")

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
        "missing_links":    missing_links[:30],
        "url_changed":      url_changed[:30],
        "extra_links":      extra_links[:30],
        "legacy_ignored":   legacy_ignored[:30],
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
        l_nav_set = set(t.strip() for t,_ in l.get("left_nav",[]) if t.strip())
        g_nav_set = set(t.strip() for t,_ in g.get("left_nav",[]) if t.strip())
        if l_nav_set and g_nav_set:
            nav_ok      = ('<span style="color:#1D7A6B;font-weight:600">MATCHING</span>'
                           if not (l_nav_set - g_nav_set)
                           else '<span style="color:#A32D2D;font-weight:600">MISMATCH</span>')
            nav_data    = "MATCHING" if not (l_nav_set - g_nav_set) else "MISMATCH"
        elif not l_nav_set and g_nav_set:
            nav_ok   = '<span style="color:#888;font-weight:600">PRESENT</span>'
            nav_data = "PRESENT"
        elif l_nav_set and not g_nav_set:
            nav_ok   = '<span style="color:#A32D2D;font-weight:600">MISSING</span>'
            nav_data = "MISSING"
        else:
            nav_ok   = '<span style="color:#888">N/A</span>'
            nav_data = "N/A"
        # Links column
        ml_count = len(r.get("missing_links", []))
        uc_count = len(r.get("url_changed", []))
        ll_count = len(l.get("links", []))
        if ml_count > 0:
            links_col = '<span style="color:#A32D2D;font-weight:600">FAILED</span>'
            links_data = "FAILED"
        elif uc_count > 0 and ml_count == 0:
            links_col = '<span style="color:#888;font-weight:600">PASSED</span>'
            links_data = "PASSED"
        elif ll_count == 0:
            links_col = '<span style="color:#888">N/A</span>'
            links_data = "N/A"
        else:
            links_col = '<span style="color:#1D7A6B;font-weight:600">PASSED</span>'
            links_data = "PASSED"

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
        # Link diff detail
        ml = r.get("missing_links", [])
        el = r.get("extra_links", [])
        l_all_links = l.get("links", [])
        g_all_links = g.get("links", [])
        l_link_set  = {(t.strip(), h.strip()) for t,h in l_all_links if t.strip()}
        g_link_set  = {(t.strip(), h.strip()) for t,h in g_all_links if t.strip()}

        def is_ignorable(href):
            h = href.strip().lower()
            return (h.startswith("file://") or "http://crikit/" in h or "//crikit/" in h
                    or h.replace("/","").replace("\\","").startswith("file:"))

        import os as _os, re as _re
        from urllib.parse import urlparse as _urlparse, unquote as _unquote
        def get_fn(href):
            from urllib.parse import parse_qs as _parse_qs
            href = href.strip()
            if '#' in href:
                href = href.split('#')[0]
            if not href:
                return ''
            href = href.replace(';', '%3B')
            parsed   = _urlparse(href)
            path     = parsed.path if parsed.scheme else href.split('?')[0]
            filename = _os.path.basename(_unquote(path)).lower().rstrip('/')
            if filename.endswith('.aspx'):
                filename = filename[:-5]
            query = parsed.query
            if query:
                qs = _parse_qs(query)
                if filename == 'dispform':
                    id_val = qs.get('ID', qs.get('id', ['']))[0]
                    filename = filename + 'id' + id_val.lower()
                elif filename == 'xlviewer':
                    id_path = qs.get('id', qs.get('ID', ['']))[0]
                    filename = _os.path.basename(_unquote(id_path)).lower()
            filename = _re.sub(r'^crikit', '', filename)
            return _re.sub(r'[^a-z0-9]', '', filename)

        uc_legacy = {tuple(le) for le,ge in r.get("url_changed",[])}
        uc_gcc    = {tuple(ge) for le,ge in r.get("url_changed",[])}
        # Filenames that were matched (URL format changed) — GCC links with these are not "new"
        matched_fns = {get_fn(ge[1]) for _,ge in r.get("url_changed",[])}

        def link_row(t, h, highlight="", faded=False, note=""):
            bg       = f'background:{highlight};' if highlight else ""
            color    = "color:#aaa;" if faded else ""
            note_html= f' <span style="font-size:10px;color:#888">{note}</span>' if note else ""
            return (f'<tr style="{bg}">'
                    f'<td style="padding:4px 8px;font-size:12px;{color};width:30%;min-width:120px">{html_lib.escape(t)}{note_html}</td>'
                    f'<td style="padding:4px 8px;font-size:11px;color:#555;word-break:break-all;{color};width:70%">'
                    f'<a href="{html_lib.escape(h)}" target="_blank" style="color:#185FA5;text-decoration:none">{html_lib.escape(h)}</a></td>'
                    f'</tr>')

        # Legacy links table
        if l_all_links:
            l_rows = "".join(
                link_row(t, h,
                    highlight="" if (is_ignorable(h) or (t.strip(),h.strip()) in uc_legacy) else
                              ("#FCEBEB" if (t.strip(),h.strip()) in (l_link_set - g_link_set) else ""),
                    faded=(is_ignorable(h) or (t.strip(),h.strip()) in uc_legacy),
                    note="(URL format changed)" if (t.strip(),h.strip()) in uc_legacy else
                         "(LAN/file - ignored)" if is_ignorable(h) else ""
                )
                for t,h in l_all_links if t.strip()
            )
            legacy_links_table = (
                f'<table style="width:100%;border-collapse:collapse;border:1px solid #e0e0e0;margin-top:4px">' +
                f'<tr style="background:#f0f0f0"><th style="padding:4px 8px;text-align:left;font-size:12px">Link text</th>' +
                f'<th style="padding:4px 8px;text-align:left;font-size:12px">URL</th></tr>' +
                l_rows + '</table>'
            )
        else:
            legacy_links_table = "<em style='font-size:12px;color:#999'>No links on this page</em>"

        # GCC links table — highlight ones not in legacy
        if g_all_links:
            g_rows = "".join(
                link_row(t, h,
                    highlight="#FCEBEB" if (t.strip(),h.strip()) in (g_link_set - l_link_set) and
                                          get_fn(t.strip()) not in matched_fns else "",
                    note="(not in legacy)" if (t.strip(),h.strip()) in (g_link_set - l_link_set) and
                                              get_fn(t.strip()) not in matched_fns else ""
                )
                for t,h in g_all_links if t.strip()
            )
            gcc_links_table = (
                f'<table style="width:100%;border-collapse:collapse;border:1px solid #e0e0e0;margin-top:4px">' +
                f'<tr style="background:#f0f0f0"><th style="padding:4px 8px;text-align:left;font-size:12px">Link text</th>' +
                f'<th style="padding:4px 8px;text-align:left;font-size:12px">URL</th></tr>' +
                g_rows + '</table>'
            )
        else:
            gcc_links_table = "<em style='font-size:12px;color:#999'>No links on this page</em>"

        if l_all_links or g_all_links:
            missing_links_html = (
                f'<div style="margin-top:12px">' +
                f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">' +
                f'<div><b style="font-size:13px">Legacy links ({len(l_all_links)}) ' +
                f'<span style="color:#A32D2D;font-size:11px">red = missing in GCC</span></b>' +
                legacy_links_table + '</div>' +
                f'<div><b style="font-size:13px">GCC links ({len(g_all_links)}) ' +
                f'<span style="color:#A32D2D;font-size:11px">red = in GCC but not in legacy</span></b>' +
                gcc_links_table + '</div>' +
                '</div></div>'
            )
        else:
            missing_links_html = ""
        extra_links_html = ""

        # Nav — presence only, no detail comparison
        gnav_html = ""

        rows.append(f"""
        <tr class="mr" onclick="td('{pid}')" style="cursor:pointer"
            data-status="{r['status']}" data-pid="{html_lib.escape(pid.lower())}"
            data-body="{('MATCH' if r.get('body_sim',0)==1.0 else 'DIFF')}"
            data-nav="{nav_data}"
            data-links="{links_data}">
          <td style="font-size:11px;word-break:break-all;max-width:220px"><a href="{lurl}" target="_blank" style="color:#185FA5">{html_lib.escape(l.get('url',''))[:80]}{"..." if len(l.get("url",""))>80 else ""}</a></td>
          <td style="font-size:11px;word-break:break-all;max-width:220px"><a href="{gurl}" target="_blank" style="color:#185FA5">{html_lib.escape(g.get('url',''))[:80]}{"..." if len(g.get("url",""))>80 else ""}</a></td>
          <td>{badge(r['status'])}</td>
          <td style="text-align:center">{sim_pct}</td>
          <td style="text-align:center;font-size:11px;line-height:1.8">W: {lw}→{gw}<br>P: {lp}→{gp}<br>H: {lh}→{gh}</td>
          <td style="text-align:center">{nav_ok}</td>
          <td style="text-align:center">{links_col}</td>
          <td style="font-size:12px">{issue_col} &nbsp; {warn_col}</td>
        </tr>
        <tr id="d-{pid}" class="dr" style="display:none;background:#fafafa">
          <td colspan="8" style="padding:14px 22px">
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:10px;font-size:12px">
              <div><b>Legacy URL:</b><br><a href="{lurl}" target="_blank" style="color:#185FA5;word-break:break-all">{lurl}</a><br><b>Page title:</b> {lt}</div>
              <div><b>GCC URL:</b><br><a href="{gurl}" target="_blank" style="color:#185FA5;word-break:break-all">{gurl}</a><br><b>Page title:</b> {gt}</div>
            </div>
            {issue_blocks}{warn_blocks}{info_blocks}
            {missing_paras_html}{missing_headings_html}{extra_paras_html}{missing_links_html}{extra_links_html}
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
  <ul style="margin:6px 0 2px 18px;padding:0;font-size:12px;line-height:1.7">
    <li><b>Body text</b> — must match exactly (whitespace ignored) → FAIL if different</li>
    <li><b>Word count</b> — must match exactly → FAIL if different</li>
    <li><b>Paragraphs</b> — every legacy paragraph must exist in GCC → FAIL if any missing</li>
    <li><b>Headings</b> — every legacy heading must exist in GCC with exact text → FAIL if any missing</li>
    <li><b>Left nav MATCHING</b> — both legacy and GCC have nav and all items match</li>
    <li><b>Left nav MISMATCH</b> — both have nav but one or more items in legacy are missing from GCC</li>
    <li><b>Left nav MISSING</b> — legacy had a nav but GCC does not have one</li>
    <li><b>Left nav PRESENT</b> — only GCC has a nav, legacy did not have one (expected in modern SharePoint)</li>
    <li><b>Left nav N/A</b> — neither legacy nor GCC has a nav</li>
    <li><b>Note:</b> left nav result never affects the overall PASS/FAIL status — it is informational only</li>
    <li><b>Links FAILED</b> — one or more real links missing from GCC → FAIL</li>
    <li><b>Links PASSED (grey)</b> — links matched relatively (URL format or encoding differences) → not a failure</li>
    <li><b>Links PASSED (green)</b> — all links match exactly</li>
    <li><b>Links ignored</b> — file:// paths, LAN/UNC (\server), intranet (http://crikit/) — expected to be missing in GCC, not counted as failures</li>
  </ul>
</div>

<div class="controls">
  <input type="text" id="searchBox" placeholder="Search URL or page name..." oninput="filt()">
  <select id="sf_status" onchange="filt()">
    <option value="">All statuses</option>
    <option value="FAIL">FAIL</option>
    <option value="WARN">WARN</option>
    <option value="PASS">PASS</option>
    <option value="ERR">ERR</option>
  </select>
  <select id="sf_body" onchange="filt()">
    <option value="">Body text</option>
    <option value="MATCH">MATCH</option>
    <option value="DIFF">DIFF</option>
  </select>
  <select id="sf_nav" onchange="filt()">
    <option value="">Left nav</option>
    <option value="MATCHING">MATCHING</option>
    <option value="MISMATCH">MISMATCH</option>
    <option value="PRESENT">PRESENT</option>
    <option value="MISSING">MISSING</option>
    <option value="N/A">N/A</option>
  </select>
  <select id="sf_links" onchange="filt()">
    <option value="">Links</option>
    <option value="FAILED">FAILED</option>
    <option value="PASSED">PASSED</option>
    <option value="PASSED">PASSED</option>
    <option value="N/A">N/A</option>
  </select>
  <button onclick="expandAll()">Expand all</button>
  <button onclick="collapseAll()">Collapse all</button>
</div>

<div class="wrap">
<table>
  <thead><tr>
    <th>Legacy URL</th>
    <th>GCC URL</th>
    <th>Status</th>
    <th style="text-align:center">Body text</th>
    <th style="text-align:center">Content (W/P/H)</th>
    <th style="text-align:center">Left nav</th>
    <th style="text-align:center">Links</th>
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
  const s       = document.getElementById('searchBox').value.toLowerCase();
  const sfSt    = document.getElementById('sf_status').value;
  const sfBody  = document.getElementById('sf_body').value;
  const sfNav   = document.getElementById('sf_nav').value;
  const sfLinks = document.getElementById('sf_links').value;
  document.querySelectorAll('.mr').forEach(row => {{
    const show = (!s       || (row.dataset.pid||'').includes(s) || row.cells[0].textContent.toLowerCase().includes(s) || row.cells[1].textContent.toLowerCase().includes(s)) &&
                 (!sfSt    || (row.dataset.status||'').includes(sfSt))   &&
                 (!sfBody  || (row.dataset.body||'')  === sfBody)        &&
                 (!sfNav   || (row.dataset.nav||'')   === sfNav)         &&
                 (!sfLinks || (row.dataset.links||'') === sfLinks);
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
