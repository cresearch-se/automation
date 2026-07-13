"""
SP Permission Matrix Comparison
Source: CornerZone (cresearch1.sharepoint.com) — classic environment
Target: GCC (cresearch3.sharepoint.com) — modern environment

Match key: normalized WebUrl (Classic → Modern path transform) + PrincipalName + PermissionLevel

Rules:
  - MISSING IN GCC : permission exists in CornerZone but not in GCC  (failure)
  - EXTRA IN GCC   : permission exists in GCC but not in CornerZone   (informational)

URL transform (classic sub-sites become flat /sites/ collections in modern):
  Standard rule : /corp/hr/advisors  →  /sites/corp_hr_advisors
  Exceptions    : see CLASSIC_TO_MODERN lookup table below
"""

import csv
import os
import openpyxl
from collections import defaultdict

CORNERZONE_CSV = "tests/GCCSPPermissions/fixtures/Cornerzone_SP_PermissionMatrix_20260708_170344.csv"
GCC_XLSX       = "tests/GCCSPPermissions/fixtures/GG_SP_PermissionMatrix.xlsx"
OUTPUT_DIR     = "tests/GCCSPPermissions/output"

# Explicit overrides for URLs that don't follow the standard slash→underscore rule.
# Key: Classic path (no domain, no trailing slash)
# Value: Modern path (no domain, no trailing slash)
CLASSIC_TO_MODERN = {
    # corp/it collapses to the token "corpIT" (no underscore between corp and it)
    "/corp/it":                                              "/sites/corpIT",
    "/corp/it/Appdevprojects":                               "/sites/corpIT_Appdevprojects",
    "/corp/it/dsg":                                          "/sites/corpIT_dsg",
    "/corp/it/litigationsupport":                            "/sites/corpIT_litigationsupport",
    "/corp/it/private":                                      "/sites/corpIT_private",
    "/corp/it/private/budget":                               "/sites/corpIT_budget",
    "/corp/it/Projects/Desktop":                             "/sites/corpIT_Desktop",
    "/corp/it/Projects/Pipeline":                            "/sites/corpIT_Pipeline",
    "/corp/it/Projects/SV-Migration":                        "/sites/corpIT_migration",
    "/corp/mobiledevices/MobileDevices":                     "/sites/corpIT_mobiledevices_MD",
    # "facilities" middle segment was dropped during migration
    "/corp/facilities/CRBEOffice":                           "/sites/corp_CRBEOffice",
    "/corp/facilities/CRBOffice":                            "/sites/corp_CRBOffice",
    # CRIKIT — deep intermediate segments dropped
    "/CRIKIT/InformationResources/DataSources/DataSourcesA": "/sites/CRIKIT_DataSourcesA",
    # Sites renamed during migration
    "/corp/hr/ben":                                          "/sites/corp_hr_benefits",
    "/consulting/util/WhereInLAN":                           "/sites/consulting_WhereInLan",
}

_CZ_DOMAIN  = "https://cresearch1.sharepoint.com"
_GCC_DOMAIN = "https://cresearch3.sharepoint.com"


def strip_domain(url):
    url = (url or "").strip()
    for d in (_CZ_DOMAIN, _GCC_DOMAIN):
        if url.startswith(d):
            return url[len(d):].rstrip("/") or "/"
    return url.rstrip("/") or "/"


def normalize_classic_url(url):
    """Transform a Classic SharePoint WebUrl path to its expected Modern GCC path."""
    path = strip_domain(url)
    if path in CLASSIC_TO_MODERN:
        return CLASSIC_TO_MODERN[path]
    if path.startswith("/sites/"):
        return path  # Already modern-style, leave untouched
    segments = [s for s in path.split("/") if s]
    if not segments:
        return "/"
    return "/sites/" + "_".join(segments)


def make_key_classic(row):
    return (
        normalize_classic_url((row.get("WebUrl") or "").strip()),
        (row.get("PrincipalName") or "").strip(),
        (row.get("PermissionLevel") or "").strip(),
    )


def make_key_modern(row):
    return (
        strip_domain((row.get("WebUrl") or "").strip()),
        (row.get("PrincipalName") or "").strip(),
        (row.get("PermissionLevel") or "").strip(),
    )


def load_cornerzone():
    rows = []
    with open(CORNERZONE_CSV, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            rows.append(row)
    print(f"CornerZone rows loaded : {len(rows)}")
    return rows


def load_gcc():
    wb = openpyxl.load_workbook(GCC_XLSX, read_only=True, data_only=True)
    ws = wb.active
    rows = []
    headers = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            headers = [str(c).strip() if c else "" for c in row]
        else:
            rows.append(dict(zip(headers, [
                str(c).strip() if c is not None else "" for c in row
            ])))
    wb.close()
    print(f"GCC rows loaded        : {len(rows)}")
    return rows


def compare(cz_rows, gcc_rows):
    gcc_keys = set(make_key_modern(r) for r in gcc_rows)
    cz_keys  = set(make_key_classic(r) for r in cz_rows)

    missing_in_gcc = []
    matched = []
    for row in cz_rows:
        k = make_key_classic(row)
        entry = {
            "NormalizedWebUrl"   : k[0],
            "PrincipalName"      : k[1],
            "PermissionLevel"    : k[2],
            "OriginalWebUrl"     : row.get("WebUrl", ""),
            "WebTitle"           : row.get("WebTitle", ""),
            "ObjectType"         : row.get("ObjectType", ""),
            "PrincipalType"      : row.get("PrincipalType", ""),
            "PermissionCategory" : row.get("PermissionCategory", ""),
        }
        if k not in gcc_keys:
            missing_in_gcc.append(entry)
        else:
            matched.append(entry)

    extra_in_gcc = []
    for row in gcc_rows:
        k = make_key_modern(row)
        if k not in cz_keys:
            extra_in_gcc.append({
                "NormalizedWebUrl"   : k[0],
                "PrincipalName"      : k[1],
                "PermissionLevel"    : k[2],
                "OriginalWebUrl"     : row.get("WebUrl", ""),
                "WebTitle"           : row.get("WebTitle", ""),
                "ObjectType"         : row.get("ObjectType", ""),
                "PrincipalType"      : row.get("PrincipalType", ""),
                "PermissionCategory" : row.get("PermissionCategory", ""),
            })

    return missing_in_gcc, extra_in_gcc, matched


def print_summary(missing, extra):
    print("\n" + "=" * 70)
    print("PERMISSION COMPARISON SUMMARY")
    print("=" * 70)
    print(f"  MISSING IN GCC (failures)      : {len(missing)}")
    print(f"  EXTRA IN GCC   (informational) : {len(extra)}")
    print("=" * 70)


def print_missing(missing):
    print(f"\n--- MISSING IN GCC ({len(missing)} total) ---")
    by_site = defaultdict(list)
    for r in missing:
        by_site[r["NormalizedWebUrl"]].append(r)

    print(f"  Affected sites: {len(by_site)}")
    for site, rows in sorted(by_site.items()):
        print(f"\n  Site: {site or '(blank)'} — {len(rows)} missing permissions")
        for r in rows[:10]:
            print(f"    [{r['ObjectType']}] {r['WebTitle']}")
            print(f"      Principal : {r['PrincipalName']} ({r['PrincipalType']})")
            print(f"      Permission: {r['PermissionLevel']} / {r['PermissionCategory']}")
        if len(rows) > 10:
            print(f"    ... and {len(rows) - 10} more (see CSV)")


def print_extra(extra):
    print(f"\n--- EXTRA IN GCC ({len(extra)} total) [informational] ---")
    by_site = defaultdict(list)
    for r in extra:
        by_site[r["NormalizedWebUrl"]].append(r)

    print(f"  Affected sites: {len(by_site)}")
    for site, rows in sorted(by_site.items()):
        print(f"\n  Site: {site or '(blank)'} — {len(rows)} extra permissions")
        for r in rows[:5]:
            print(f"    [{r['ObjectType']}] {r['WebTitle']}")
            print(f"      Principal : {r['PrincipalName']} ({r['PrincipalType']})")
            print(f"      Permission: {r['PermissionLevel']} / {r['PermissionCategory']}")
        if len(rows) > 5:
            print(f"    ... and {len(rows) - 5} more (see CSV)")


def write_csv(rows, filename, label):
    if not rows:
        print(f"\n{label}: none found.")
        return
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    path = os.path.join(OUTPUT_DIR, filename)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)
    print(f"\n{label} saved to: {path}")


def diagnose_matching(cz_rows, gcc_rows):
    """Show how Classic URLs normalize and how many match Modern sites."""
    print("\n--- DIAGNOSTIC: URL normalization (Classic → Modern) ---")
    seen_cz = set()
    for r in cz_rows:
        orig = (r.get("WebUrl") or "").strip()
        if orig in seen_cz:
            continue
        seen_cz.add(orig)
        print(f"  {strip_domain(orig):50s}  →  {normalize_classic_url(orig)}")
        if len(seen_cz) >= 15:
            break

    print("\n--- DIAGNOSTIC: Modern WebUrls (sample) ---")
    seen_gcc = set()
    for r in gcc_rows:
        url = strip_domain((r.get("WebUrl") or "").strip())
        if url not in seen_gcc:
            seen_gcc.add(url)
            print(f"  {url}")
        if len(seen_gcc) >= 15:
            break

    gcc_sites = set(strip_domain((r.get("WebUrl") or "").strip()) for r in gcc_rows)
    cz_normalized = set(normalize_classic_url((r.get("WebUrl") or "").strip()) for r in cz_rows)
    matched   = cz_normalized & gcc_sites
    unmatched = cz_normalized - gcc_sites

    print(f"\nClassic unique WebUrls (after normalize) : {len(cz_normalized)}")
    print(f"Modern unique WebUrls                    : {len(gcc_sites)}")
    print(f"Matched sites                            : {len(matched)}")
    print(f"Unmatched Classic sites (no Modern pair) : {len(unmatched)}")
    if unmatched:
        print("  Unmatched (first 20 — add to CLASSIC_TO_MODERN if needed):")
        for u in sorted(unmatched)[:20]:
            print(f"    {u}")


if __name__ == "__main__":
    cz_rows  = load_cornerzone()
    gcc_rows = load_gcc()

    diagnose_matching(cz_rows, gcc_rows)

    missing, extra, matched = compare(cz_rows, gcc_rows)

    print_summary(missing, extra)
    print(f"  MATCHED (found in both)        : {len(matched)}")
    print_missing(missing)

    write_csv(missing, "missing_in_gcc.csv", "Missing in GCC")
    write_csv(extra,   "extra_in_gcc.csv",   "Extra in GCC")
    write_csv(matched, "matched.csv",         "Matched permissions")
