"""
SP Permission Matrix Comparison
Source: CornerZone (cresearch1.sharepoint.com) — classic environment
Target: GCC (cresearch3.sharepoint.com) — modern environment

URL resolution order for each Classic site:
  1. CLASSIC_TO_MODERN lookup table (explicit overrides for known exceptions)
  2. Standard rule: /corp/hr/advisors → /sites/corp_hr_advisors
  3. Fuzzy fallback: token search across GCC URLs (auto, no manual table needed)

Output files:
  comparison_report.csv — every CZ row: Status (MATCHED/MISSING IN GCC),
                          CZ_WebUrl, GCC_WebUrl, PrincipalName, PermissionLevel
  extra_in_gcc.csv      — permissions in GCC not in CornerZone (informational)
"""

import csv
import os
import openpyxl
from collections import defaultdict

CORNERZONE_CSV = "tests/GCCSPPermissions/fixtures/Cornerzone_SP_PermissionMatrix_20260708_170344.csv"
GCC_XLSX       = "tests/GCCSPPermissions/fixtures/GG_SP_PermissionMatrix.xlsx"
OUTPUT_DIR     = "tests/GCCSPPermissions/output"

# Explicit overrides for the handful of sites that need special handling.
# Only needed when standard rule AND fuzzy fallback would both give wrong results.
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

# URL fragments that identify Nintex/workflow app sites — never migrated to GCC
NINTEX_URL_KEYWORDS = [
    "nintexworkflow",
    "formsapp",
    "nintex",
]

# Classic site paths that were intentionally NOT migrated (test sites, decommissioned)
EXCLUDED_CLASSIC_PATHS = {
    "/PMSTest",
    "/SPTest",
    "/sites/BI-Test",
    "/sites/TestAR",
    "/sites/pwa",
    "/sites/apps",
    "/corp/pawg/TestEP",
    "/AzureAISearch",
}


def _is_nintex_url(url):
    url_lower = (url or "").lower()
    return any(k in url_lower for k in NINTEX_URL_KEYWORDS)


def _is_excluded_path(url):
    path = strip_domain(url)
    return path in EXCLUDED_CLASSIC_PATHS

# Built at runtime in build_url_mapping() — Classic path → resolved GCC path
_URL_MAP = {}


def strip_domain(url):
    url = (url or "").strip()
    for d in (_CZ_DOMAIN, _GCC_DOMAIN):
        if url.startswith(d):
            return url[len(d):].rstrip("/") or "/"
    return url.rstrip("/") or "/"


def _standard_normalize(path):
    """Apply the standard slash→underscore rule (no lookup table, no fuzzy)."""
    if path.startswith("/sites/"):
        return path
    segments = [s for s in path.split("/") if s]
    return ("/sites/" + "_".join(segments)) if segments else "/"


def build_url_mapping(cz_rows, gcc_rows):
    """
    Pre-compute Classic path → GCC path for every unique Classic WebUrl.
    Resolution order: lookup table → standard rule → fuzzy token fallback.
    Populates the module-level _URL_MAP used by make_key_classic().
    """
    global _URL_MAP
    gcc_sites = sorted(set(strip_domain((r.get("WebUrl") or "").strip()) for r in gcc_rows))
    gcc_set   = set(gcc_sites)

    fuzzy_used    = []
    unresolved    = []

    classic_paths = sorted(set(strip_domain((r.get("WebUrl") or "").strip()) for r in cz_rows))

    for path in classic_paths:
        # 1. Explicit lookup table
        if path in CLASSIC_TO_MODERN:
            _URL_MAP[path] = CLASSIC_TO_MODERN[path]
            continue

        # 2. Standard rule
        normalized = _standard_normalize(path)
        if normalized in gcc_set:
            _URL_MAP[path] = normalized
            continue

        # 3. Fuzzy token fallback — search GCC URLs for ones containing all tokens
        tokens = normalized.replace("/sites/", "").lower().split("_")
        candidates = [g for g in gcc_sites if all(t in g.lower() for t in tokens)]
        if not candidates:
            # Relax: match on just the last (most specific) token
            candidates = [g for g in gcc_sites if tokens[-1] in g.lower()]

        if len(candidates) == 1:
            _URL_MAP[path] = candidates[0]
            fuzzy_used.append((path, candidates[0]))
        elif len(candidates) > 1:
            # Multiple candidates — pick shortest (closest structural match)
            best = min(candidates, key=len)
            _URL_MAP[path] = best
            fuzzy_used.append((path, best + f"  [picked from {len(candidates)} candidates]"))
        else:
            _URL_MAP[path] = normalized  # leave as-is, will show as MISSING IN GCC
            unresolved.append(path)

    print(f"\nURL mapping built: {len(_URL_MAP)} Classic sites")
    print(f"  Resolved via lookup table  : {sum(1 for p in classic_paths if p in CLASSIC_TO_MODERN)}")
    print(f"  Resolved via standard rule : {len(classic_paths) - sum(1 for p in classic_paths if p in CLASSIC_TO_MODERN) - len(fuzzy_used) - len(unresolved)}")
    print(f"  Resolved via fuzzy match   : {len(fuzzy_used)}")
    print(f"  Unresolved (no GCC match)  : {len(unresolved)}")

    if fuzzy_used:
        print("\n  Fuzzy-matched sites (review these):")
        for orig, resolved in fuzzy_used:
            print(f"    {orig:50s} -> {resolved}")

    if unresolved:
        print("\n  Unresolved sites (will appear as MISSING IN GCC):")
        for u in unresolved:
            print(f"    {u}")


def resolve_classic_url(url):
    """Return the GCC path for a Classic WebUrl using the pre-built mapping."""
    path = strip_domain(url)
    return _URL_MAP.get(path, _standard_normalize(path))


def make_key_classic(row):
    return (
        resolve_classic_url((row.get("WebUrl") or "").strip()),
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
    gcc_key_to_row = {}
    for row in gcc_rows:
        k = make_key_modern(row)
        gcc_key_to_row[k] = row

    cz_keys = set(make_key_classic(r) for r in cz_rows)

    comparison = []
    for row in cz_rows:
        k = make_key_classic(row)
        cz_url    = row.get("WebUrl", "")
        gcc_match = gcc_key_to_row.get(k)

        comparison.append({
            "Status"             : "MATCHED" if gcc_match else "MISSING IN GCC",
            "CZ_WebUrl"          : cz_url,
            "GCC_WebUrl"         : gcc_match.get("WebUrl", "") if gcc_match else "",
            "CZ_NormalizedUrl"   : resolve_classic_url(cz_url),
            "PrincipalName"      : k[1],
            "PermissionLevel"    : k[2],
            "WebTitle_CZ"        : row.get("WebTitle", ""),
            "WebTitle_GCC"       : gcc_match.get("WebTitle", "") if gcc_match else "",
            "ObjectType"         : row.get("ObjectType", ""),
            "PrincipalType"      : row.get("PrincipalType", ""),
            "PermissionCategory" : row.get("PermissionCategory", ""),
        })

    extra_in_gcc = []
    for row in gcc_rows:
        k = make_key_modern(row)
        if k not in cz_keys:
            extra_in_gcc.append({
                "GCC_WebUrl"         : row.get("WebUrl", ""),
                "PrincipalName"      : k[1],
                "PermissionLevel"    : k[2],
                "WebTitle"           : row.get("WebTitle", ""),
                "ObjectType"         : row.get("ObjectType", ""),
                "PrincipalType"      : row.get("PrincipalType", ""),
                "PermissionCategory" : row.get("PermissionCategory", ""),
            })

    return comparison, extra_in_gcc


def print_summary(comparison, extra):
    matched = sum(1 for r in comparison if r["Status"] == "MATCHED")
    missing = sum(1 for r in comparison if r["Status"] == "MISSING IN GCC")
    print("\n" + "=" * 70)
    print("PERMISSION COMPARISON SUMMARY")
    print("=" * 70)
    print(f"  Total CornerZone rows          : {len(comparison)}")
    print(f"  MATCHED (found in GCC)         : {matched}")
    print(f"  MISSING IN GCC (not found)     : {missing}")
    print(f"  EXTRA IN GCC   (informational) : {len(extra)}")
    print("=" * 70)


def print_missing(comparison):
    missing = [r for r in comparison if r["Status"] == "MISSING IN GCC"]
    if not missing:
        print("\nNo missing permissions found.")
        return

    print(f"\n--- MISSING IN GCC ({len(missing)} total) ---")
    by_site = defaultdict(list)
    for r in missing:
        by_site[r["CZ_NormalizedUrl"]].append(r)

    print(f"  Affected sites: {len(by_site)}")
    for site, rows in sorted(by_site.items()):
        print(f"\n  Site: {site or '(blank)'} — {len(rows)} missing permissions")
        for r in rows[:10]:
            print(f"    [{r['ObjectType']}] {r['WebTitle_CZ']}")
            print(f"      Principal : {r['PrincipalName']} ({r['PrincipalType']})")
            print(f"      Permission: {r['PermissionLevel']} / {r['PermissionCategory']}")
        if len(rows) > 10:
            print(f"    ... and {len(rows) - 10} more (see comparison_report.csv)")


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
    print(f"{label} saved to: {path}")


if __name__ == "__main__":
    cz_rows  = load_cornerzone()
    gcc_rows = load_gcc()

    # Filter out system-generated permissions (e.g. Limited Access auto-granted by SharePoint).
    # These are not explicitly assigned by admins and don't appear consistently in Modern SP.
    cz_before  = len(cz_rows)
    gcc_before = len(gcc_rows)
    cz_rows  = [r for r in cz_rows  if (r.get("PermissionCategory") or "").strip() != "System"]
    gcc_rows = [r for r in gcc_rows if (r.get("PermissionCategory") or "").strip() != "System"]
    print(f"System rows filtered out       -- CZ: {cz_before - len(cz_rows)}, GCC: {gcc_before - len(gcc_rows)}")

    # Filter out Nintex/workflow app sites (different subdomain, never migrated to GCC)
    n = len(cz_rows)
    cz_rows = [r for r in cz_rows if not _is_nintex_url(r.get("WebUrl", ""))]
    print(f"Nintex/workflow rows filtered  -- CZ: {n - len(cz_rows)}")

    # Filter out intentionally excluded (test/decommissioned) classic sites
    n = len(cz_rows)
    cz_rows = [r for r in cz_rows if not _is_excluded_path(r.get("WebUrl", ""))]
    print(f"Excluded test sites filtered   -- CZ: {n - len(cz_rows)}")

    # Build URL mapping (lookup table → standard rule → fuzzy fallback)
    build_url_mapping(cz_rows, gcc_rows)

    comparison, extra = compare(cz_rows, gcc_rows)

    print_summary(comparison, extra)
    print_missing(comparison)

    print()
    write_csv(comparison, "comparison_report.csv", "Comparison report")
    write_csv(extra,      "extra_in_gcc.csv",      "Extra in GCC     ")
