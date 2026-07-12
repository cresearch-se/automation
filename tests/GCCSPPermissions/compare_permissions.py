"""
SP Permission Matrix Comparison
Source: CornerZone (cresearch1.sharepoint.com) — classic environment
Target: GCC (cresearch3.sharepoint.com) — modern environment

Rules:
  - MISSING IN GCC : permission exists in CornerZone but not in GCC  (failure)
  - EXTRA IN GCC   : permission exists in GCC but not in CornerZone   (informational)

Match key: SiteCollectionTitle + relative ObjectUrl + PrincipalName + PermissionLevel
"""

import csv
import os
import openpyxl
from collections import defaultdict

CORNERZONE_CSV = "tests/GCCSPPermissions/fixtures/Cornerzone_SP_PermissionMatrix_20260708_170344.csv"
GCC_XLSX       = "tests/GCCSPPermissions/fixtures/GG_SP_PermissionMatrix.xlsx"
OUTPUT_DIR     = "tests/GCCSPPermissions/output"


def relative_url(url):
    """Strip domain so cresearch1 and cresearch3 URLs can be compared by path."""
    for prefix in [
        "https://cresearch1.sharepoint.com",
        "https://cresearch3.sharepoint.com",
    ]:
        if url.startswith(prefix):
            return url[len(prefix):].rstrip("/") or "/"
    return url.rstrip("/") or "/"


def make_key(row):
    return (
        (row.get("SiteCollectionTitle") or "").strip(),
        relative_url((row.get("ObjectUrl") or "").strip()),
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
            rows.append(dict(zip(headers, [str(c).strip() if c is not None else "" for c in row])))
    wb.close()
    print(f"GCC rows loaded        : {len(rows)}")
    return rows


def compare(cz_rows, gcc_rows):
    gcc_keys = set(make_key(r) for r in gcc_rows)
    cz_keys  = set(make_key(r) for r in cz_rows)

    missing_in_gcc = []
    for row in cz_rows:
        k = make_key(row)
        if k not in gcc_keys:
            missing_in_gcc.append({
                "SiteCollectionTitle": k[0],
                "ObjectUrl_relative" : k[1],
                "PrincipalName"      : k[2],
                "PermissionLevel"    : k[3],
                "OriginalObjectUrl"  : row.get("ObjectUrl", ""),
                "ObjectType"         : row.get("ObjectType", ""),
                "PrincipalType"      : row.get("PrincipalType", ""),
                "PermissionCategory" : row.get("PermissionCategory", ""),
            })

    extra_in_gcc = []
    for row in gcc_rows:
        k = make_key(row)
        if k not in cz_keys:
            extra_in_gcc.append({
                "SiteCollectionTitle": k[0],
                "ObjectUrl_relative" : k[1],
                "PrincipalName"      : k[2],
                "PermissionLevel"    : k[3],
                "OriginalObjectUrl"  : row.get("ObjectUrl", ""),
                "ObjectType"         : row.get("ObjectType", ""),
                "PrincipalType"      : row.get("PrincipalType", ""),
                "PermissionCategory" : row.get("PermissionCategory", ""),
            })

    return missing_in_gcc, extra_in_gcc


def print_summary(missing, extra):
    print("\n" + "=" * 70)
    print("PERMISSION COMPARISON SUMMARY")
    print("=" * 70)
    print(f"  MISSING IN GCC (failures)      : {len(missing)}")
    print(f"  EXTRA IN GCC   (informational) : {len(extra)}")
    print("=" * 70)


def print_missing(missing, limit=50):
    print(f"\n--- MISSING IN GCC (first {min(limit, len(missing))} of {len(missing)}) ---")
    by_site = defaultdict(list)
    for r in missing:
        by_site[r["SiteCollectionTitle"]].append(r)

    print(f"  Affected sites: {len(by_site)}")
    for site, rows in sorted(by_site.items()):
        print(f"\n  Site: {site or '(blank)'} — {len(rows)} missing permissions")
        for r in rows[:10]:
            print(f"    [{r['ObjectType']}] {r['ObjectUrl_relative'] or '/'}")
            print(f"      Principal : {r['PrincipalName']} ({r['PrincipalType']})")
            print(f"      Permission: {r['PermissionLevel']} / {r['PermissionCategory']}")
        if len(rows) > 10:
            print(f"    ... and {len(rows) - 10} more")


def print_extra(extra, limit=20):
    print(f"\n--- EXTRA IN GCC (first {min(limit, len(extra))} of {len(extra)}) [informational] ---")
    by_site = defaultdict(list)
    for r in extra:
        by_site[r["SiteCollectionTitle"]].append(r)

    print(f"  Affected sites: {len(by_site)}")
    for site, rows in sorted(by_site.items()):
        print(f"\n  Site: {site or '(blank)'} — {len(rows)} extra permissions")
        for r in rows[:5]:
            print(f"    [{r['ObjectType']}] {r['ObjectUrl_relative'] or '/'}")
            print(f"      Principal : {r['PrincipalName']} ({r['PrincipalType']})")
            print(f"      Permission: {r['PermissionLevel']} / {r['PermissionCategory']}")
        if len(rows) > 5:
            print(f"    ... and {len(rows) - 5} more")


def write_csv(rows, filename, label):
    if not rows:
        return
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    path = os.path.join(OUTPUT_DIR, filename)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)
    print(f"\n{label} saved to: {path}")


def diagnose_matching(cz_rows, gcc_rows):
    """Print sample keys from both sides to help debug zero-match situations."""
    print("\n--- DIAGNOSTIC: Sample match keys ---")
    print("CornerZone sample keys:")
    for r in cz_rows[:5]:
        print(f"  {make_key(r)}")
    print("GCC sample keys:")
    for r in gcc_rows[:5]:
        print(f"  {make_key(r)}")

    cz_sites  = sorted(set(r.get("SiteCollectionTitle","").strip() for r in cz_rows))
    gcc_sites = sorted(set(r.get("SiteCollectionTitle","").strip() for r in gcc_rows))
    common    = set(cz_sites) & set(gcc_sites)
    print(f"\nCornerZone unique sites : {len(cz_sites)}")
    print(f"GCC unique sites        : {len(gcc_sites)}")
    print(f"Common site titles      : {len(common)}")
    if common:
        print("  Matched:", list(common)[:10])
    else:
        print("  NO COMMON SITE TITLES — sites likely have different names between environments")
        print("  CornerZone sites (first 10):", cz_sites[:10])
        print("  GCC sites (first 10)        :", gcc_sites[:10])


if __name__ == "__main__":
    cz_rows = load_cornerzone()
    gcc_rows = load_gcc()

    diagnose_matching(cz_rows, gcc_rows)

    missing, extra = compare(cz_rows, gcc_rows)

    print_summary(missing, extra)
    print_missing(missing)
    print_extra(extra)

    write_csv(missing, "missing_in_gcc.csv", "Missing in GCC")
    write_csv(extra,   "extra_in_gcc.csv",   "Extra in GCC")
