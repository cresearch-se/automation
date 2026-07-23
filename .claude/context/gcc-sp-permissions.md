# GCC SharePoint Permissions Validation

## Background
Migration from classic SharePoint (CornerZone / cresearch1.sharepoint.com) to modern SharePoint
(GCC / cresearch3.sharepoint.com). Goal: validate that permissions were correctly transferred.

## Objective
Compare site permissions between source (CornerZone) and target (GCC).
- **MISSING IN GCC** — in CornerZone but not in GCC → real issue, must be fixed
- **EXTRA IN GCC** — in GCC but not in CornerZone → migration additions, not a problem, kept separate

## Data Files
- `tests/GCCSPPermissions/fixtures/Cornerzone_SP_PermissionMatrix_20260708_170344.csv` — 8139 rows
- `tests/GCCSPPermissions/fixtures/GG_SP_PermissionMatrix.xlsx` — 4855 rows

## Script
`tests/GCCSPPermissions/compare_permissions.py` — run with:
```
py tests/GCCSPPermissions/compare_permissions.py
```

## Output Files
- `tests/GCCSPPermissions/output/comparison_report.csv` — every CZ row with Status (MATCHED / MISSING IN GCC),
  showing CZ_WebUrl, GCC_WebUrl side by side, PrincipalName, PermissionLevel
- `tests/GCCSPPermissions/output/extra_in_gcc.csv` — GCC-only permissions (informational)

## Match Key
`normalized WebUrl + PrincipalName + PermissionLevel`
- Dropped SiteCollectionTitle (unreliable across environments — confirmed by Shridhar)
- Dropped ObjectUrl (sub-path structure also changes between environments)

## URL Normalization
Classic sub-sites become flat `/sites/` collections in Modern. Standard rule:
`/corp/hr/advisors` → `/sites/corp_hr_advisors` (join segments with `_`, prepend `/sites/`)

Known exceptions handled in `CLASSIC_TO_MODERN` lookup table:
- `corp/it` → `corpIT` (special token, no underscore)
- `corp/it/Projects/X` → `corpIT_X` (intermediate segments dropped)
- `corp/facilities/X` → `corp_X` ("facilities" segment dropped)
- `CRIKIT/InformationResources/DataSources/X` → `CRIKIT_X` (deep nesting dropped)
- `/corp/hr/ben` → `corp_hr_benefits` (renamed)
- `/consulting/util/WhereInLAN` → `consulting_WhereInLan` (renamed)

## System Permission Filter
`PermissionCategory = "System"` rows are filtered out from BOTH files before comparison.
Reason: SharePoint auto-generates "Limited Access / System" for individual users when they have
permissions on sub-items. Classic exposes these explicitly; Modern handles them internally.
They are not admin-assigned and should not be compared. This removed a large amount of false
MISSING IN GCC noise (Individual User + Limited Access + System rows).

## Key Decisions
- Extra GCC permissions are NOT printed to screen — only written to extra_in_gcc.csv
- Screen output shows: diagnostic URL normalization, summary counts, MISSING IN GCC details only
- Shridhar (team lead) confirmed: no universal join key exists; URL-based mapping is the approach

## Current Results (after all filters)
- Total CZ rows after filtering: 1976
- MATCHED: 1427
- MISSING IN GCC: 549
- EXTRA IN GCC: 142

## Filters Applied
1. `PermissionCategory = System` — removed 5808 CZ / 3321 GCC rows
2. Nintex/FormsApp URLs (different subdomain `cresearch1-7e6a...sharepoint.com`) — removed 276 CZ rows
3. Known test/decommissioned sites — removed 79 CZ rows

## Waiting On (from Shridhar)
Correct GCC equivalents for these Classic paths (fuzzy match had too many candidates):
- `/` (root site)
- `/corp`
- `/consulting`
- `/corp/hr`
- `/corp/ir`
- `/corp/ARC`
- `/corp/pawg`

Also asked: any other site collections intentionally not migrated?

## Status
- [x] Script written and pushed to gcc-sp-permission branch
- [x] URL normalization with 3-tier resolution (lookup table → standard rule → fuzzy)
- [x] System permission filter added
- [x] Nintex/workflow filter added
- [x] Test site exclusion list added
- [x] Script runs successfully, output CSVs generated
- [ ] Awaiting Shridhar's reply on URL mappings and additional excluded sites
- [ ] Once mappings confirmed, add to CLASSIC_TO_MODERN and rerun for final count
