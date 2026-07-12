# GCC SharePoint Permissions Validation

## Background
Migration from classic SharePoint (CornerZone) to modern SharePoint (GCC). Goal is to validate that sites migrated from the classic permission model have correct permissions in their corresponding modern sites.

## Objective
Compare site permissions between source (CornerZone) and target (GCC) environments.

## What to Check
- **Source → Target match:** All source (CornerZone) permissions must exist in target (GCC) — this is the primary check
- **Missing permissions:** Permissions present in source but absent in target
- **Additional permissions:** Permissions in GCC not in CornerZone — these can be safely ignored (not a failure)
- **Permission mismatches/exceptions:** Same site/user but different permission level

## Key Rule
Extra permissions in GCC are acceptable — only missing permissions (from source) are failures.

## Data
- `tests/GCCSPPermissions/fixtures/GCC_Security_Matrix.xlsx` — target (modern GCC sites)
- `tests/GCCSPPermissions/fixtures/CornerZone_Security_Matrix.xlsx` — source (classic CornerZone sites)

## Files
- `tests/GCCSPPermissions/` — test folder
- `tests/GCCSPPermissions/fixtures/` — input Excel files
- `tests/GCCSPPermissions/output/` — generated reports

## Context
- This is likely a one-time activity (migration validation)
- Team lead asked for a quick feasibility check (~1 hour) before committing to full automation
- If too complex, manual validation continues; automation would help for detailed library/folder validation starting early next week

## Status
- [ ] Excel files received and structure understood
- [ ] Comparison logic designed
- [ ] Tests written
- [ ] Report output generated
