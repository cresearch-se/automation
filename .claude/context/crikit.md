# CRIKIT — SharePoint to GCC Migration Content Validator

## Overview
Automated content comparison tool that validates migrated SharePoint pages between the legacy environment (cresearch1) and the new GCC environment (cresearch3). Uses a real Edge browser to extract page content and compares them side by side.

## Key Files
```
tests/crikit_validations/
├── step1_extract.py          — Browser-based extractor (Selenium + Edge)
├── step2_compare.py          — Comparison engine + HTML report generator
├── step2_compare_nonav.py    — Variant without nav checks
├── step3_check_links.py      — Link-level validation pass
├── urls.csv                  — All page pairs to compare (950 pages)
├── msedgedriver.exe           — Edge WebDriver (must match installed Edge version)
├── legacy_content.json       — Extracted legacy page content (generated)
├── gcc_content.json          — Extracted GCC page content (generated)
└── diff_report.html          — Final comparison report (generated)
```

## Validation Checks (per page)
| Check | Rule | Result |
|---|---|---|
| Body text | Must match exactly (whitespace ignored) | FAIL if different |
| Word count | GCC must not have >20% fewer words than legacy | FAIL if >20% drop, WARN if 5–20% |
| Paragraphs | Every paragraph in legacy must exist in GCC | FAIL if any missing |
| Headings | Every heading in legacy must exist in GCC (exact text) | FAIL if any missing |
| GCC left nav | Left nav must be present on every GCC page | FAIL if missing |
| Link count | Informational — flags significant drops | WARN if >20% drop |

## Report Statuses
| Status | Meaning |
|---|---|
| PASS | All checks passed |
| WARN | Minor differences — worth reviewing |
| FAIL | Content mismatch — must fix before go-live |
| ERR-L | Legacy page failed to load |
| ERR-G | GCC page failed to load |

## How to Run

### Prerequisites (one-time)
```powershell
py -m pip install selenium beautifulsoup4 lxml
```
EdgeDriver must match installed Edge version exactly — place `msedgedriver.exe` in the `crikit_validations/` folder.

### Step 1a — Extract legacy content
```powershell
cd tests/crikit_validations
py step1_extract.py --env legacy --urls urls.csv --output legacy_content.json
```
- Edge opens — log in with legacy SSO account, complete MFA, then press Enter
- Script visits all pages automatically (~30–60 min for 950 pages)

### Step 1b — Extract GCC content
```powershell
py step1_extract.py --env gcc --urls urls.csv --output gcc_content.json
```
- Edge opens — log in with GCC account, press Enter

### Step 2 — Compare and generate report
```powershell
py step2_compare.py --legacy legacy_content.json --gcc gcc_content.json --output diff_report.html
start diff_report.html
```

### Step 3 — Link check (optional additional pass)
```powershell
py step3_check_links.py
```

## urls.csv Format
```
page_id,legacy_url,gcc_url
home,https://cresearch1.sharepoint.com/CRIKIT/Pages/home.aspx,https://cresearch3.sharepoint.com/sites/CRIKIT/SitePages/home.aspx
```

## Notes
- Run `step1_extract.py` from within `tests/crikit_validations/` — the WebDriver path is relative
- Each environment's extraction is independent — run in any order
- The JSON files are plain text and can be inspected directly
- The generated HTML report is self-contained — share it directly with the team, no server needed
- SSO session stays alive for the full run; if it expires, re-run step 1 from the beginning for that environment
- `diff_report_nonav.html` and the `*_failed*` / `*_retest*` variants in the folder are past runs kept for reference
