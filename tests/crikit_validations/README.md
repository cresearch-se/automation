# CRIKIT — SharePoint to GCC Migration Content Validator

Automated content comparison tool that validates migrated SharePoint pages between the legacy environment (cresearch1) and the new GCC environment (cresearch3).

---

## What it does

Visits every page in both environments using a real Edge browser, extracts the page content, and compares them side by side. Generates a shareable HTML report showing which pages passed, warned, or failed.

### The 5 validation checks run on every page

| Check | Rule | Result |
|---|---|---|
| Body text | Must match exactly (whitespace ignored) | FAIL if different |
| Word count | GCC must not have >20% fewer words than legacy | FAIL if >20% drop, WARN if 5–20% |
| Paragraphs | Every paragraph in legacy must exist in GCC | FAIL if any missing |
| Headings | Every heading in legacy must exist in GCC (exact text) | FAIL if any missing |
| GCC left nav | Left navigation must be present on every GCC page | FAIL if missing |
| Link count | Informational — flags significant drops | WARN if >20% drop |

---

## Project structure

```
crikit_validations/
├── step1_extract.py      # Extracts page content from legacy and GCC
├── step2_compare.py      # Compares both outputs and generates HTML report
├── urls.csv              # List of all page pairs to compare
├── msedgedriver.exe      # Edge WebDriver (must match your Edge version)
├── README.md             # This file
│
│   (generated when you run the scripts)
├── legacy_content.json   # Extracted legacy page content
├── gcc_content.json      # Extracted GCC page content
└── diff_report.html      # Final comparison report (open in any browser)
```

---

## Prerequisites

**One-time setup — run these once ever:**

```
py -m pip install selenium beautifulsoup4 lxml
```

**EdgeDriver** — must match your installed Edge version exactly.
- Check your Edge version: `(Get-Item "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe").VersionInfo.ProductVersion`
- Download matching driver: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
- Place `msedgedriver.exe` in this project folder

---

## How to run

### Step 1a — Extract legacy content

```
py step1_extract.py --env legacy --urls urls.csv --output legacy_content.json
```

- Edge opens and navigates to the first legacy URL
- Log in with your **legacy SSO account**, complete any MFA
- Come back to PowerShell and press **Enter**
- Script visits all pages automatically — do not touch the browser
- Takes ~30–60 min for 950 pages

### Step 1b — Extract GCC content

```
py step1_extract.py --env gcc --urls urls.csv --output gcc_content.json
```

- Edge opens again
- Log in with your **GCC account and password**
- Press **Enter**, let it run

### Step 2 — Compare and generate report

```
py step2_compare.py --legacy legacy_content.json --gcc gcc_content.json --output diff_report.html
```

- Takes about 10 seconds
- Open the report in any browser:

```
start diff_report.html
```

---

## Understanding the report

| Status | Meaning |
|---|---|
| **PASS** | All 5 checks passed — content matches |
| **WARN** | Minor differences — worth reviewing but not blocking |
| **FAIL** | Content mismatch — needs to be fixed before go-live |
| **ERR-L** | Legacy page failed to load (timeout or access error) |
| **ERR-G** | GCC page failed to load (timeout or access error) |

- Click any row to expand and see the exact issues, missing paragraphs, missing headings, and a line-by-line text diff
- Use the filter box to show only FAIL or WARN pages
- The report is self-contained — share the `.html` file directly with your team, no server needed

---

## urls.csv format

Three columns, one row per page pair:

```
page_id,legacy_url,gcc_url
home,https://cresearch1.sharepoint.com/CRIKIT/Pages/home.aspx,https://cresearch3.sharepoint.com/sites/CRIKIT/SitePages/home.aspx
about,https://cresearch1.sharepoint.com/CRIKIT/Pages/about.aspx,https://cresearch3.sharepoint.com/sites/CRIKIT/SitePages/about.aspx
```

- `page_id` — unique name for each page, used in the report
- `legacy_url` — full URL of the page on the legacy SharePoint
- `gcc_url` — full URL of the corresponding page on GCC

---

## Troubleshooting

**Edge fails to start:**
- Confirm `msedgedriver.exe` is in this folder: `dir msedgedriver.exe`
- Confirm Edge is installed: `dir "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"`
- Make sure driver version matches Edge version exactly

**Pages timing out:**
- Check your VPN / network connection
- Re-run the script — it overwrites the output file cleanly

**SSO session expires mid-run:**
- Re-run step1_extract.py for that environment from the beginning

**DeprecationWarning about utcnow():**
- Harmless warning from Python 3.13 — does not affect results

---

## Notes

- You log in **once per run** — the browser session stays alive for all pages
- Legacy and GCC extractions are independent — run them in any order
- The JSON files are plain text — you can open them to inspect any page's extracted content
- Test with a small batch of URLs first before running all 950
