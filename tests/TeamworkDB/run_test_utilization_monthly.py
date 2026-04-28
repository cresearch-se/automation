#!/usr/bin/env python3
"""
run_utilization_monthly_tests.py

Runs all utilization QA test groups, parses the output files, and generates
Excel and HTML reports saved to tests/TeamworkDB/output/.

Usage (run from Code/ directory):
    py tests/TeamworkDB/run_utilization_monthly_tests.py
"""

import subprocess
import sys
import os
import re
from pathlib import Path
from datetime import date

# ==========================================
# CONFIGURATION
# ==========================================

TEST_FILE  = "tests/TeamworkDB/test_utilization_monthly.py"
OUTPUT_DIR = Path("tests/TeamworkDB/output")

TEST_GROUPS = [
    {
        "name": "Format Validations",
        "filter": (
            "test_format_validations_by_location or "
            "test_totals_validation_by_location or "
            "test_format_validations_by_employee"
        ),
        "output_file": OUTPUT_DIR / "format-validations.txt"
    },
    {
        "name": "DB Location",
        "filter": (
            "test_db_comparison_us_monthly or "
            "test_db_comparison_europe_monthly or "
            "test_db_comparison_us_ytd or "
            "test_db_comparison_europe_ytd"
        ),
        "output_file": OUTPUT_DIR / "db-location.txt"
    },
    {
        "name": "DB Employee Monthly",
        "filter": "test_db_comparison_employee_monthly",
        "output_file": OUTPUT_DIR / "db-employee-monthly.txt"
    },
    {
        "name": "DB Employee YTD",
        "filter": "test_db_comparison_employee_ytd",
        "output_file": OUTPUT_DIR / "db-employee-ytd.txt"
    }
]


# ==========================================
# STEP 1 — RUN TESTS
# ==========================================

def run_tests():
    """Run all test groups and write output to individual .txt files."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for group in TEST_GROUPS:
        print(f"  Running: {group['name']}...")
        cmd = [
            sys.executable, "-m", "pytest", TEST_FILE,
            "-k", group["filter"],
            "-v", "--no-cov", "-s", "--tb=short"
        ]
        with open(group["output_file"], "w", encoding="utf-8") as f:
            result = subprocess.run(cmd, stdout=f, stderr=subprocess.STDOUT, text=True)
        status = "PASS" if result.returncode == 0 else "FAILURES FOUND"
        print(f"    → {status}")


# ==========================================
# STEP 2 — PARSE OUTPUT FILES
# ==========================================

def parse_output_file(filepath):
    """
    Parse a pytest output .txt file and return structured results:
    {
        "passed":       [test_name, ...],
        "failed":       [test_name, ...],
        "mismatches":   [{"sheet": str, "columns": [...], "rows": [{col: val}]}, ...],
        "missing":      [line, ...],
        "other_errors": [line, ...]
    }
    """
    content = Path(filepath).read_text(encoding="utf-8", errors="replace")
    lines   = content.splitlines()

    passed       = []
    failed       = []
    mismatches   = []
    missing_rows = []
    other_errors = []

    # Extract pass/fail per test
    for line in lines:
        m = re.search(r'::(test_\S+)\s+(PASSED|FAILED)', line)
        if m:
            (passed if m.group(2) == "PASSED" else failed).append(m.group(1))

    # Extract mismatch tables, missing rows, format errors
    i = 0
    while i < len(lines):
        line = lines[i].strip()

        # VALUE MISMATCH table block
        if "UTILIZATION QA" in line and "VALUE MISMATCHES" in line:
            sheet = ""
            if i + 1 < len(lines):
                m = re.search(r'Sheet:\s*(.+?)\s*\|', lines[i + 1])
                if m:
                    sheet = m.group(1).strip()

            # Find header row (has 'Column', 'Excel', 'DB')
            header_idx = None
            for j in range(i + 1, min(i + 10, len(lines))):
                if "Column" in lines[j] and "Excel" in lines[j] and "DB" in lines[j]:
                    header_idx = j
                    break

            if header_idx is None:
                i += 1
                continue

            col_names  = [c.strip() for c in lines[header_idx].split("|")]
            data_start = header_idx + 2  # skip separator line
            table_rows = []

            for j in range(data_start, len(lines)):
                data_line = lines[j].strip()
                if not data_line or "|" not in data_line:
                    break
                values = [v.strip() for v in data_line.split("|")]
                if len(values) == len(col_names):
                    table_rows.append(dict(zip(col_names, values)))

            if table_rows:
                mismatches.append({
                    "sheet":   sheet,
                    "columns": col_names,
                    "rows":    table_rows
                })

            i = data_start + len(table_rows)
            continue

        # Missing rows
        if line.startswith("[MISSING IN DB]") or line.startswith("[MISSING IN XLS]"):
            missing_rows.append(line)

        # Format validation errors
        if any(line.startswith(tag) for tag in [
            "[MISSING OFFICE]", "[MISSING TITLE]", "[BLANK VALUE]",
            "[DUPLICATE", "[WRONG ORDER]", "[ZERO VALUE]", "[MISSING SUBTOTAL",
            "[MISSING GRAND TOTAL"
        ]):
            other_errors.append(line)

        i += 1

    return {
        "passed":       passed,
        "failed":       failed,
        "mismatches":   mismatches,
        "missing":      missing_rows,
        "other_errors": other_errors
    }


def parse_all_groups():
    """Parse all output files and attach group name to each result."""
    results = []
    for group in TEST_GROUPS:
        if not Path(group["output_file"]).exists():
            results.append({"name": group["name"], "passed": [], "failed": [],
                             "mismatches": [], "missing": [], "other_errors": [],
                             "error": "Output file not found"})
            continue
        parsed       = parse_output_file(group["output_file"])
        parsed["name"] = group["name"]
        results.append(parsed)
    return results


def get_report_name():
    """Read FIXTURE_FILE from the test file and derive the report name."""
    try:
        with open(TEST_FILE, encoding="utf-8") as f:
            for line in f:
                m = re.search(r'FIXTURE_FILE\s*=\s*["\'].*?([^/\\]+\.xlsx)["\']', line)
                if m:
                    month = m.group(1).replace("Utilization_", "").replace(".xlsx", "")
                    return f"utilization_failures_report_{month}"
    except Exception:
        pass
    return f"utilization_failures_report_{date.today().strftime('%Y%m')}"


# ==========================================
# STEP 3A — EXCEL REPORT
# ==========================================

def generate_excel_report(results, report_name):
    """Generate an Excel report with a Summary sheet and one sheet per test group."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Font, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("  [SKIP] openpyxl not installed. Run: pip install openpyxl")
        return

    wb = Workbook()
    wb.remove(wb.active)

    # Styles
    green_fill  = PatternFill("solid", fgColor="C6EFCE")
    red_fill    = PatternFill("solid", fgColor="FFC7CE")
    header_fill = PatternFill("solid", fgColor="2F5496")
    header_font = Font(color="FFFFFF", bold=True)
    bold_font   = Font(bold=True)
    title_font  = Font(bold=True, size=13)
    thin        = Side(style="thin")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)

    def set_header_row(ws, values):
        ws.append(values)
        row = ws.max_row
        for col_idx in range(1, len(values) + 1):
            cell = ws.cell(row, col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border

    def auto_width(ws):
        for col in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=10)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 3, 55)

    # ---- Summary sheet ----
    ws_sum = wb.create_sheet("Summary")
    ws_sum.append(["UTILIZATION QA REPORT"])
    ws_sum["A1"].font = title_font
    ws_sum.append([f"Report: {report_name}   |   Run Date: {date.today().strftime('%Y-%m-%d')}"])
    ws_sum.append([])
    set_header_row(ws_sum, ["Group", "Passed", "Failed", "Mismatches", "Missing", "Status"])

    for r in results:
        passed   = len(r.get("passed", []))
        failed   = len(r.get("failed", []))
        mismatch = sum(len(m["rows"]) for m in r.get("mismatches", []))
        missing  = len(r.get("missing", []))
        status   = "PASS" if failed == 0 else "FAIL"
        ws_sum.append([r["name"], passed, failed, mismatch, missing, status])
        row_num = ws_sum.max_row
        fill = green_fill if status == "PASS" else red_fill
        for col in range(1, 7):
            ws_sum.cell(row_num, col).fill   = fill
            ws_sum.cell(row_num, col).border = border

    auto_width(ws_sum)

    # ---- One sheet per group ----
    for r in results:
        sheet_name = r["name"][:31]
        ws = wb.create_sheet(sheet_name)
        passed = len(r.get("passed", []))
        failed = len(r.get("failed", []))

        ws.append([r["name"]])
        ws["A1"].font = title_font
        ws.append([f"Run: {date.today().strftime('%Y-%m-%d')}   |   Passed: {passed}   |   Failed: {failed}"])
        ws.append([])

        # Value mismatch tables
        for block in r.get("mismatches", []):
            ws.append([f"Sheet: {block['sheet']}"])
            ws.cell(ws.max_row, 1).font = bold_font
            set_header_row(ws, block["columns"])
            for row_data in block["rows"]:
                ws.append([row_data.get(c, "") for c in block["columns"]])
                row_num = ws.max_row
                for col_idx in range(1, len(block["columns"]) + 1):
                    ws.cell(row_num, col_idx).fill   = red_fill
                    ws.cell(row_num, col_idx).border = border
            ws.append([])

        # Missing rows
        if r.get("missing"):
            ws.append(["MISSING ROWS"])
            ws.cell(ws.max_row, 1).font = bold_font
            for line in r["missing"]:
                ws.append([line])
            ws.append([])

        # Format/other errors
        if r.get("other_errors"):
            ws.append(["FORMAT ERRORS"])
            ws.cell(ws.max_row, 1).font = bold_font
            for line in r["other_errors"]:
                ws.append([line])

        auto_width(ws)

    out_path = OUTPUT_DIR / f"{report_name}.xlsx"
    wb.save(out_path)
    print(f"  Excel report: {out_path}")


# ==========================================
# STEP 3B — HTML REPORT
# ==========================================

def generate_html_report(results, report_name):
    """Generate a self-contained, collapsible HTML report."""
    run_date      = date.today().strftime("%Y-%m-%d")
    total_passed  = sum(len(r.get("passed", [])) for r in results)
    total_failed  = sum(len(r.get("failed", [])) for r in results)
    overall       = "PASS" if total_failed == 0 else "FAIL"
    overall_color = "#27ae60" if overall == "PASS" else "#e74c3c"

    def badge(status):
        color = "#27ae60" if status == "PASS" else "#e74c3c"
        return (f'<span style="background:{color};color:white;padding:2px 10px;'
                f'border-radius:4px;font-weight:bold;font-size:0.85em">{status}</span>')

    def mismatch_table_html(block):
        cols    = block["columns"]
        headers = "".join(f"<th>{c}</th>" for c in cols)
        body    = "".join(
            "<tr>" + "".join(f"<td>{row.get(c, '')}</td>" for c in cols) + "</tr>"
            for row in block["rows"]
        )
        return f"""
            <p class="sheet-label">Sheet: <strong>{block['sheet']}</strong>
            &nbsp;({len(block['rows'])} mismatch(es))</p>
            <table><thead><tr>{headers}</tr></thead><tbody>{body}</tbody></table>"""

    group_sections = ""
    for r in results:
        passed   = len(r.get("passed", []))
        failed   = len(r.get("failed", []))
        status   = "PASS" if failed == 0 else "FAIL"
        bg_color = "#eafaf1" if status == "PASS" else "#fdedec"
        border_c = "#27ae60" if status == "PASS" else "#e74c3c"

        body_html = ""

        for block in r.get("mismatches", []):
            body_html += mismatch_table_html(block)

        if r.get("missing"):
            items = "".join(f"<li>{line}</li>" for line in r["missing"])
            body_html += f'<p class="sheet-label"><strong>Missing Rows</strong></p><ul>{items}</ul>'

        if r.get("other_errors"):
            items = "".join(f"<li>{line}</li>" for line in r["other_errors"])
            body_html += f'<p class="sheet-label"><strong>Format Errors</strong></p><ul>{items}</ul>'

        if not body_html:
            body_html = "<p class='all-pass'>&#10003; All checks passed.</p>"

        group_sections += f"""
        <details open>
          <summary style="background:{bg_color};border-left:5px solid {border_c};
                          padding:10px 16px;cursor:pointer;font-size:1.05em;font-weight:bold">
            {r['name']} &nbsp;&mdash;&nbsp; Passed: {passed} &nbsp;|&nbsp; Failed: {failed}
            &nbsp;&nbsp;{badge(status)}
          </summary>
          <div class="group-body" style="border:1px solid {border_c};border-top:none">
            {body_html}
          </div>
        </details>"""

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Utilization QA — {report_name}</title>
  <style>
    body        {{ font-family: Arial, sans-serif; margin: 30px; color: #333; background: #f9f9f9; }}
    h1          {{ color: #2c3e50; margin-bottom: 5px; }}
    .summary    {{ background: white; border: 1px solid #ddd; border-radius: 6px;
                   padding: 15px 20px; margin-bottom: 25px; font-size: 0.95em; }}
    .summary span {{ margin-right: 20px; }}
    details      {{ margin-bottom: 10px; border-radius: 4px; overflow: hidden; }}
    summary      {{ list-style: none; user-select: none; }}
    summary::-webkit-details-marker {{ display: none; }}
    .group-body  {{ padding: 15px 20px; background: white; }}
    table        {{ border-collapse: collapse; width: 100%; margin-bottom: 15px;
                   font-size: 0.88em; }}
    th           {{ background: #2f5496; color: white; padding: 7px 12px; text-align: left; }}
    td           {{ padding: 6px 12px; border: 1px solid #ddd; }}
    tr:nth-child(even) td {{ background: #fef9f9; }}
    tr:nth-child(odd)  td {{ background: #fff; }}
    .sheet-label {{ margin: 12px 0 4px; color: #555; }}
    .all-pass    {{ color: #27ae60; font-weight: bold; }}
    ul           {{ margin: 5px 0 15px 20px; font-size: 0.9em; }}
  </style>
</head>
<body>
  <h1>Utilization QA Report</h1>
  <div class="summary">
    <span><strong>Report:</strong> {report_name}</span>
    <span><strong>Run Date:</strong> {run_date}</span>
    <span><strong>Total Passed:</strong> {total_passed}</span>
    <span><strong>Total Failed:</strong> {total_failed}</span>
    <span><strong>Overall:</strong>
      <span style="color:{overall_color};font-weight:bold">{overall}</span>
    </span>
  </div>
  {group_sections}
</body>
</html>"""

    out_path = OUTPUT_DIR / f"{report_name}.html"
    out_path.write_text(html, encoding="utf-8")
    print(f"  HTML  report: {out_path}")


# ==========================================
# MAIN
# ==========================================

if __name__ == "__main__":
    print("=" * 55)
    print("  UTILIZATION QA TEST RUNNER")
    print("=" * 55)

    #print("\n[1/3] Running tests...")
    #run_tests()

    print("\n[2/3] Parsing output files...")
    results = parse_all_groups()

    report_name = get_report_name()
    print(f"\n[3/3] Generating reports ({report_name})...")
    generate_excel_report(results, report_name)
    generate_html_report(results, report_name)

    total_failed = sum(len(r.get("failed", [])) for r in results)
    print("\n" + "=" * 55)
    if total_failed == 0:
        print("  ALL TESTS PASSED")
    else:
        print(f"  {total_failed} TEST(S) FAILED — check reports in {OUTPUT_DIR}")
    print("=" * 55)
