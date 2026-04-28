import os
from pathlib import Path
from typing import Optional

import pytest

from cornerstone_automation.utils.excel_utils import read_excel, get_excel_headers

def _find_fixture(fixtures_dir: Path, name_fragment: str) -> Optional[Path]:
    """
    Find a file in fixtures_dir whose filename contains name_fragment and has an xls/xlsx extension.
    Returns Path or None.
    """
    for ext in ("*.xlsx", "*.xls"):
        for p in fixtures_dir.glob(ext):
            if name_fragment.lower() in p.name.lower():
                return p
    return None

def _normalize_name(val):
    if val is None:
        return ""
    if isinstance(val, str):
        return val.strip().lower()
    return str(val).strip().lower()

def test_compare_employee_rows_between_timeoff_sheets():
    """
    For each employee name in column A of frm_Rpt_Det_given, attempt to find a matching
    row in VHSE_Holiday_Report_202510 (matching on column A / employee name).
    If a match is found print the row from both sheets; otherwise print a missing message.
    This test is informational (prints matches / misses) and will not fail the run.
    """
    base = Path(__file__).parent
    fixtures_dir = base / "fixtures"

    # Locate expected files under tests/TimeOffSheets/fixtures
    frm_file = _find_fixture(fixtures_dir, "frm_rpt_det_given")
    vhse_file = _find_fixture(fixtures_dir, "vhse_holiday_report_202510")

    assert frm_file is not None, f"Could not find frm_Rpt_Det_given file in {fixtures_dir}"
    assert vhse_file is not None, f"Could not find VHSE_Holiday_Report_202510 file in {fixtures_dir}"

    # Read headers and rows
    frm_headers = get_excel_headers(str(frm_file))
    vhse_headers = get_excel_headers(str(vhse_file))

    # Determine key (first column header) for both sheets
    frm_key = frm_headers[0] if frm_headers else ""  # empty string if header cell empty
    vhse_key = vhse_headers[0] if vhse_headers else ""

    frm_rows = read_excel(str(frm_file))
    vhse_rows = read_excel(str(vhse_file))

    # If read_excel produced empty list but file has rows, we still handle gracefully
    if not frm_rows:
        print(f"No data rows found in {frm_file}.")
        return

    # Build lookup map for VHSE by normalized employee name -> list of rows
    vhse_lookup = {}
    for r in vhse_rows:
        name = _normalize_name(r.get(vhse_key) if isinstance(r, dict) else None)
        if not name:
            # Attempt to fallback: if dict keys are numeric indices (unlikely) try first value
            if isinstance(r, dict):
                # try first value in the dict
                vals = list(r.values())
                name = _normalize_name(vals[0] if vals else None)
        if name:
            vhse_lookup.setdefault(name, []).append(r)

    # Iterate frm rows and attempt to match
    for r in frm_rows:
        frm_name_raw = r.get(frm_key) if isinstance(r, dict) else None
        frm_name = _normalize_name(frm_name_raw)
        if not frm_name:
            # Skip rows without a name
            print(f"Skipping row with empty name in {frm_file}: {r}")
            continue

        matches = vhse_lookup.get(frm_name)
        if matches:
            for m in matches:
                print(f"Match found for '{frm_name_raw}':")
                print(f"  frm ({frm_file.name}): {r}")
                print(f"  vhse ({vhse_file.name}): {m}")
        else:
            print(f"Missing on {vhse_file.name}: employee '{frm_name_raw}' (from {frm_file.name})")

    # This test is informational and will always pass unless fixtures are missing
    assert True