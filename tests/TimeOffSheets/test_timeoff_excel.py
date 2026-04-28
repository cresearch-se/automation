# -*- coding: utf-8 -*-
import os
from pathlib import Path
from typing import Optional, List, Dict, Any

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

def _normalize_header(h: str) -> str:
    if h is None:
        return ""
    return "".join(ch for ch in h.lower() if ch.isalnum())

def _find_header(headers: List[str], fragments: List[str], fallback_index: Optional[int] = None) -> Optional[str]:
    """
    Try to find a header in headers that contains all fragments (case-insensitive, non-alnum ignored).
    If not found and fallback_index is provided and valid, return headers[fallback_index].
    """
    normalized_fragments = ["".join(f.split()).lower() for f in fragments]
    for h in headers:
        nh = _normalize_header(h)
        if all(frag.replace(" ", "") in nh for frag in normalized_fragments):
            return h
    if fallback_index is not None and 0 <= fallback_index < len(headers):
        return headers[fallback_index]
    return None

def _to_number(val: Any) -> Optional[float]:
    """
    Try to convert a value to float. Handles:
      * plain numbers (int/float)
      * strings with commas or NBSP thousands separators
      * parentheses for negatives: (1,234.56)
      * a single dash/hyphen treated as 0
      * percentages like "100.0 %" or "12.5%" -> returns numeric portion (100.0 or 12.5)
      * currency prefixes like dollar, pound, Euro signs
    Returns None only if conversion not possible.
    """
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)

    s = str(val).strip()
    if s == "":
        return None

    # Treat common placeholders as zero (change to None if preferred)
    if s in ("-", "N/A", "n/a"):
        return 0.0

    # Remove non-breaking space and surrounding whitespace
    s = s.replace("\u00A0", "").strip()

    # Handle percentage values: "100.0 %" or "12.5%"
    if "%" in s:
        s_pct = s.replace("%", "").strip()
        # s_pct = s_pct.replace(",", "").lstrip("$€£")
        s_pct = s_pct.replace(",", "").lstrip("$\u20AC\u00A3")
        try:
            return float(s_pct)
        except ValueError:
            return None

    # parentheses -> negative
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1].strip()

    # strip currency symbols
    s = s.lstrip("$\u20AC\u00A3")
    # s = s.lstrip("$€£")

    # remove thousands separators
    s = s.replace(",", "")

    # final attempt
    try:
        num = float(s)
        return -num if negative else num
    except ValueError:
        return None

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
            # Attempt to fallback: if dict keys are numeric indices try first value
            if isinstance(r, dict):
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

def test_compare_timeoff_numeric_columns():
    """
    For each employee in frm_Rpt_Det_given (column A), find matching row in VHSE_Holiday_Report_202510 (column A)
    and compare the following columns:
      * frm: Accrue Month (col D)          <-> vhse: HolAccrueMnth (col E)
      * frm: Accrue YTD                    <-> vhse: HolAccrueYTD
      * frm: Actual Month                  <-> vhse: HolActualMnth
      * frm: Actual YTD                    <-> vhse: HolActualYTD
      * frm: End Bal                       <-> vhse: End_Bal (or matching header)
    Print employee name and the mismatching values for any discrepancy. This test is informational.
    """
    base = Path(__file__).parent
    fixtures_dir = base / "fixtures"

    frm_file = _find_fixture(fixtures_dir, "frm_rpt_det_given")
    vhse_file = _find_fixture(fixtures_dir, "vhse_holiday_report_202510")

    assert frm_file is not None, f"Could not find frm_Rpt_Det_given file in {fixtures_dir}"
    assert vhse_file is not None, f"Could not find VHSE_Holiday_Report_202510 file in {fixtures_dir}"

    frm_headers = get_excel_headers(str(frm_file))
    vhse_headers = get_excel_headers(str(vhse_file))

    frm_rows = read_excel(str(frm_file))
    vhse_rows = read_excel(str(vhse_file))

    if not frm_rows:
        print(f"No data rows found in {frm_file}.")
        return

    # header resolution: try to find by semantic header first, fall back to index (0-based)
    frm_name_header = frm_headers[0] if frm_headers else ""
    vhse_name_header = vhse_headers[0] if vhse_headers else ""

    # frm expected columns (try semantic, fallback to column indices)
    frm_accrue_month_hdr = _find_header(frm_headers, ["accrue", "month"], fallback_index=3)
    frm_accrue_ytd_hdr = _find_header(frm_headers, ["accrue", "ytd"], fallback_index=4)
    frm_actual_month_hdr = _find_header(frm_headers, ["actual", "month"], fallback_index=5)
    frm_actual_ytd_hdr = _find_header(frm_headers, ["actual", "ytd"], fallback_index=6)
    frm_end_bal_hdr = _find_header(frm_headers, ["end", "bal"], fallback_index=None)

    # vhse expected columns (try semantic names that user specified)
    vhse_accrue_month_hdr = _find_header(vhse_headers, ["holaccrue", "mnth"], fallback_index=4)
    vhse_accrue_ytd_hdr = _find_header(vhse_headers, ["holaccrue", "ytd"], fallback_index=None)
    vhse_actual_month_hdr = _find_header(vhse_headers, ["holactual", "mnth"], fallback_index=None)
    vhse_actual_ytd_hdr = _find_header(vhse_headers, ["holactual", "ytd"], fallback_index=None)
    vhse_end_bal_hdr = _find_header(vhse_headers, ["end", "bal"], fallback_index=None)

    # Print diagnostic header resolution
    print("\nResolved headers:")
    print(f" frm name: {frm_name_header}")
    print(f" frm accrue month: {frm_accrue_month_hdr}")
    print(f" frm accrue ytd: {frm_accrue_ytd_hdr}")
    print(f" frm actual month: {frm_actual_month_hdr}")
    print(f" frm actual ytd: {frm_actual_ytd_hdr}")
    print(f" frm end bal: {frm_end_bal_hdr}")
    print(f" vhse name: {vhse_name_header}")
    print(f" vhse accrue month: {vhse_accrue_month_hdr}")
    print(f" vhse accrue ytd: {vhse_accrue_ytd_hdr}")
    print(f" vhse actual month: {vhse_actual_month_hdr}")
    print(f" vhse actual ytd: {vhse_actual_ytd_hdr}")
    print(f" vhse end bal: {vhse_end_bal_hdr}\n")

    # Build vhse lookup by normalized name
    vhse_lookup: Dict[str, List[Dict[str, Any]]] = {}
    for r in vhse_rows:
        name = _normalize_name(r.get(vhse_name_header) if isinstance(r, dict) else None)
        if not name and isinstance(r, dict):
            vals = list(r.values())
            name = _normalize_name(vals[0] if vals else None)
        if name:
            vhse_lookup.setdefault(name, []).append(r)

    mismatches = []

    for r in frm_rows:
        frm_name_raw = r.get(frm_name_header) if isinstance(r, dict) else None
        frm_name = _normalize_name(frm_name_raw)
        if not frm_name:
            continue

        matches = vhse_lookup.get(frm_name)
        if not matches:
            print(f"Missing on VHSE file for employee '{frm_name_raw}'")
            continue

        # compare against each matching vhse row
        for vhse_row in matches:
            comparisons = [
                (frm_accrue_month_hdr, vhse_accrue_month_hdr, "Accrue Month", "HolAccrueMnth"),
                (frm_accrue_ytd_hdr, vhse_accrue_ytd_hdr, "Accrue YTD", "HolAccrueYTD"),
                (frm_actual_month_hdr, vhse_actual_month_hdr, "Actual Month", "HolActualMnth"),
                (frm_actual_ytd_hdr, vhse_actual_ytd_hdr, "Actual YTD", "HolActualYTD"),
                (frm_end_bal_hdr, vhse_end_bal_hdr, "End Bal", "End_Bal"),
            ]
            for frm_hdr, vhse_hdr, friendly_frm, friendly_vhse in comparisons:
                if frm_hdr is None or vhse_hdr is None:
                    # if either header missing, print warning and skip comparison
                    print(f"Skipping comparison {friendly_frm} <-> {friendly_vhse} for '{frm_name_raw}': header not resolved.")
                    continue

                val_frm = r.get(frm_hdr) if isinstance(r, dict) else None
                val_vhse = vhse_row.get(vhse_hdr) if isinstance(vhse_row, dict) else None

                num_frm = _to_number(val_frm)
                num_vhse = _to_number(val_vhse)

                equal = False
                if num_frm is not None and num_vhse is not None:
                    # allow small rounding differences
                    equal = abs(num_frm - num_vhse) < 0.0001
                else:
                    # fallback to trimmed string comparison
                    sf = "" if val_frm is None else str(val_frm).strip()
                    sv = "" if val_vhse is None else str(val_vhse).strip()
                    equal = sf == sv
                """
                if not equal:
                    msg = (
                        f"Mismatch for '{frm_name_raw}' on {friendly_frm}:\n"
                        f"  frm ({frm_file.name}) [{frm_hdr}] = {val_frm}\n"
                        f"  vhse ({vhse_file.name}) [{vhse_hdr}] = {val_vhse}\n"
                    )
                    print(msg)
                    mismatches.append((frm_name_raw, friendly_frm, val_frm, val_vhse))
                """
                if not equal:
                    # Define column widths for a table-like look
                    # Name: 25 chars, Field: 15 chars, Values: 12 chars each
                    name_col = f"{str(frm_name_raw)[:23]:<25}"
                    field_col = f"{friendly_frm:<15}"
                    frm_val_col = f"FRM: {str(val_frm):<12}"
                    vhse_val_col = f"VHSE: {str(val_vhse):<12}"

                    msg = f"{name_col} | {field_col} | {frm_val_col} | {vhse_val_col}"
                    print(msg)
                    mismatches.append((frm_name_raw, friendly_frm, val_frm, val_vhse))

    print(f"\nTotal mismatches found: {len(mismatches)}")
    # Informational test: do not fail run, but if you want it to fail when mismatches exist,
    # replace the following line with: assert len(mismatches) == 0, f"{len(mismatches)} mismatches found"
    assert True