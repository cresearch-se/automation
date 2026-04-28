import os
from pathlib import Path
import pytest
from cornerstone_automation.utils.excel_utils import read_excel, get_excel_headers
from typing import List, Optional, Dict, Any, Set

def _find_header(headers: List[str], candidates: List[str]) -> Optional[str]:
    """
    Return the first header from headers that matches any candidate (case-insensitive,
    spaces/underscores ignored). Returns None if no match.
    """
    def normalize(s: str) -> str:
        return "".join(ch for ch in (s or "").lower() if ch.isalnum())

    normalized = {h: normalize(h) for h in headers}
    for cand in candidates:
        nc = normalize(cand)
        for h, nh in normalized.items():
            if nc == nh:
                return h
    return None

def test_reviewerlogin_pdf_filenames_in_wowd():
    """
    Existing test: For each ReviewerLogin in AllOfficers-SelfEval.xlsx, generate a file name
    in the format WOWD_<ReviewerLogin>_2025.pdf and compare against the Name column in WOWD.xlsx.
    """
    base = Path(__file__).parent
    fixtures_dir = base / "fixtures"

    seleval_file = fixtures_dir / "AllOfficers-SelfEval.xlsx"
    wowd_file = fixtures_dir / "WOWD.xlsx"

    assert seleval_file.exists(), f"Missing fixture: {seleval_file}"
    assert wowd_file.exists(), f"Missing fixture: {wowd_file}"

    # Read data
    seleval_rows = read_excel(str(seleval_file))
    wowd_rows = read_excel(str(wowd_file))

    # Resolve headers robustly
    seleval_headers = get_excel_headers(str(seleval_file))
    wowd_headers = get_excel_headers(str(wowd_file))

    reviewer_col = _find_header(seleval_headers, ["ReviewerLogin", "reviewer_login", "reviewer login", "reviewer"])
    assert reviewer_col is not None, "Could not find ReviewerLogin column in AllOfficers-SelfEval.xlsx"

    wowd_name_col = _find_header(wowd_headers, ["Name", "FileName", "Filename", "file name", "name"])
    assert wowd_name_col is not None, "Could not find Name column in WOWD.xlsx"

    # Get ReviewerLogin values and build expected filenames
    reviewer_logins = [str(row.get(reviewer_col)).strip() for row in seleval_rows if row.get(reviewer_col)]
    expected_filenames = {f"WOWD_{login}_2025.pdf" for login in reviewer_logins}

    # Get actual file names from WOWD.xlsx Name column
    wowd_names = {str(row.get(wowd_name_col)).strip() for row in wowd_rows if row.get(wowd_name_col)}

    # Find which expected filenames are present and which are missing
    present = sorted(expected_filenames & wowd_names)
    missing = sorted(expected_filenames - wowd_names)

    print("\nFiles present in WOWD folder for the given officers:")
    for fname in present:
        print(f"  {fname}")

    print("\nFiles NOT present in WOWD folder for the given officers:")
    for fname in missing:
        print(f"  {fname}")

    # informational: do not fail by default
    # assert not missing, f"Missing files: {missing}"

def test_reviewerlogin_staffreport_filenames_in_staffreportslist():
    """
    New test: For each ReviewerLogin in AllOfficers-SelfEval.xlsx, generate a file name
    in the format StaffReport_<ReviewerLogin>_2025.pdf and compare that set with the Name
    column in StaffReportsList.xlsx. Print files present and files missing.
    """
    base = Path(__file__).parent
    fixtures_dir = base / "fixtures"

    seleval_file = fixtures_dir / "AllOfficers-SelfEval.xlsx"
    staffreports_file = fixtures_dir / "StaffReportsList.xlsx"

    assert seleval_file.exists(), f"Missing fixture: {seleval_file}"
    assert staffreports_file.exists(), f"Missing fixture: {staffreports_file}"

    # Read data
    seleval_rows = read_excel(str(seleval_file))
    staff_rows = read_excel(str(staffreports_file))

    # Resolve headers robustly
    seleval_headers = get_excel_headers(str(seleval_file))
    staff_headers = get_excel_headers(str(staffreports_file))

    reviewer_col = _find_header(seleval_headers, ["ReviewerLogin", "reviewer_login", "reviewer login", "reviewer"])
    assert reviewer_col is not None, "Could not find ReviewerLogin column in AllOfficers-Self-Eval.xlsx"

    staff_name_col = _find_header(staff_headers, ["Name", "FileName", "Filename", "file name", "name"])
    assert staff_name_col is not None, "Could not find Name column in StaffReportsList.xlsx"

    # Build expected filenames from ReviewerLogin column
    def _clean(v):
        return "" if v is None else str(v).strip()

    reviewer_logins = [ _clean(r.get(reviewer_col)) for r in seleval_rows if _clean(r.get(reviewer_col)) ]
    expected_files = { f"StaffReport_{login}_2025.pdf" for login in reviewer_logins }

    # Get actual file names from StaffReportsList.xlsx Name column
    actual_files = { _clean(r.get(staff_name_col)) for r in staff_rows if _clean(r.get(staff_name_col)) }

    present = sorted(expected_files & actual_files)
    missing = sorted(expected_files - actual_files)

    print("\n--- Files present in StaffReportsList.xlsx ---")
    if present:
        for p in present:
            print(f"  {p}")
    else:
        print("  (none)")

    print("\n--- Files NOT present in StaffReportsList.xlsx ---")
    if missing:
        for m in missing:
            print(f"  {m}")
    else:
        print("  (none)")

    # informational by default; uncomment to enforce
    # assert not missing, f"Missing files in StaffReportsList.xlsx: {missing}"

# -----------------------------
# New reusable helpers + test using spec array
# -----------------------------

def build_expected_filenames_from_rows(rows: List[Dict[str, Any]], source_header: str, prefix: str, year: str, suffix: str = ".pdf") -> Set[str]:
    """
    Build set of filenames from rows using the value in source_header.
    Example: prefix="WOWD_", year="2025" -> "WOWD_<value>_2025.pdf"
    """
    def _clean(v):
        return "" if v is None else str(v).strip()
    values = [_clean(r.get(source_header)) for r in rows if _clean(r.get(source_header))]
    return {f"{prefix}{v}_{year}{suffix}" for v in values}

def compare_files_by_spec(spec: Dict[str, Any]) -> Dict[str, Any]:
    # ... (Keep the path and header resolution code at the top the same) ...
    source_path = Path(spec["source_path"])
    target_path = Path(spec["target_path"])
    prefix = spec.get("prefix", "")
    year = str(spec.get("year", "2025"))
    suffix = spec.get("suffix", ".pdf")

    source_rows = read_excel(str(source_path))
    target_rows = read_excel(str(target_path))
    source_headers = get_excel_headers(str(source_path))
    target_headers = get_excel_headers(str(target_path))

    source_header = _find_header(source_headers, spec.get("source_header_candidates", []))
    target_header = _find_header(target_headers, spec.get("target_header_candidates", []))

    # 1. Generate Expected Filenames
    expected_filenames = build_expected_filenames_from_rows(source_rows, source_header, prefix, year, suffix)
    
    # 2. Build a CASE-INSENSITIVE map of what is actually in the target Excel
    # Key: lowercase filename, Value: original cased filename
    def _clean(v):
        return "" if v is None else str(v).strip()

    actual_map = {
        _clean(r.get(target_header)).casefold(): _clean(r.get(target_header))
        for r in target_rows if _clean(r.get(target_header))
    }

    present = []
    missing = []

    # 3. Perform the comparison
    for expected in sorted(expected_filenames):
        lookup_key = expected.casefold()
        if lookup_key in actual_map:
            # Match found! Use the actual filename from the Excel for the report
            present.append(actual_map[lookup_key])
        else:
            # No match
            missing.append(expected)

    return {
        "spec": spec,
        "expected_count": len(expected_filenames),
        "actual_count": len(actual_map),
        "present": present,
        "missing": missing,
    }

def compare_files_by_spec_old(spec: Dict[str, Any]) -> Dict[str, Any]:
    """
    Compare source -> target using spec dictionary. Expected spec keys:
      - source_path: Path or str (Excel file containing source values)
      - source_header_candidates: List[str] header name candidates in source file
      - target_path: Path or str (Excel file containing target file names)
      - target_header_candidates: List[str] header candidates in target file
      - prefix: filename prefix (e.g. "WOWD_" or "StaffReport_")
      - year: string year (e.g. "2025")
      - suffix: optional (default ".pdf")
    Returns a result dict with 'present' and 'missing' sets and diagnostic info.
    """
    # Normalize paths
    source_path = Path(spec["source_path"])
    target_path = Path(spec["target_path"])
    prefix = spec.get("prefix", "")
    year = str(spec.get("year", "2025"))
    suffix = spec.get("suffix", ".pdf")

    assert source_path.exists(), f"Missing source fixture: {source_path}"
    assert target_path.exists(), f"Missing target fixture: {target_path}"

    source_rows = read_excel(str(source_path))
    target_rows = read_excel(str(target_path))

    source_headers = get_excel_headers(str(source_path))
    target_headers = get_excel_headers(str(target_path))

    source_header = _find_header(source_headers, spec.get("source_header_candidates", []))
    assert source_header is not None, f"Could not find source header in {source_path}: {spec.get('source_header_candidates', [])}"

    target_header = _find_header(target_headers, spec.get("target_header_candidates", []))
    assert target_header is not None, f"Could not find target header in {target_path}: {spec.get('target_header_candidates', [])}"

    expected = build_expected_filenames_from_rows(source_rows, source_header, prefix, year, suffix)
    # Normalize target values to strings and strip
    def _clean(v):
        return "" if v is None else str(v).strip()
    actual = {_clean(r.get(target_header)) for r in target_rows if _clean(r.get(target_header))}

    present = sorted(expected & actual)
    missing = sorted(expected - actual)

    return {
        "spec": spec,
        "expected_count": len(expected),
        "actual_count": len(actual),
        "present": present,
        "missing": missing,
    }

def get_comparison_specs(fixtures_dir: Path) -> List[Dict[str, Any]]:
    """
    Return an array of spec dictionaries you can extend in future.
    Current examples include the two cases already present in the file.
    """
    return [
        {
            "source_path": str(fixtures_dir / "AllOfficers-SelfEval.xlsx"),
            "source_header_candidates": ["ReviewerLogin", "reviewer_login", "reviewer"],
            "target_path": str(fixtures_dir / "WOWD.xlsx"),
            "target_header_candidates": ["Name", "FileName", "Filename", "name"],
            "prefix": "WOWD_",
            "year": "2025",
            "suffix": ".pdf",
        },
        {
            "source_path": str(fixtures_dir / "AllOfficers-SelfEval.xlsx"),
            "source_header_candidates": ["ReviewerLogin", "reviewer_login", "reviewer"],
            "target_path": str(fixtures_dir / "StaffReportsList.xlsx"),
            "target_header_candidates": ["Name", "FileName", "Filename", "name"],
            "prefix": "StaffReport_",
            "year": "2025",
            "suffix": ".pdf",
        },
        {
            "source_path": str(fixtures_dir / "AllOfficers-SelfEval.xlsx"),
            "source_header_candidates": ["ReviewerLogin", "reviewer_login", "reviewer"],
            "target_path": str(fixtures_dir / "IndividualReports.xlsx"),
            "target_header_candidates": ["Name", "FileName", "Filename", "name"],
            "prefix": "Ind_",
            "year": "2025",
            "suffix": ".pdf",
        },
        {
            "source_path": str(fixtures_dir / "AllOfficers-SelfEval.xlsx"),
            "source_header_candidates": ["ReviewerLogin", "reviewer_login", "reviewer"],
            "target_path": str(fixtures_dir / "TeamworkMatrixReports.xlsx"),
            "target_header_candidates": ["Name", "FileName", "Filename", "name"],
            "prefix": "TWMatrix_",
            "year": "2025",
            "suffix": ".pdf",
        },
        {
            "source_path": str(fixtures_dir / "AllOfficers-SelfEval.xlsx"),
            "source_header_candidates": ["ReviewerLogin", "reviewer_login", "reviewer"],
            "target_path": str(fixtures_dir / "TeamworkSurveyReports.xlsx"),
            "target_header_candidates": ["Name", "FileName", "Filename", "name"],
            "prefix": "TWS_",
            "year": "2025",
            "suffix": ".pdf",
        }
    ]

def test_compare_files_using_spec_array():
    """
    New test: iterate the spec array returned by get_comparison_specs() and run compare_files_by_spec()
    for each spec. Prints present/missing lists for each spec. Keeps test informational.
    """
    base = Path(__file__).parent
    fixtures_dir = base / "fixtures"
    specs = get_comparison_specs(fixtures_dir)

    any_missing = False
    for spec in specs:
        result = compare_files_by_spec(spec)
        prefix = spec.get("prefix", "")
        target_name = Path(spec["target_path"]).name
        print(f"\n--- Comparison for target file: {target_name} (prefix={prefix}) ---")
        print(f"Expected count: {result['expected_count']}, Actual count: {result['actual_count']}")
        if result["present"]:
            print("Present:")
            for p in result["present"]:
                print(f"  {p}")
        else:
            print("Present: (none)")

        if result["missing"]:
            any_missing = True
            print("Missing:")
            for m in result["missing"]:
                print(f"  {m}")
        else:
            print("Missing: (none)")

    # informational - do not fail by default
    # if you want to fail when any missing, uncomment:
    # assert not any_missing, "One or more comparisons had missing files"