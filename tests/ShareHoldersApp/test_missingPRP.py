# -*- coding: latin-1 -*-

import pytest
from pathlib import Path
from typing import List, Dict

# run as  py -m pytest -s tests/ShareHoldersApp/test_missingPRP.py -k test_read_and_filter_emp_full_details > C:\Users\sowjikarumuri\Documents\Code\tests\ShareHoldersApp\output\prp_missing.txt

def read_emp_full_details(file_path: Path) -> List[Dict[str, str]]:
    """
    Read the EmpFullDetailsPRP.txt file and parse it into a list of dictionaries.

    File format: Folder Path||Emp ID||FY Folder Name||File Names

    This function attempts to read as UTF-8 and falls back to latin-1 (Windows-compatible)
    if a UnicodeDecodeError occurs.
    """
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # Try utf-8 first, fall back to latin-1
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            lines = f.read().splitlines()
    except UnicodeDecodeError:
        with open(file_path, "r", encoding="latin-1") as f:
            lines = f.read().splitlines()

    employees: List[Dict[str, str]] = []
    # skip header line if present
    start_idx = 1 if lines and "Folder Path" in lines[0] else 0
    for line in lines[start_idx:]:
        line = line.strip()
        if not line:
            continue

        parts = line.split("||")
        # ensure we have at least 4 parts; pad if necessary
        while len(parts) < 4:
            parts.append("")
        employees.append({
            "folder_path": parts[0].strip(),
            "emp_id": parts[1].strip(),
            "fy_folder_name": parts[2].strip(),
            "file_names": parts[3].strip()
        })

    return employees


def filter_employees_with_fy_folder(employees: List[Dict[str, str]]) -> List[Dict[str, str]]:
    """
    Return employees that have a non-empty FY folder name but NO file names.

    This treats whitespace-only file_names as empty.
    """
    result = []
    for e in employees:
        fy = (e.get("fy_folder_name") or "").strip()
        files = (e.get("file_names") or "").strip()
        if fy and not files:
            result.append({
                "folder_path": e.get("folder_path", ""),
                "emp_id": e.get("emp_id", ""),
                "fy_folder_name": fy
            })
    return result


def filter_employees_without_fy_folder(employees: List[Dict[str, str]]) -> List[Dict[str, str]]:
    """
    Return employees that do NOT have an FY folder name.
    """
    return [
        {"folder_path": e["folder_path"], "emp_id": e["emp_id"]}
        for e in employees if not e.get("fy_folder_name")
    ]


def test_read_and_filter_emp_full_details():
    """
    Read EmpFullDetailsPRP.txt, parse it, and print two lists:
    1. Employees WITH FY folder name and no files (folder_path, emp_id, fy_folder_name)
    2. Employees WITHOUT FY folder name (folder_path, emp_id)

    Use -s when running pytest to see the printed output.
    """
    fixture_file = Path.cwd() / "tests" / "ShareHoldersApp" / "fixtures" / "EmpFullDetailsPRP.txt"

    if not fixture_file.exists():
        pytest.skip(f"{fixture_file} not found; skipping test")

    employees = read_emp_full_details(fixture_file)
    assert employees, "No employees found in file"

    with_fy_folder = filter_employees_with_fy_folder(employees)
    without_fy_folder = filter_employees_without_fy_folder(employees)

    # Print results for inspection (use pytest -s to see output)
    print("\n" + "=" * 80)
    print(f"Total employees: {len(employees)}")
    print(f"Employees WITH FY Folder Name: {len(with_fy_folder)}")
    print(f"Employees WITHOUT FY Folder Name: {len(without_fy_folder)}")
    print("=" * 80 + "\n")

    print("LIST 1: EMPLOYEES WITH FY FOLDER NAME (Folder Path || Emp ID || FY Folder Name)")
    print("-" * 80)
    for emp in with_fy_folder:
        print(f"{emp['folder_path']} || {emp['emp_id']} || {emp['fy_folder_name']}")

    print("\n" + "=" * 80 + "\n")

    print("LIST 2: EMPLOYEES WITHOUT FY FOLDER NAME (Folder Path || Emp ID)")
    print("-" * 80)
    for emp in without_fy_folder:
        print(f"{emp['folder_path']} || {emp['emp_id']}")

    print("\n" + "=" * 80 + "\n")

    # Basic assertions
    assert len(with_fy_folder) + len(without_fy_folder) == len(employees), \
        "Sum of filtered lists should equal total employees"
    assert len(with_fy_folder) > 0, "Expected at least one employee with FY folder name"
    assert len(without_fy_folder) > 0, "Expected at least one employee without FY folder name"