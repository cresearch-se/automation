"""Refresh the monthly TeamworkDB utilization fixture from Power BI, then run the tests.

This automates the manual monthly workflow:
  1. export the Power BI paginated report (default = latest period in the dropdown),
  2. save it as tests/TeamworkDB/fixtures/Utilization_<YYYYMM>.xlsx,
  3. repoint FIXTURE_FILE in test_utilization_monthly.py at the new fixture,
  4. run run_test_utilization_monthly.py (Excel + HTML report to output/).

Run from the repo root (Code/) with the venv active:
    python tests/TeamworkDB/refresh_utilization_fixture.py            # previous month, then run tests
    python tests/TeamworkDB/refresh_utilization_fixture.py 202505     # force a specific YYYYMM
    python tests/TeamworkDB/refresh_utilization_fixture.py --no-run   # download + patch only
    python tests/TeamworkDB/refresh_utilization_fixture.py --check    # only verify Power BI access

Credentials come from config/creds/powerbi.env (see config/powerbi.env.example).
"""

import argparse
import datetime
import os
import re
import subprocess
import sys

# Repo root = .../Code  (this file is Code/tests/TeamworkDB/refresh_utilization_fixture.py)
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "src"))  # ensure the package is importable

from cornerstone_automation.utils import powerbi_utils  # noqa: E402

TEST_FILE = os.path.join(ROOT, "tests", "TeamworkDB", "test_utilization_monthly.py")
FIXTURES_DIR = os.path.join(ROOT, "tests", "TeamworkDB", "fixtures")
RUNNER = os.path.join(ROOT, "tests", "TeamworkDB", "run_test_utilization_monthly.py")


def previous_month_yyyymm(today=None):
    """Return the previous calendar month as 'YYYYMM' (e.g. run in June -> '202505')."""
    today = today or datetime.date.today()
    last_day_prev = today.replace(day=1) - datetime.timedelta(days=1)
    return last_day_prev.strftime("%Y%m")


def patch_fixture_file(test_file, fixture_relpath):
    """Repoint the FIXTURE_FILE assignment in the test file. Returns # replacements."""
    with open(test_file, "r", encoding="utf-8") as f:
        content = f.read()
    new_content, n = re.subn(
        r'(FIXTURE_FILE\s*=\s*)["\'].*?["\']',
        lambda m: f'{m.group(1)}"{fixture_relpath}"',
        content,
        count=1,
    )
    if n == 0:
        raise RuntimeError(
            f"Could not find a FIXTURE_FILE assignment to patch in {test_file}."
        )
    if new_content != content:
        with open(test_file, "w", encoding="utf-8") as f:
            f.write(new_content)
    return n


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("yyyymm", nargs="?",
                    help="Target period YYYYMM (default: previous month).")
    ap.add_argument("--no-run", action="store_true",
                    help="Download + patch only; do not run the tests.")
    ap.add_argument("--check", action="store_true",
                    help="Only verify Power BI access, then exit.")
    ap.add_argument("--period-param",
                    help="Name of the report's period parameter, used to FORCE a "
                         "specific month (needed only when YYYYMM differs from the "
                         "report's default period).")
    args = ap.parse_args()

    if args.check:
        powerbi_utils.check_access()
        return

    yyyymm = args.yyyymm or previous_month_yyyymm()
    if not re.fullmatch(r"\d{6}", yyyymm):
        sys.exit(f"Invalid YYYYMM: {yyyymm!r} (expected 6 digits like 202505).")

    fixture_name = f"Utilization_{yyyymm}.xlsx"
    out_path = os.path.join(FIXTURES_DIR, fixture_name)
    fixture_relpath = f"tests/TeamworkDB/fixtures/{fixture_name}"

    # By default we export the report's DEFAULT parameters (= latest period the
    # dropdown shows). Only when the user explicitly forces a month AND gives the
    # parameter name do we override it.
    parameter_values = None
    if args.yyyymm and args.period_param:
        parameter_values = [{"name": args.period_param, "value": yyyymm}]
    elif args.yyyymm:
        print(f"NOTE: exporting with the report's DEFAULT period. If the report "
              f"default is not {yyyymm}, the file contents won't match its name. "
              f"Pass --period-param <name> to force a specific month.")

    print(f"Exporting Power BI report -> {out_path}")
    powerbi_utils.export_report_to_file(out_path, fmt="XLSX", parameter_values=parameter_values)
    print(f"  Saved fixture ({os.path.getsize(out_path):,} bytes)")

    patch_fixture_file(TEST_FILE, fixture_relpath)
    print(f"  Patched FIXTURE_FILE -> {fixture_relpath}")

    if args.no_run:
        print("Done (--no-run): skipping the test run.")
        return

    print("Running utilization tests...")
    result = subprocess.run([sys.executable, RUNNER], cwd=ROOT)
    sys.exit(result.returncode)


if __name__ == "__main__":
    main()
