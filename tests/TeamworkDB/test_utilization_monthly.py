import calendar
import datetime
import pandas as pd
import pytest
import os
from dotenv import load_dotenv

# Load environment variables from config/db.env
load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), '..', '..', 'config', 'db.env'))

from cornerstone_automation.utils.pandas_utilis import read_excel_file, get_excel_sheet_names, find_column_by_keywords, check_totals_match, compare_db_to_excel, safe_to_numeric
from cornerstone_automation.utils.db_utils import get_db_connection_from_env, call_stored_procedure

# ==========================================
# CHANGE ONLY THIS LINE EACH MONTH
# ==========================================
FIXTURE_FILE  = "tests/TeamworkDB/fixtures/Utilization_202603.xlsx"

# Derived automatically from the filename — no other changes needed
REPORT_MONTH  = os.path.basename(FIXTURE_FILE).replace("Utilization_", "").replace(".xlsx", "")  # e.g. "202602"
REPORT_YEAR   = REPORT_MONTH[:4]   # e.g. "2026"
REPORT_MM     = REPORT_MONTH[4:]   # e.g. "02"

_last_day     = calendar.monthrange(int(REPORT_YEAR), int(REPORT_MM))[1]
MONTHLY_START = f"{REPORT_YEAR}-{REPORT_MM}-01"
MONTHLY_END   = f"{REPORT_YEAR}-{REPORT_MM}-{_last_day:02d}"
YTD_START     = f"{REPORT_YEAR}-01-01"
YTD_END       = MONTHLY_END

# ==========================================
# LOCATION-LEVEL COMPARISON CONFIGURATIONS
# ==========================================
# You can add as many sheets or different files here as needed
CONFIG_LOCATION = {
    "comparisons": [
        {
            "name": "US_Utilization_Monthly",
            "sheet_name": f"{REPORT_MONTH}_US",
            "tolerance": 3.0,
            "expected_offices": ['CRB', 'CRCH', 'CRDC', 'CRLA', 'CRNY', 'CRSF', 'CRSV'],
            "expected_titles": [
                'Officer', 'Principal', 'Manager',
                'Associate-Exp', 'Associate-1st Yr', 'Associate',
                'Analyst-Exp', 'Analyst-1st Yr', 'Analyst'
            ],
            "header_row": 1,
            "subtotal_label": "OFFICE TOTAL",
            "grand_total_label": "US TOTAL"
        },
        {
            "name": "EU_Utilization_Monthly",
            "sheet_name": f"{REPORT_MONTH}_Europe",
            "tolerance": 3.0,
            "expected_offices": ['Brussels', 'London'],
            "expected_titles": [
                'Officer', 'Principal', 'Manager',
                'Associate-Exp', 'Associate-1st Yr', 'Associate',
                'Analyst-Exp', 'Analyst-1st Yr', 'Analyst'
            ],
            "header_row": 1,
            "subtotal_label": "OFFICE TOTAL",
            "grand_total_label": "Europe TOTAL"
        },
        {
            "name": "Total_Utilization_Monthly",
            "sheet_name": f"{REPORT_MONTH}_Total",
            "tolerance": 3.0,
            "expected_offices": ['US', 'Europe', 'Cornerstone Research'],
            "expected_titles": [
                'Officer', 'Principal', 'Manager',
                'Associate-Exp', 'Associate-1st Yr', 'Associate',
                'Analyst-Exp', 'Analyst-1st Yr', 'Analyst'
            ],
            "header_row": 2,
            "subtotal_label": ["US TOTAL", "Europe TOTAL"],
            "grand_total_label": "GRAND TOTAL",
            "ignore_subtotal_for": ["Cornerstone Research"]
        },
        {
            "name": "US_Utilization_YTD",
            "sheet_name": f"{REPORT_YEAR}_US",
            "tolerance": 3.0,
            "expected_offices": ['CRB', 'CRCH', 'CRDC', 'CRLA', 'CRNY', 'CRSF', 'CRSV'],
            "expected_titles": [
                'Officer', 'Principal', 'Manager',
                'Associate-Exp', 'Associate-1st Yr', 'Associate',
                'Analyst-Exp', 'Analyst-1st Yr', 'Analyst'
            ],
            "header_row": 1,
            "subtotal_label": "OFFICE TOTAL",
            "grand_total_label": "US TOTAL"
        },
        {
            "name": "EU_Utilization_YTD",
            "sheet_name": f"{REPORT_YEAR}_Europe",
            "tolerance": 3.0,
            "expected_offices": ['Brussels', 'London'],
            "expected_titles": [
                'Officer', 'Principal', 'Manager',
                'Associate-Exp', 'Associate-1st Yr', 'Associate',
                'Analyst-Exp', 'Analyst-1st Yr', 'Analyst'
            ],
            "header_row": 1,
            "subtotal_label": "OFFICE TOTAL",
            "grand_total_label": "Europe TOTAL"
        },
        {
            "name": "Total_Utilization_YTD",
            "sheet_name": f"{REPORT_YEAR}_Total",
            "tolerance": 3.0,
            "expected_offices": ['US', 'Europe', 'Cornerstone Research'],
            "expected_titles": [
                'Officer', 'Principal', 'Manager',
                'Associate-Exp', 'Associate-1st Yr', 'Associate',
                'Analyst-Exp', 'Analyst-1st Yr', 'Analyst'
            ],
            "header_row": 2,
            "subtotal_label": ["US TOTAL", "Europe TOTAL"],
            "grand_total_label": "GRAND TOTAL",
            "ignore_subtotal_for": ["Cornerstone Research"]
        }
    ]
}

SUBTOTAL_TITLES = ['Officer', 'Principal', 'Manager', 'Associate', 'Analyst']

CONFIG_EMPLOYEE = {
    "comparisons": [
        {
            "name": "Boston_Monthly",
            "sheet_name": "Month - CRB ",
            "grand_total_label": "OFFICE TOTAL"
        },
        {
            "name": "Chicago_Monthly",
            "sheet_name": "Month - CRCH",
            "grand_total_label": "OFFICE TOTAL"
        },
        {
            "name": "DC_Monthly",
            "sheet_name": "Month - CRDC",
            "grand_total_label": "OFFICE TOTAL"
        },
        {
            "name": "LA_Monthly",
            "sheet_name": "Month - CRLA",
            "grand_total_label": "OFFICE TOTAL"
        },
        {
            "name": "NY_Monthly",
            "sheet_name": "Month - CRNY",
            "grand_total_label": "OFFICE TOTAL"
        },
        {
            "name": "SF_Monthly",
            "sheet_name": "Month - CRSF",
            "grand_total_label": "OFFICE TOTAL"
        },
        {
            "name": "SV_Monthly",
            "sheet_name": "Month - CRSV",
            "grand_total_label": "OFFICE TOTAL"
        },
        {
            "name": "Brussels_Monthly",
            "sheet_name": "Month - CRBE",
            "grand_total_label": "OFFICE TOTAL"
        },
        {
            "name": "UK_Monthly",
            "sheet_name": "Month - CRUK",
            "grand_total_label": "OFFICE TOTAL"
        },
        {
            "name": "Boston_YTD",
            "sheet_name": "YTD - CRB ",
            "grand_total_label": "OFFICE TOTAL"
        },
        {
            "name": "Chicago_YTD",
            "sheet_name": "YTD - CRCH",
            "grand_total_label": "OFFICE TOTAL"
        },
        {
            "name": "DC_YTD",
            "sheet_name": "YTD - CRDC",
            "grand_total_label": "OFFICE TOTAL"
        },
        {
            "name": "LA_YTD",
            "sheet_name": "YTD - CRLA",
            "grand_total_label": "OFFICE TOTAL"
        },
        {
            "name": "NY_YTD",
            "sheet_name": "YTD - CRNY",
            "grand_total_label": "OFFICE TOTAL"
        },
        {
            "name": "SF_YTD",
            "sheet_name": "YTD - CRSF",
            "grand_total_label": "OFFICE TOTAL"
        },
        {
            "name": "SV_YTD",
            "sheet_name": "YTD - CRSV",
            "grand_total_label": "OFFICE TOTAL"
        },
        {
            "name": "Brussels_YTD",
            "sheet_name": "YTD - CRBE",
            "grand_total_label": "OFFICE TOTAL"
        },
        {
            "name": "UK_YTD",
            "sheet_name": "YTD - CRUK",
            "grand_total_label": "OFFICE TOTAL"
        },
        {
            "name": "DataScience_Monthly",
            "sheet_name": "Month - Data Science",
            "grand_total_label": "TOTAL"
        },
        {
            "name": "AppliedResearch_Monthly",
            "sheet_name": "Month - Applied Research",
            "grand_total_label": "TOTAL"
        },
        {
            "name": "DataScience_YTD",
            "sheet_name": "YTD - Data Science",
            "grand_total_label": "TOTAL"
        },
        {
            "name": "AppliedResearch_YTD",
            "sheet_name": "YTD - Applied Research",
            "grand_total_label": "TOTAL"
        }
    ]
}

OFFICE_MAP = {
    'Boston': 'CRB',
    'Chicago': 'CRCH',
    'Washington, D.C.': 'CRDC',
    'Los Angeles': 'CRLA',
    'New York': 'CRNY',
    'San Francisco': 'CRSF',
    'Silicon Valley': 'CRSV',
    'Europe': 'Europe'
}

NUMERIC_COLS = ['Target_Hours', 'Target_Rev', 'Actual_Hours', 'Standard_Rev']

# DB column name → Excel column name mappings (used when comparing SP results to Excel)
COLUMN_MAP_BY_LOCATION = {
    "Office_Code":  "Office",
    "VA_Target_Hrs": "Target_Hours",
    "TargetRev":    "Target_Rev",
    "Actual_hrs":   "Actual_Hours",
    "Actual_Amount": "Standard_Rev"
}

COLUMN_MAP_BY_EMPLOYEE = {
    "Office_Code":     "Office",
    "EMPLOYEE_CODE":   "EmpNo",
    "EMPLOYEE_NAME":   "Name",
    "VA_Target_Hrs":   "Target_Hours",
    "TargetRev":       "Target_Rev",
    "Actual_hrs":      "Actual_Hours",
    "Actual_Amount":   "Standard_Rev"
}

# Database connection settings — read from config/db.env (loaded via db_utils import)
DB_SERVER   = os.getenv("T1_DB_SERVER")
DB_DATABASE = os.getenv("ReportDevl_DATABASE")
SP_NAME     = "SP_Utilization_Validation"

# Join keys used when merging Excel vs DB data (constant per test type)
JOIN_KEYS_BY_LOCATION = ["Office", "Title"]
JOIN_KEYS_BY_EMPLOYEE = ["EmpNo"]

# US office locations for employee-level DB comparison tests
US_EMPLOYEE_LOCATION_CONFIG = [
    {"name": "Boston",    "office_code": "CRB",  "monthly_sheet": "Month - CRB ",  "ytd_sheet": "YTD - CRB "},
    {"name": "Chicago",   "office_code": "CRCH", "monthly_sheet": "Month - CRCH",  "ytd_sheet": "YTD - CRCH"},
    {"name": "DC",        "office_code": "CRDC", "monthly_sheet": "Month - CRDC",  "ytd_sheet": "YTD - CRDC"},
    {"name": "LA",        "office_code": "CRLA", "monthly_sheet": "Month - CRLA",  "ytd_sheet": "YTD - CRLA"},
    {"name": "NY",        "office_code": "CRNY", "monthly_sheet": "Month - CRNY",  "ytd_sheet": "YTD - CRNY"},
    {"name": "SF",        "office_code": "CRSF", "monthly_sheet": "Month - CRSF",  "ytd_sheet": "YTD - CRSF"},
    {"name": "SV",        "office_code": "CRSV", "monthly_sheet": "Month - CRSV",  "ytd_sheet": "YTD - CRSV"},
    {"name": "EU-UK",  "office_code": "Europe", "monthly_sheet": "Month - CRUK", "ytd_sheet": "YTD - CRUK", "filter_by_excel_empnos": True},
    {"name": "EU-BE",  "office_code": "Europe", "monthly_sheet": "Month - CRBE", "ytd_sheet": "YTD - CRBE", "filter_by_excel_empnos": True}
]

# ==========================================
# HELPER FUNCTIONS (Normalization & Row Tagging)
# ==========================================

def normalize_excel_data_by_location(file_path, sheet_name, header_row=1):
    """
    Transforms the Month Utilization Summary Excel sheet into a clean flat table.

    The sheet has:
      - A title row at the top (e.g. "Month Utilization Summary - 202602")
      - A header row with: OFFC, Title, Target Hours, Target Revenue,
                           Actual Hours, Standard Revenue, % of Target Hours, % of Target Revenue
      - Office section headers (e.g. "Boston", "Chicago") with no numeric data
      - Data rows: one row per title within each office
      - OFFICE TOTAL rows at the end of each section (excluded)

    Returns a DataFrame with columns:
      Office, Title, Target_Hours, Target_Rev, Actual_Hours, Standard_Rev
    """
    # Row 0 is the sheet title; header_row points to the actual column header row
    df = read_excel_file(file_path, sheet_name=sheet_name, header=header_row)

    # 3. Locate columns by matching header names (case-insensitive, partial match)
    offc_col         = find_column_by_keywords(df, ['offc']) or find_column_by_keywords(df, ['office'])
    title_col        = find_column_by_keywords(df, ['title'])
    target_hours_col = find_column_by_keywords(df, ['target', 'hour'])
    target_rev_col   = find_column_by_keywords(df, ['target', 'rev'])
    actual_hours_col = find_column_by_keywords(df, ['actual', 'hour'])
    standard_rev_col = find_column_by_keywords(df, ['standard', 'rev'])

    missing = [name for name, col in [
        ('OFFC', offc_col), ('Title', title_col),
        ('Target Hours', target_hours_col), ('Target Revenue', target_rev_col),
        ('Actual Hours', actual_hours_col), ('Standard Revenue', standard_rev_col)
    ] if col is None]
    if missing:
        raise ValueError(f"Could not find columns {missing} in sheet '{sheet_name}'. Found: {list(df.columns)}")

    # 4. Forward-fill office names — office header rows have no numeric Target Hours
    df['Office_Name_Raw'] = df[offc_col].where(pd.to_numeric(df[target_hours_col], errors='coerce').isna())
    df['Office_Name_Raw'] = df['Office_Name_Raw'].ffill().str.strip()

    # 5. Map office names to codes
    df['Office'] = df['Office_Name_Raw'].map(OFFICE_MAP).fillna(df['Office_Name_Raw'].str.strip())

    # 6. Keep only data rows (Target Hours is numeric)
    df_clean = df[pd.to_numeric(df[target_hours_col], errors='coerce').notna()].copy()

    # 7. Remove OFFICE TOTAL rows
    df_clean = df_clean[~df_clean[title_col].astype(str).str.contains('TOTAL', na=False, case=False)]

    # 8. Build final clean DataFrame
    df_final = df_clean[['Office', title_col, target_hours_col, target_rev_col, actual_hours_col, standard_rev_col]].copy()
    df_final.columns = ['Office', 'Title', 'Target_Hours', 'Target_Rev', 'Actual_Hours', 'Standard_Rev']
    df_final['Title'] = df_final['Title'].str.strip()

    for col in NUMERIC_COLS:
        df_final[col] = pd.to_numeric(df_final[col], errors='coerce')

    return df_final.reset_index(drop=True)

def normalize_excel_data_with_totals(file_path, sheet_name, subtotal_label, grand_total_label, header_row=1):
    """
    Same as normalize_excel_data_by_location but keeps subtotal and grand total rows.
    Each row is tagged with a 'row_type':
      - 'data'         : regular title rows
      - 'subtotal'     : OFFICE TOTAL / US TOTAL / EU TOTAL rows within a section
      - 'grand_total'  : the final grand total row for the whole sheet

    :param subtotal_label:    Label(s) to match subtotal rows. Either a single string
                              (e.g. "OFFICE TOTAL" for US/EU sheets) or a list of strings
                              (e.g. ["US TOTAL", "Europe TOTAL", "Cornerstone Research"] for Total sheets).
    :param grand_total_label: Exact label to match the grand total row (e.g. "GRAND TOTAL").
    """
    df = read_excel_file(file_path, sheet_name=sheet_name, header=header_row)

    offc_col         = find_column_by_keywords(df, ['offc']) or find_column_by_keywords(df, ['office'])
    title_col        = find_column_by_keywords(df, ['title'])
    target_hours_col = find_column_by_keywords(df, ['target', 'hour'])
    target_rev_col   = find_column_by_keywords(df, ['target', 'rev'])
    actual_hours_col = find_column_by_keywords(df, ['actual', 'hour'])
    standard_rev_col = find_column_by_keywords(df, ['standard', 'rev'])

    missing = [name for name, col in [
        ('OFFC', offc_col), ('Title', title_col),
        ('Target Hours', target_hours_col), ('Target Revenue', target_rev_col),
        ('Actual Hours', actual_hours_col), ('Standard Revenue', standard_rev_col)
    ] if col is None]
    if missing:
        raise ValueError(f"Could not find columns {missing} in sheet '{sheet_name}'. Found: {list(df.columns)}")

    # Forward-fill office names
    df['Office_Name_Raw'] = df[offc_col].where(pd.to_numeric(df[target_hours_col], errors='coerce').isna())
    df['Office_Name_Raw'] = df['Office_Name_Raw'].ffill().str.strip()
    df['Office'] = df['Office_Name_Raw'].map(OFFICE_MAP).fillna(df['Office_Name_Raw'].str.strip())

    # Keep all rows that have numeric Target Hours
    df_clean = df[pd.to_numeric(df[target_hours_col], errors='coerce').notna()].copy()

    # Tag each row by type
    title_upper = df_clean[title_col].astype(str).str.strip().str.upper()
    grand_upper = grand_total_label.strip().upper()

    df_clean['row_type'] = 'data'
    df_clean.loc[title_upper == grand_upper, 'row_type'] = 'grand_total'
    if isinstance(subtotal_label, list):
        sub_uppers = [s.strip().upper() for s in subtotal_label]
        df_clean.loc[title_upper.isin(sub_uppers), 'row_type'] = 'subtotal'
    else:
        df_clean.loc[title_upper == subtotal_label.strip().upper(), 'row_type'] = 'subtotal'

    # Build final DataFrame
    df_final = df_clean[['Office', title_col, target_hours_col, target_rev_col,
                          actual_hours_col, standard_rev_col, 'row_type']].copy()
    df_final.columns = ['Office', 'Title', 'Target_Hours', 'Target_Rev',
                        'Actual_Hours', 'Standard_Rev', 'row_type']

    df_final['Title'] = df_final['Title'].str.strip()

    for col in NUMERIC_COLS:
        df_final[col] = pd.to_numeric(df_final[col], errors='coerce')

    return df_final.reset_index(drop=True)

def normalize_excel_data_by_employee(file_path, sheet_name, header_row=1, grand_total_label="OFFICE TOTAL"):
    """
    Normalizes an employee-level utilization sheet (e.g. "Month - CRB").

    Sheet structure (1-indexed):
      - Row 1: Skip (report title)
      - Row 2: Column headers → header_row=1
      - Row 3 becomes index 0 after read → dropped via df.iloc[1:]
      - Row 4+: Title rows + employee rows + subtotals + grand total

    Title is forward-filled down to employee rows below each title header.

    Row types tagged:
      - 'data'       : rows with a valid Emp#
      - 'subtotal'   : no Emp#, numeric values (total per title group)
      - 'grand_total': row matching grand_total_label (e.g. "OFFICE TOTAL" or "TOTAL")

    Returns DataFrame with columns:
      EmpNo, Name, Title, Target_Hours, Target_Rev, Actual_Hours, Standard_Rev, row_type
    """
    df = read_excel_file(file_path, sheet_name=sheet_name, header=header_row)

    # Drop location name row (sheet row 3)
    df = df.iloc[1:].reset_index(drop=True)

    # Locate numeric columns first
    target_hours_col = find_column_by_keywords(df, ['target', 'hour'])
    target_rev_col   = find_column_by_keywords(df, ['target', 'rev'])
    actual_hours_col = find_column_by_keywords(df, ['actual', 'hour'])
    standard_rev_col = find_column_by_keywords(df, ['standard', 'rev'])

    # The header row has a merged cell; the 'emp' column contains either Title text
    # (non-numeric, used as section headers) or numeric EmpNo (employee data rows)
    combo_col = find_column_by_keywords(df, ['emp'])

    # Name column is the unnamed column just before Target Hours
    target_hours_idx = list(df.columns).index(target_hours_col)
    name_col = df.columns[target_hours_idx - 1]

    missing = [label for label, col in [
        ('Combo (Title/Emp#)', combo_col), ('Target Hours', target_hours_col),
        ('Target Revenue', target_rev_col), ('Actual Hours', actual_hours_col),
        ('Standard Revenue', standard_rev_col)
    ] if col is None]
    if missing:
        raise ValueError(f"Could not find columns {missing} in sheet '{sheet_name}'. Found: {list(df.columns)}")

    # Forward-fill Title: extract text (non-numeric, non-null) values from combo column as Title header rows
    df['Title_Filled'] = df[combo_col].where(
        pd.to_numeric(df[combo_col], errors='coerce').isna() & df[combo_col].notna()
    ).ffill().str.strip()

    # Emp# = numeric values in the combo column
    df['EmpNo_Raw'] = pd.to_numeric(df[combo_col], errors='coerce')

    # Keep only rows with numeric Target Hours
    df_clean = df[pd.to_numeric(df[target_hours_col], errors='coerce').notna()].copy()

    # Tag row types
    title_upper = df_clean['Title_Filled'].astype(str).str.strip().str.upper()
    has_emp = df_clean['EmpNo_Raw'].notna()

    grand_upper = grand_total_label.strip().upper()
    df_clean['row_type'] = 'data'
    df_clean.loc[~has_emp & (title_upper != grand_upper), 'row_type'] = 'subtotal'
    df_clean.loc[title_upper == grand_upper, 'row_type'] = 'grand_total'

    # Build final DataFrame
    df_final = df_clean[['Title_Filled', 'EmpNo_Raw', name_col, target_hours_col,
                          target_rev_col, actual_hours_col, standard_rev_col, 'row_type']].copy()
    df_final.columns = ['Title', 'EmpNo', 'Name', 'Target_Hours', 'Target_Rev',
                        'Actual_Hours', 'Standard_Rev', 'row_type']

    # Canonical type: EmpNo as str — take original value from combo column to preserve leading zeros
    # (e.g. "0543" stays "0543"; float 543.0 becomes "543" via split on '.')
    df_final['EmpNo'] = df_clean[combo_col].where(df_clean['EmpNo_Raw'].notna()).apply(
        lambda x: str(x).split('.')[0] if pd.notna(x) else x
    )

    for col in NUMERIC_COLS:
        df_final[col] = pd.to_numeric(df_final[col], errors='coerce')

    return df_final.reset_index(drop=True)


# ==========================================
# TEST SUITE
# ==========================================

@pytest.mark.parametrize("comp", CONFIG_LOCATION["comparisons"], ids=[c["name"] for c in CONFIG_LOCATION["comparisons"]])
def test_format_validations_by_location(comp):
    """
    Parametrized test that runs a fresh comparison for every entry in CONFIG_LOCATION.
    """
    #print(f"\nRunning Comparison: {comp['sheet_name']}")
    
    # 1. Check file existence
    assert os.path.exists(FIXTURE_FILE), f"File not found: {FIXTURE_FILE}"

    # 2. Normalize Excel Data
    df_excel = normalize_excel_data_by_location(FIXTURE_FILE, comp['sheet_name'], comp['header_row'])

    #print(f"\n--- Loaded sheet: {comp.get('sheet_name')} ---")
    #print("\nColumns:", list(df_excel.columns))
    #preview_rows = 200
    #print(df_excel.head(preview_rows).to_string(index=False, float_format='{:,.2f}'.format))

    # 3. Validate all expected offices and titles are present
    errors = []
    missing_offices = set()

    actual_offices = df_excel['Office'].unique().tolist()
    for office in comp['expected_offices']:
        if office not in actual_offices:
            errors.append(f"[MISSING OFFICE] {office}")
            missing_offices.add(office)

    for office in comp['expected_offices']:
        if office in missing_offices:
            continue
        actual_titles = df_excel[df_excel['Office'] == office]['Title'].str.strip().tolist()
        for title in comp['expected_titles']:
            if title not in actual_titles:
                errors.append(f"[MISSING TITLE] {office} — {title}")

    # 4. Validate title order per office matches expected order
    for office in comp['expected_offices']:
        if office in missing_offices:
            continue
        actual_titles = df_excel[df_excel['Office'] == office]['Title'].str.strip().tolist()
        actual_ordered = [t for t in actual_titles if t in comp['expected_titles']]
        if actual_ordered != comp['expected_titles']:
            errors.append(f"[WRONG ORDER] {office} — expected {comp['expected_titles']} but got {actual_ordered}")

    # 5. Validate no blank or zero values in numeric columns
    for _, row in df_excel.iterrows():
        for col in NUMERIC_COLS:
            val = row[col]
            if pd.isna(val) or val == 0:
                status = "blank" if pd.isna(val) else "zero"
                errors.append(f"[{status.upper()} VALUE] {row['Office']} — {row['Title']} — {col} is {status}")

    if errors:
        print(f"\t")
        for error in errors:
            print(f"  {error}")
        pytest.fail(f"{len(errors)} error(s) found")
    else:
        print(f"\n[PASS] All {len(comp['expected_offices'])} offices and all {len(comp['expected_titles'])} titles are present as expected, and no blank or zero values found.")

@pytest.mark.parametrize("comp", CONFIG_LOCATION["comparisons"], ids=[c["name"] for c in CONFIG_LOCATION["comparisons"]])
def test_totals_validation_by_location(comp):
    """
    Validates that subtotal and grand total rows in each sheet match the calculated sums.

    For all sheets:
      - Each office subtotal row must equal the sum of its data rows.
      - The grand total row must equal the sum of all office subtotal rows.

    subtotal_label can be a single string (US/EU sheets: "OFFICE TOTAL")
    or a list (Total sheets: ["US TOTAL", "Europe TOTAL", "Cornerstone Research"]).
    """
    #print(f"\nRunning Totals Validation: {comp['sheet_name']}")

    assert os.path.exists(FIXTURE_FILE), f"File not found: {FIXTURE_FILE}"

    df = normalize_excel_data_with_totals(
        FIXTURE_FILE, comp['sheet_name'],
        comp['subtotal_label'], comp['grand_total_label'], comp['header_row']
    )

    errors = []
    tolerance = comp['tolerance']

    # 1. For each office: sum of data rows == its subtotal row
    for office in comp['expected_offices']:
        if office in comp.get('ignore_subtotal_for', []):
            continue
        df_data     = df[(df['Office'] == office) & (df['row_type'] == 'data') & (df['Title'].isin(SUBTOTAL_TITLES))]
        df_subtotal = df[(df['Office'] == office) & (df['row_type'] == 'subtotal')]

        if df_subtotal.empty:
            errors.append(f"[MISSING SUBTOTAL ROW] {office} — subtotal row not found")
            continue

        errors.extend(check_totals_match(df_data, df_subtotal, NUMERIC_COLS, office, tolerance))

    # 2. Grand total == sum of all subtotal rows
    df_all_subtotals = df[df['row_type'] == 'subtotal']
    df_grand         = df[df['row_type'] == 'grand_total']

    if df_grand.empty:
        errors.append(f"[MISSING GRAND TOTAL ROW] '{comp['grand_total_label']}' row not found")
    else:
        errors.extend(check_totals_match(df_all_subtotals, df_grand, NUMERIC_COLS,
                                         comp['grand_total_label'], tolerance))

    if errors:
        print(f"\t")
        for error in errors:
            print(f"  {error}")
        pytest.fail(f"{len(errors)} error(s) found")
    else:
        print(f"\n[PASS] All subtotals and grand total verified correctly for '{comp['sheet_name']}'.")

@pytest.mark.parametrize("comp", CONFIG_EMPLOYEE["comparisons"], ids=[c["name"] for c in CONFIG_EMPLOYEE["comparisons"]])
def test_format_validations_by_employee(comp):
    """
    Format validations for employee-level utilization sheets (e.g. "Month - CRB").

    Validates on data rows only:
      1. No blank values in Title, EmpNo, Name, or any numeric column
      2. EmpNo is unique across the sheet
      3. EmpNo + Name combo is unique across the sheet
      4. Subtotal rows match sum of data rows for each title group
      5. Grand total (OFFICE TOTAL) matches sum of all subtotal rows
    """
    # print(f"\nRunning Format Validation: {comp['name']}")

    assert os.path.exists(FIXTURE_FILE), f"File not found: {FIXTURE_FILE}"

    df = normalize_excel_data_by_employee(FIXTURE_FILE, comp['sheet_name'], grand_total_label=comp['grand_total_label'])
    df_data     = df[df['row_type'] == 'data'].copy()
    df_subtotal = df[df['row_type'] == 'subtotal'].copy()
    df_grand    = df[df['row_type'] == 'grand_total'].copy()
    tolerance   = comp.get('tolerance', 0.1)

    errors = []

    # 1. No blank values in key columns
    for col in ['Title', 'EmpNo', 'Name'] + NUMERIC_COLS:
        blank_rows = df_data[df_data[col].isna() | (df_data[col].astype(str).str.strip() == '')]
        for _, row in blank_rows.iterrows():
            errors.append(f"[BLANK VALUE] EmpNo={row['EmpNo']} | Name={row['Name']} | {col} is blank")

    # 2. EmpNo is unique
    dup_empnos = df_data[df_data['EmpNo'].duplicated(keep=False)]['EmpNo'].unique()
    for emp in dup_empnos:
        names = df_data[df_data['EmpNo'] == emp]['Name'].tolist()
        errors.append(f"[DUPLICATE EMP#] EmpNo={emp} appears {len(names)} time(s): {names}")

    # 3. EmpNo + Name combo is unique
    dup_combos = df_data[df_data.duplicated(subset=['EmpNo', 'Name'], keep=False)][['EmpNo', 'Name']].drop_duplicates()
    for _, row in dup_combos.iterrows():
        errors.append(f"[DUPLICATE EMP# + NAME] EmpNo={row['EmpNo']} | Name={row['Name']}")

    # 4. Subtotals: sum of data rows per title == subtotal row for that title
    for title in df_subtotal['Title'].unique():
        df_title_data     = df_data[df_data['Title'] == title]
        df_title_subtotal = df_subtotal[df_subtotal['Title'] == title]
        if df_title_subtotal.empty:
            errors.append(f"[MISSING SUBTOTAL] Title='{title}' — subtotal row not found")
            continue
        errors.extend(check_totals_match(df_title_data, df_title_subtotal, NUMERIC_COLS, f"Subtotal: {title}", tolerance))

    # 5. Grand total: sum of all subtotal rows == grand total row
    if df_grand.empty:
        errors.append(f"[MISSING GRAND TOTAL] '{comp['grand_total_label']}' row not found")
    else:
        errors.extend(check_totals_match(df_subtotal, df_grand, NUMERIC_COLS, comp['grand_total_label'], tolerance))

    if errors:
        print(f"\t")
        for error in errors:
            print(f"  {error}")
        pytest.fail(f"{len(errors)} error(s) found")
    else:
        print(f"\n[PASS] All format validations passed for '{comp['sheet_name']}'.")
        print(f"[PASS] All Subtotals and grand total verified for '{comp['sheet_name']}'.")


# ==========================================
# ==========================================
# DB COMPARISON HELPERS
# ==========================================

def fetch_summary_from_db(start_date, end_date):
    """Call SP_Utilization_Validation with Type='Summary' and return the first result set as a DataFrame."""
    print(f"\nFetching summary data from DB for {DB_SERVER} to {DB_DATABASE}...")
    conn = get_db_connection_from_env(DB_SERVER, DB_DATABASE, trusted_connection=True)
    result_sets = call_stored_procedure(
        conn, SP_NAME,
        named_params={"StartDate": start_date, "EndDate": end_date, "Type": "Summary"},
        as_dataframe=True
    )
    conn.close()
    return result_sets[0]


def fetch_detail_from_db(start_date, end_date):
    """Call SP_Utilization_Validation with Type='Detail' and return the first result set as a DataFrame."""
    conn = get_db_connection_from_env(DB_SERVER, DB_DATABASE, trusted_connection=True)
    result_sets = call_stored_procedure(
        conn, SP_NAME,
        named_params={"StartDate": start_date, "EndDate": end_date, "Type": "Detail"},
        as_dataframe=True
    )
    conn.close()
    return result_sets[0]


def normalize_db_data_by_employee(df_db, numeric_cols):
    """
    Normalize DB data for employee-level comparisons.

    Canonical types (must match normalize_excel_data_by_employee output):
      - EmpNo: str — strip trailing .0 via split('.')[0] to preserve leading zeros (e.g. "0543" stays "0543")
      - Numeric columns: float
    """
    df = df_db.copy()

    if 'EmpNo' in df.columns:
        df['EmpNo'] = df['EmpNo'].apply(
            lambda x: str(x).split('.')[0] if pd.notna(x) and str(x).strip() != '' else x
        )

    for col in numeric_cols:
        if col in df.columns:
            df[col] = safe_to_numeric(df[col], remove_commas=True)
    
    return df


def normalize_db_data_by_location(df_db, numeric_cols):
    """
    Normalize DB data for location-level comparisons.
    
    Converts:
      - Numeric columns: comma-formatted strings → float
    """
    df = df_db.copy()
    
    for col in numeric_cols:
        if col in df.columns:
            df[col] = safe_to_numeric(df[col], remove_commas=True)
    
    return df


def print_error_summary(errors, label):
    """
    Prints errors in a formatted table matching the QA report style.
      - Header block with report month, sheet, and run date
      - MISSING rows printed as plain lines
      - VALUE MISMATCH rows printed as a table with individual key columns
    """
    missing  = [e for e in errors if '[MISSING' in e]
    mismatch = [e for e in errors if '[VALUE MISMATCH]' in e]

    month_name = calendar.month_name[int(REPORT_MM)]
    run_date   = datetime.date.today().strftime("%Y-%m-%d")

    print(f"\n  UTILIZATION QA — VALUE MISMATCHES")
    print(f"  Report: {month_name} {REPORT_YEAR} | Sheet: {label} | Run: {run_date}")

    if missing:
        print(f"\n  MISSING ROWS:")
        for e in missing:
            print(f"    {e}")

    if mismatch:
        rows = []
        key_names = None
        for e in mismatch:
            body      = e[len('[VALUE MISMATCH] '):]
            parts     = body.split(' | ')
            col_idx   = next(i for i, p in enumerate(parts) if ': XLS=' in p)
            key_parts = parts[:col_idx]

            # Extract key column names and values from "Key=Value" pairs
            kv = [p.strip().split('=', 1) for p in key_parts]
            if key_names is None:
                key_names = [k for k, v in kv]
            key_vals = [v for k, v in kv]

            col_part = parts[col_idx]
            col_name = col_part.split(':')[0].strip()
            rest     = col_part.split(': ', 1)[1]
            xls_val  = rest.split(' vs ')[0].replace('XLS=', '').strip()
            db_part  = rest.split(' vs ')[1]
            db_val   = db_part.split(' (diff=')[0].replace('DB=', '').strip()
            diff_str = db_part.split('(diff=')[1].rstrip(')')
            diff_val, pct = diff_str.split(', ')
            rows.append(key_vals + [col_name, xls_val, db_val, diff_val.strip(), pct.strip()])

        all_cols   = key_names + ['Column', 'Util Report', 'DB', 'Diff', 'Diff%']
        col_widths = [max(len(all_cols[i]), max(len(str(r[i])) for r in rows)) for i in range(len(all_cols))]

        header = ' | '.join(f"{all_cols[i]:<{col_widths[i]}}" for i in range(len(all_cols)))
        sep    = '-+-'.join('-' * col_widths[i] for i in range(len(all_cols)))
        print(f"\n  {header}")
        print(f"  {sep}")
        for r in rows:
            row_str = ' | '.join(f"{str(r[i]):<{col_widths[i]}}" for i in range(len(all_cols)))
            print(f"  {row_str}")


def run_location_comparison(df_excel, df_db, label, tolerance=0.1):
    """Rename DB columns, normalize data types, outer-merge with Excel data, and assert no mismatches."""
    df_db = df_db.rename(columns=COLUMN_MAP_BY_LOCATION)
    # Normalize DB data: convert numeric columns from strings to float
    df_db = normalize_db_data_by_location(df_db, NUMERIC_COLS)
    
    merged = pd.merge(
        df_excel, df_db[JOIN_KEYS_BY_LOCATION + NUMERIC_COLS],
        on=JOIN_KEYS_BY_LOCATION, how='outer', suffixes=('_XLS', '_DB')
    )
    errors = compare_db_to_excel(merged, JOIN_KEYS_BY_LOCATION, NUMERIC_COLS, tolerance)
    if errors:
        print_error_summary(errors, label)
        pytest.fail(f"{len(errors)} error(s) found")
    else:
        print(f"\n[PASS] '{label}' matches database.")


def run_employee_comparison(df_excel, df_db, label, tolerance=0.1):
    """Rename DB columns, normalize data types, outer-merge with Excel employee data, and assert no mismatches."""
    df_db = df_db.rename(columns=COLUMN_MAP_BY_EMPLOYEE)
    # Normalize DB data: convert numeric columns from strings to float, EmpNo to string
    df_db = normalize_db_data_by_employee(df_db, NUMERIC_COLS)

    # Both EmpNo columns are now str — no conversion needed here
    merged = pd.merge(
        df_excel, df_db[JOIN_KEYS_BY_EMPLOYEE + NUMERIC_COLS],
        on=JOIN_KEYS_BY_EMPLOYEE, how='outer', suffixes=('_XLS', '_DB')
    )
    #print(f"\n[DEBUG] Merged DataFrame for '{label}':\n{merged.to_string()}")

    errors = compare_db_to_excel(merged, ["EmpNo", "Name"], NUMERIC_COLS, tolerance)
    if errors:
        print_error_summary(errors, label)
        pytest.fail(f"{len(errors)} error(s) found")
    else:
        print(f"\n[PASS] '{label}' matches database.")


# ==========================================
# SESSION-SCOPED DB FIXTURES
# ==========================================

@pytest.fixture(scope="session")
def db_summary_monthly():
    """Fetches Summary SP result for the report month once and caches for all monthly location tests."""
    return fetch_summary_from_db(MONTHLY_START, MONTHLY_END)

@pytest.fixture(scope="session")
def db_summary_ytd():
    """Fetches Summary SP result for YTD once and caches for all YTD location tests."""
    return fetch_summary_from_db(YTD_START, YTD_END)

@pytest.fixture(scope="session")
def db_detail_monthly():
    """Fetches Detail SP result for the report month once and caches for all monthly employee tests."""
    return fetch_detail_from_db(MONTHLY_START, MONTHLY_END)

@pytest.fixture(scope="session")
def db_detail_ytd():
    """Fetches Detail SP result for YTD once and caches for all YTD employee tests."""
    return fetch_detail_from_db(YTD_START, YTD_END)


# ==========================================
# DB COMPARISON TESTS
# ==========================================

def test_db_comparison_us_monthly(db_summary_monthly):
    """Validates Excel US monthly sheet against DB for the report month."""
    df_excel = normalize_excel_data_by_location(FIXTURE_FILE, f"{REPORT_MONTH}_US", header_row=1)
    df_excel = df_excel[df_excel['Title'].isin(SUBTOTAL_TITLES)].copy()

    df_db = db_summary_monthly[db_summary_monthly['Office_Code'] != 'Europe'].copy()

    run_location_comparison(df_excel, df_db, f"{REPORT_MONTH}_US", tolerance=3.0)


def test_db_comparison_europe_monthly(db_summary_monthly):
    """Validates Europe rows in Excel Total monthly sheet against DB for the report month."""
    df_excel = normalize_excel_data_by_location(FIXTURE_FILE, f"{REPORT_MONTH}_Total", header_row=2)
    df_excel = df_excel[
        (df_excel['Office'] == 'Europe') &
        (df_excel['Title'].isin(SUBTOTAL_TITLES))
    ].copy()

    df_db = db_summary_monthly[db_summary_monthly['Office_Code'] == 'Europe'].copy()

    run_location_comparison(df_excel, df_db, f"{REPORT_MONTH}_Total (Europe)", tolerance=3.0)


def test_db_comparison_us_ytd(db_summary_ytd):
    """Validates Excel US YTD sheet against DB for YTD through the report month."""
    df_excel = normalize_excel_data_by_location(FIXTURE_FILE, f"{REPORT_YEAR}_US", header_row=1)
    df_excel = df_excel[df_excel['Title'].isin(SUBTOTAL_TITLES)].copy()

    df_db = db_summary_ytd[db_summary_ytd['Office_Code'] != 'Europe'].copy()

    run_location_comparison(df_excel, df_db, f"{REPORT_YEAR}_US (YTD)", tolerance=3.0)


def test_db_comparison_europe_ytd(db_summary_ytd):
    """Validates Europe rows in Excel Total YTD sheet against DB for YTD through the report month."""
    df_excel = normalize_excel_data_by_location(FIXTURE_FILE, f"{REPORT_YEAR}_Total", header_row=2)
    df_excel = df_excel[
        (df_excel['Office'] == 'Europe') &
        (df_excel['Title'].isin(SUBTOTAL_TITLES))
    ].copy()

    df_db = db_summary_ytd[db_summary_ytd['Office_Code'] == 'Europe'].copy()

    run_location_comparison(df_excel, df_db, f"{REPORT_YEAR}_Total (Europe YTD)", tolerance=3.0)


@pytest.mark.parametrize("loc", US_EMPLOYEE_LOCATION_CONFIG, ids=[l["name"] for l in US_EMPLOYEE_LOCATION_CONFIG])
def test_db_comparison_employee_monthly(loc, db_detail_monthly):
    """Validates employee data in each office monthly sheet against DB for the report month.
    For Europe offices (EU-UK, EU-BE), DB returns all EU employees combined — filtered
    to this sheet's EmpNos only to avoid false mismatches with the other EU office.
    """
    df_excel = normalize_excel_data_by_employee(FIXTURE_FILE, loc["monthly_sheet"], header_row=1)
    df_excel = df_excel[df_excel['row_type'] == 'data'].copy()

    df_db = db_detail_monthly[db_detail_monthly['Office_Code'] == loc["office_code"]].copy()

    if loc.get("filter_by_excel_empnos"):
        df_db = df_db[df_db['EMPLOYEE_CODE'].isin(df_excel['EmpNo'])].copy()

    run_employee_comparison(df_excel, df_db, f"{loc['monthly_sheet']} (Monthly)", tolerance=3.0)


@pytest.mark.parametrize("loc", US_EMPLOYEE_LOCATION_CONFIG, ids=[l["name"] for l in US_EMPLOYEE_LOCATION_CONFIG])
def test_db_comparison_employee_ytd(loc, db_detail_ytd):
    """Validates employee data in each office YTD sheet against DB for YTD through the report month.
    For Europe offices (EU-UK, EU-BE), DB returns all EU employees combined — filtered
    to this sheet's EmpNos only to avoid false mismatches with the other EU office.
    """
    df_excel = normalize_excel_data_by_employee(FIXTURE_FILE, loc["ytd_sheet"], header_row=1)
    df_excel = df_excel[df_excel['row_type'] == 'data'].copy()
    #print(f"[DEBUG] df_excel:\n{df_excel.to_string()}")

    df_db = db_detail_ytd[db_detail_ytd['Office_Code'] == loc["office_code"]].copy()

    if loc.get("filter_by_excel_empnos"):
        df_db = df_db[df_db['EMPLOYEE_CODE'].isin(df_excel['EmpNo'])].copy()
   
    #print(f"[DEBUG] df_db:\n{df_db.to_string()}")
    run_employee_comparison(df_excel, df_db, f"{loc['ytd_sheet']} (YTD)", tolerance=3.0)

if __name__ == "__main__":
    pytest.main([__file__])