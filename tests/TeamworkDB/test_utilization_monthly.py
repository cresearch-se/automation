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
from cornerstone_automation.sqls.loader import load_query

# ==========================================
# CHANGE ONLY THIS LINE EACH MONTH
# ==========================================
FIXTURE_FILE  = "tests/TeamworkDB/fixtures/Utilization_202608.xlsx"

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

# Cross-sheet validation configuration
# Each entry defines a summary sheet office total that must match
# the corresponding detail sheet office total
CONFIG_CROSS_SHEET = {
    "monthly": [
        {
            "summary_sheet":  f"{REPORT_MONTH}_US",
            "summary_office": "CRB",
            "detail_sheet":   "Month - CRB ",
            "label":          "Boston Monthly"
        },
        {
            "summary_sheet":  f"{REPORT_MONTH}_US",
            "summary_office": "CRCH",
            "detail_sheet":   "Month - CRCH",
            "label":          "Chicago Monthly"
        },
        {
            "summary_sheet":  f"{REPORT_MONTH}_US",
            "summary_office": "CRDC",
            "detail_sheet":   "Month - CRDC",
            "label":          "DC Monthly"
        },
        {
            "summary_sheet":  f"{REPORT_MONTH}_US",
            "summary_office": "CRLA",
            "detail_sheet":   "Month - CRLA",
            "label":          "LA Monthly"
        },
        {
            "summary_sheet":  f"{REPORT_MONTH}_US",
            "summary_office": "CRNY",
            "detail_sheet":   "Month - CRNY",
            "label":          "NY Monthly"
        },
        {
            "summary_sheet":  f"{REPORT_MONTH}_US",
            "summary_office": "CRSF",
            "detail_sheet":   "Month - CRSF",
            "label":          "SF Monthly"
        },
        {
            "summary_sheet":  f"{REPORT_MONTH}_US",
            "summary_office": "CRSV",
            "detail_sheet":   "Month - CRSV",
            "label":          "SV Monthly"
        },
        {
            "summary_sheet":  f"{REPORT_MONTH}_Europe",
            "summary_office": "Brussels",
            "detail_sheet":   "Month - CRBE",
            "label":          "Brussels Monthly"
        },
        {
            "summary_sheet":  f"{REPORT_MONTH}_Europe",
            "summary_office": "London",
            "detail_sheet":   "Month - CRUK",
            "label":          "London Monthly"
        },
        {
            "summary_sheet":  f"{REPORT_MONTH}_Total",
            "summary_office": "US",
            "detail_sheet":   f"{REPORT_MONTH}_US",
            "label":          "US Total vs US Sheet Monthly",
            "detail_is_summary": True
        },
        {
            "summary_sheet":  f"{REPORT_MONTH}_Total",
            "summary_office": "Europe",
            "detail_sheet":   f"{REPORT_MONTH}_Europe",
            "label":          "Europe Total vs Europe Sheet Monthly",
            "detail_is_summary": True
        }
    ],
    "ytd": [
        {
            "summary_sheet":  f"{REPORT_YEAR}_US",
            "summary_office": "CRB",
            "detail_sheet":   "YTD - CRB ",
            "label":          "Boston YTD"
        },
        {
            "summary_sheet":  f"{REPORT_YEAR}_US",
            "summary_office": "CRCH",
            "detail_sheet":   "YTD - CRCH",
            "label":          "Chicago YTD"
        },
        {
            "summary_sheet":  f"{REPORT_YEAR}_US",
            "summary_office": "CRDC",
            "detail_sheet":   "YTD - CRDC",
            "label":          "DC YTD"
        },
        {
            "summary_sheet":  f"{REPORT_YEAR}_US",
            "summary_office": "CRLA",
            "detail_sheet":   "YTD - CRLA",
            "label":          "LA YTD"
        },
        {
            "summary_sheet":  f"{REPORT_YEAR}_US",
            "summary_office": "CRNY",
            "detail_sheet":   "YTD - CRNY",
            "label":          "NY YTD"
        },
        {
            "summary_sheet":  f"{REPORT_YEAR}_US",
            "summary_office": "CRSF",
            "detail_sheet":   "YTD - CRSF",
            "label":          "SF YTD"
        },
        {
            "summary_sheet":  f"{REPORT_YEAR}_US",
            "summary_office": "CRSV",
            "detail_sheet":   "YTD - CRSV",
            "label":          "SV YTD"
        },
        {
            "summary_sheet":  f"{REPORT_YEAR}_Europe",
            "summary_office": "Brussels",
            "detail_sheet":   "YTD - CRBE",
            "label":          "Brussels YTD"
        },
        {
            "summary_sheet":  f"{REPORT_YEAR}_Europe",
            "summary_office": "London",
            "detail_sheet":   "YTD - CRUK",
            "label":          "London YTD"
        },
        {
            "summary_sheet":  f"{REPORT_YEAR}_Total",
            "summary_office": "US",
            "detail_sheet":   f"{REPORT_YEAR}_US",
            "label":          "US Total vs US Sheet YTD",
            "detail_is_summary": True
        },
        {
            "summary_sheet":  f"{REPORT_YEAR}_Total",
            "summary_office": "Europe",
            "detail_sheet":   f"{REPORT_YEAR}_Europe",
            "label":          "Europe Total vs Europe Sheet YTD",
            "detail_is_summary": True
        }
    ]
}

OFFICE_MAP = {
    'Boston': 'CRB',
    'Chicago': 'CRCH',
    'Washington, D.C.': 'CRDC',
    'Washington DC': 'CRDC',      # Workday spelling
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

# Workday office name → SP office code mapping
# Kept separate from OFFICE_MAP to avoid affecting Excel parsing
# London and Brussels map to Europe in the SP output
WORKDAY_OFFICE_TO_CODE = {
    'New York':       'CRNY',
    'Chicago':        'CRCH',
    'Washington DC':  'CRDC',
    'Boston':         'CRB',
    'San Francisco':  'CRSF',
    'Los Angeles':    'CRLA',
    'Silicon Valley': 'CRSV',
    'London':         'Europe',
    'Brussels':       'Europe'
}

# Workday Level → Report Title Group mapping
# Source: Master_Users_WD joined to TargetDaily joined to TBL_RANK._TITLE_GROUP
# This is the definitive mapping derived directly from the database
WORKDAY_LEVEL_TO_TITLE = {
    # Officers — _TITLE_GROUP = 'Officer'
    'SRVP':     'Officer',   # Rank 2
    'SRVPUK':   'Officer',   # Rank 2
    'SRVPCOO':  'Officer',   # Rank 2
    'SRVPCTIO': 'Officer',   # Rank 2
    'VP':       'Officer',   # Rank 3
    'VPUK':     'Officer',   # Rank 3
    'VPBE':     'Officer',   # Rank 3
    'CEO':      'Officer',   # Rank 107
    'COB':      'Officer',   # Rank 110
    # Principal — _TITLE_GROUP = 'Principal'
    'P1':    'Principal',    # Rank 101
    'P2':    'Principal',    # Rank 101
    'P3':    'Principal',    # Rank 101
    'P4':    'Principal',    # Rank 101
    'P5':    'Principal',    # Rank 101
    'P6':    'Principal',    # Rank 101
    'P7':    'Principal',    # Rank 101
    'P1UK':  'Principal',    # Rank 101
    'P2UK':  'Principal',    # Rank 101
    'P3UK':  None,           # Rank 10 = Other Admin, NULL _TITLE_GROUP
    'P4UK':  'Principal',    # Rank 101
    'P5UK':  'Principal',    # Rank 101
    'P5BE':  'Principal',    # Rank 101
    'P6UK':  'Principal',    # Rank 101
    # Manager — _TITLE_GROUP = 'Manager'
    'SM1':    'Manager',     # Rank 4
    'SM2':    'Manager',     # Rank 4
    'SM1UK':  'Manager',     # Rank 4
    'SM2UK':  'Manager',     # Rank 4
    'M1':     'Manager',     # Rank 5
    'M2':     'Manager',     # Rank 5
    'M1UK':   'Manager',     # Rank 5
    'M2UK':   'Manager',     # Rank 5
    'M1BE':   'Manager',     # Rank 5
    'SEM1':   'Manager',     # Rank 100
    'SEM2':   'Manager',     # Rank 100
    'SEA2':   'Manager',     # Rank 100
    'SEA3':   'Manager',     # Rank 100
    'SEP2':   'Manager',     # Rank 100
    'SESM1':  'Manager',     # Rank 100
    'SESM2':  'Manager',     # Rank 100
    'SCM2':   'Manager',     # Rank 105
    'SRADVREE':'Manager',    # Rank 120
    'A4':     'Manager',     # Rank 5
    # Associate — _TITLE_GROUP = 'Associate'
    'A1':    'Associate',    # Rank 6
    'A2':    'Associate',    # Rank 6
    'A3':    'Associate',    # Rank 6
    'A1UK':  'Associate',    # Rank 6
    'A2UK':  'Associate',    # Rank 6
    'A3UK':  'Associate',    # Rank 6
    'A3BE':  'Associate',    # Rank 6
    'SRA':   'Associate',    # Rank 6
    # Analyst — _TITLE_GROUP = 'Analyst'
    'RA1':   'Analyst',      # Rank 7
    'RA2':   'Analyst',      # Rank 7
    'RA3':   'Analyst',      # Rank 7
    'RA4':   'Analyst',      # Rank 7
    'RA1UK': 'Analyst',      # Rank 7
    'RA2UK': 'Analyst',      # Rank 7
    'RA3UK': 'Analyst',      # Rank 7
    'RA4UK': 'Analyst',      # Rank 7
    'SA1':   'Analyst',      # Rank 8
    'SA1UK': 'Analyst',      # Rank 8
    'SA1BE': 'Analyst',      # Rank 8
    'AN1':   'Analyst',      # Rank 9
    'AN2':   'Analyst',      # Rank 9
    'AN1BE': 'Analyst',      # Rank 9
    'AN2BE': 'Analyst',      # Rank 9
    'AN2UK': 'Analyst',      # Rank 9
}

# SP Rank_title → broad title group mapping
# Maps SP Detail output Rank_title to our 5 broad groups
SP_TITLE_TO_GROUP = {
    'SVP/VP':             'Officer',
    'Principal':          'Principal',
    'Senior Manager':     'Manager',
    'Manager':            'Manager',
    'SRA/Associate':      'Associate',
    'Research Associate': 'Analyst',
    'Senior Analyst':     'Analyst',
    'Analyst':            'Analyst',
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

    # Ensure EmpNo is clean str on both sides before merge — prevents same EmpNo
    # appearing in both MISSING IN DB and MISSING IN XLS due to whitespace mismatch
    df_excel = df_excel.copy()
    df_excel['EmpNo'] = df_excel['EmpNo'].astype(str).str.strip()
    df_db['EmpNo']    = df_db['EmpNo'].astype(str).str.strip()

    # Exclude DB rows with NULL Target_Hours — compare_db_to_excel uses Target_Hours_DB
    # as the presence indicator, so NULL Target_Hours fires false MISSING IN DB for every
    # employee absent from Excel, producing duplicate MISSING IN DB + MISSING IN XLS.
    df_db = df_db[df_db[NUMERIC_COLS[0]].notna()].copy()

    merged = pd.merge(
        df_excel, df_db[JOIN_KEYS_BY_EMPLOYEE + NUMERIC_COLS],
        on=JOIN_KEYS_BY_EMPLOYEE, how='outer', suffixes=('_XLS', '_DB')
    )

    # DEBUG — print full merged table so we can see why EmpNos appear in both MISSING categories
    # pd.set_option('display.max_rows', 500)
    # pd.set_option('display.max_columns', 20)
    # pd.set_option('display.width', 300)
    # print(f"\n[DEBUG MERGED — {label}]")
    # print(merged.to_string(index=False))

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

def test_terminated_employees_with_hours_appear_in_ytd_report(db_detail_ytd):
    """
    Validates that employees who were terminated during the YTD period
    but had actual billable hours still appear in the SP YTD Detail output.

    Root cause of original bug: TargetDaily stops generating rows at
    termination date, so hours worked before termination have no #TD row
    to join against and are silently dropped from the SP output.

    This test catches that by independently finding terminated employees
    with hours from raw tables (TAT_TIME + HBM_PERSNL) and verifying
    each one is present in the SP output — completely independent of
    both the SP and the Excel report.
    """
    conn = get_db_connection_from_env(DB_SERVER, DB_DATABASE, trusted_connection=True)
    try:
        sql = load_query(
            "terminated_employees_list",
            "terminated_employees_list"
        )
        terminated_rows = conn.execute(
            sql,
            (
                f"{REPORT_YEAR}-01-01",  # terminate_date >= year start
                YTD_END,                 # terminate_date <= report end
                f"{REPORT_YEAR}-01-01",  # hours from year start
                YTD_END                  # hours through report end
            )
        ).fetchall()
    finally:
        conn.close()

    if not terminated_rows:
        print("\n[PASS] No terminated employees with hours found "
              "in YTD period.")
        return

    # Build set of EmpNos present in SP Detail output
    sp_empnos = set(
        str(e).split('.')[0].strip()
        for e in db_detail_ytd['EMPLOYEE_CODE'].dropna()
    )

    errors   = []
    warnings = []

    for row in terminated_rows:
        emp_code         = str(row[0]).strip()
        emp_name         = str(row[1])
        termination_date = str(row[2])[:10]
        office           = str(row[3])
        total_hrs        = round(float(row[4]), 2)

        if emp_code not in sp_empnos:
            errors.append(
                f"[MISSING TERMINATED EMPLOYEE] "
                f"EmpNo={emp_code} | Name={emp_name} | "
                f"Office={office} | "
                f"TerminationDate={termination_date} | "
                f"YTD_Hrs={total_hrs}"
            )
        else:
            warnings.append(
                f"[OK] EmpNo={emp_code} | Name={emp_name} | "
                f"Office={office} | "
                f"TerminationDate={termination_date} | "
                f"YTD_Hrs={total_hrs}"
            )

    # Print passing employees for visibility
    if warnings:
        print(f"\n  Terminated employees correctly present in SP "
              f"output ({len(warnings)}):")
        for w in warnings:
            print(f"    {w}")

    if errors:
        print(f"\n  TERMINATED EMPLOYEES MISSING FROM YTD REPORT "
              f"({len(errors)}):")
        print(f"  These employees were terminated during {REPORT_YEAR} "
              f"but had actual hours and should appear in YTD data.\n")
        for e in errors:
            print(f"    {e}")
        pytest.fail(
            f"{len(errors)} terminated employee(s) with actual hours "
            f"are missing from the YTD SP output. Check TargetDaily "
            f"population for these employees."
        )
    else:
        print(
            f"\n[PASS] All {len(terminated_rows)} terminated employee(s) "
            f"with YTD hours are present in the SP output."
        )

# ==========================================
# ROSTER COMPLETENESS TESTS
# ==========================================

def _run_roster_check(roster_rows, sp_detail_df, label):
    """
    Core roster validation logic shared by monthly and YTD tests.

    For each employee in the Workday roster checks:
      1. Existence — is the employee in the SP output at all?
      2. Office    — are they under the correct office?
      3. Title     — are they under the correct title group?

    Returns tuple of (errors, warnings) lists.
    """

    # DEBUG — check columns and find 5905
    print(f"\n[DEBUG] SP columns: {list(sp_detail_df.columns)}")
    matches = sp_detail_df[
        sp_detail_df.apply(
            lambda row: '5905' in str(row.values), axis=1
        )
    ]
    print(f"\n[DEBUG] Rows containing 5905:\n{matches.to_string()}")

    # Build lookup: EmpNo → {office, title} from SP output

    sp_lookup = {}
    for _, row in sp_detail_df.iterrows():
        emp = str(row.get('EMPLOYEE_CODE', '')).split('.')[0].strip()
        if emp and emp != 'nan':
            sp_lookup[emp] = {
                'office': str(row.get('Office_Code', '')).strip(),
                'title':  str(row.get('Rank_title', '')).strip()
            }

    errors   = []
    warnings = []

    for row in roster_rows:
        emp_code   = str(row[0]).strip()
        emp_name   = str(row[1])
        wd_office  = str(row[2])
        dept_code  = str(row[3])
        status     = str(row[6])
        hire_date  = str(row[7])[:10]
        term_date  = str(row[8])[:10] if row[8] else 'Active'
        wd_level   = str(row[9]).strip() if row[9] else ''
        wd_title   = str(row[10]).strip() if row[10] else ''

        # Map Workday office and level to expected SP values
        expected_office = WORKDAY_OFFICE_TO_CODE.get(wd_office, wd_office)
        expected_title  = WORKDAY_LEVEL_TO_TITLE.get(wd_level, None)

        # --- Check 1: Existence ---
        if emp_code not in sp_lookup:
            errors.append(
                f"[MISSING FROM {label}] "
                f"EmpNo={emp_code} | Name={emp_name} | "
                f"Office={wd_office} | Dept={dept_code} | "
                f"Status={status} | "
                f"HireDate={hire_date} | "
                f"TermDate={term_date} | "
                f"Level={wd_level} | "
                f"JobTitle={wd_title}"
            )
            continue  # No point checking office/title if missing

        actual_office = sp_lookup[emp_code]['office']
        actual_title  = sp_lookup[emp_code]['title']

        # --- Check 2: Office ---
        if actual_office != expected_office:
            errors.append(
                f"[WRONG OFFICE] "
                f"EmpNo={emp_code} | Name={emp_name} | "
                f"Expected_Office={expected_office} | "
                f"Actual_Office={actual_office} | "
                f"Workday_Office={wd_office}"
            )

        # --- Check 3: Title ---
        if expected_title:
            actual_title_group = SP_TITLE_TO_GROUP.get(
                actual_title, None
            )
            if actual_title_group and actual_title_group != expected_title:
                errors.append(
                    f"[WRONG TITLE] "
                    f"EmpNo={emp_code} | Name={emp_name} | "
                    f"Expected_Title={expected_title} | "
                    f"Actual_SP_Title={actual_title} | "
                    f"Actual_Group={actual_title_group} | "
                    f"Workday_Level={wd_level} | "
                    f"Workday_JobTitle={wd_title}"
                )
            elif not actual_title_group:
                print(
                    f"    [WARN] Unknown SP title '{actual_title}' "
                    f"for EmpNo={emp_code} | Name={emp_name}"
                )

        # Track passing employees — only if no error logged for them
        has_error = any(f"EmpNo={emp_code}" in e for e in errors)
        if not has_error:
            warnings.append(
                f"[OK] EmpNo={emp_code} | Name={emp_name} | "
                f"Office={actual_office} | Title={actual_title}"
            )

    return errors, warnings


def test_roster_completeness_monthly(db_detail_monthly):
    """
    Validates that all regular permanent employees who should be
    in the monthly report are present in the SP monthly Detail output.

    Checks three things per employee:
      1. Existence — present in SP output
      2. Office    — under correct office
      3. Title     — under correct title group

    Uses Master_Users_WD (Workday) as independent source of truth.
    Monthly scope: employees active during the report month (~222).

    Catches employees missing from BOTH Excel and SP which existing
    comparison tests cannot detect.
    """
    conn = get_db_connection_from_env(
        DB_SERVER, DB_DATABASE, trusted_connection=True
    )
    try:
        sql = load_query("roster_completeness", "active_employees_in_scope")
        roster_rows = conn.execute(
            sql,
            (
                MONTHLY_END,    # HireDate <= report month end
                MONTHLY_START   # TerminationDate >= month start or NULL
            )
        ).fetchall()
    finally:
        conn.close()

    if not roster_rows:
        print("\n[PASS] No in-scope employees found in Workday roster.")
        return

    errors, warnings = _run_roster_check(
        roster_rows, db_detail_monthly, "MONTHLY REPORT"
    )

    if warnings:
        print(
            f"\n  Employees correctly present in monthly SP "
            f"output ({len(warnings)}/{len(roster_rows)})"
        )

    if errors:
        missing      = [e for e in errors if '[MISSING'      in e]
        wrong_office = [e for e in errors if '[WRONG OFFICE]' in e]
        wrong_title  = [e for e in errors if '[WRONG TITLE]'  in e]

        if missing:
            print(f"\n  MISSING FROM MONTHLY REPORT ({len(missing)}):")
            for e in missing:
                print(f"    {e}")

        if wrong_office:
            print(f"\n  WRONG OFFICE IN MONTHLY REPORT "
                  f"({len(wrong_office)}):")
            for e in wrong_office:
                print(f"    {e}")

        if wrong_title:
            print(f"\n  WRONG TITLE IN MONTHLY REPORT "
                  f"({len(wrong_title)}):")
            for e in wrong_title:
                print(f"    {e}")

        pytest.fail(
            f"{len(errors)} issue(s) found in monthly roster: "
            f"{len(missing)} missing, "
            f"{len(wrong_office)} wrong office, "
            f"{len(wrong_title)} wrong title."
        )
    else:
        print(
            f"\n[PASS] All {len(roster_rows)} in-scope employee(s) "
            f"from Workday roster are correctly present in monthly "
            f"SP output with correct office and title."
        )


def test_roster_completeness_ytd(db_detail_ytd):
    """
    Validates that all regular permanent employees who should be
    in the YTD report are present in the SP YTD Detail output.

    Checks three things per employee:
      1. Existence — present in SP output
      2. Office    — under correct office
      3. Title     — under correct title group

    Uses Master_Users_WD (Workday) as independent source of truth.
    YTD scope: employees active at any point Jan 1 → report end (~228).

    Catches employees missing from BOTH Excel and SP which existing
    comparison tests cannot detect — e.g. employees never populated
    in TargetDaily or terminated employees whose TargetDaily rows
    stopped before all their hours were recorded.
    """
    conn = get_db_connection_from_env(
        DB_SERVER, DB_DATABASE, trusted_connection=True
    )
    try:
        sql = load_query("roster_completeness", "active_employees_in_scope")
        roster_rows = conn.execute(
            sql,
            (
                YTD_END,    # HireDate <= YTD end
                YTD_START   # TerminationDate >= Jan 1 or NULL
            )
        ).fetchall()
    finally:
        conn.close()

    if not roster_rows:
        print("\n[PASS] No in-scope employees found in Workday roster.")
        return

    errors, warnings = _run_roster_check(
        roster_rows, db_detail_ytd, "YTD REPORT"
    )

    if warnings:
        print(
            f"\n  Employees correctly present in YTD SP "
            f"output ({len(warnings)}/{len(roster_rows)})"
        )

    if errors:
        missing      = [e for e in errors if '[MISSING'       in e]
        wrong_office = [e for e in errors if '[WRONG OFFICE]' in e]
        wrong_title  = [e for e in errors if '[WRONG TITLE]'  in e]

        if missing:
            print(f"\n  MISSING FROM YTD REPORT ({len(missing)}):")
            for e in missing:
                print(f"    {e}")

        if wrong_office:
            print(f"\n  WRONG OFFICE IN YTD REPORT "
                  f"({len(wrong_office)}):")
            for e in wrong_office:
                print(f"    {e}")

        if wrong_title:
            print(f"\n  WRONG TITLE IN YTD REPORT "
                  f"({len(wrong_title)}):")
            for e in wrong_title:
                print(f"    {e}")

        pytest.fail(
            f"{len(errors)} issue(s) found in YTD roster: "
            f"{len(missing)} missing, "
            f"{len(wrong_office)} wrong office, "
            f"{len(wrong_title)} wrong title."
        )
    else:
        print(
            f"\n[PASS] All {len(roster_rows)} in-scope employee(s) "
            f"from Workday roster are correctly present in YTD "
            f"SP output with correct office and title."
        )


# ==========================================
# CROSS-SHEET CONSISTENCY TESTS
# ==========================================

def _get_office_total_from_summary(file_path, sheet_name, office_name, header_row=1):
    """
    Reads a summary sheet and returns the OFFICE TOTAL row
    for a specific office as a dict of numeric values.

    For US/Europe sheets — finds the OFFICE TOTAL row for that office.
    For Total sheet — finds the US TOTAL or Europe TOTAL row.
    For cross-sheet US/Europe check — finds the grand total of that sheet.
    """
    # Determine subtotal and grand total labels based on sheet
    if '_Total' in sheet_name:
        subtotal_label = ["US TOTAL", "Europe TOTAL"]
        grand_total_label = "GRAND TOTAL"
    elif '_Europe' in sheet_name:
        subtotal_label = "OFFICE TOTAL"
        grand_total_label = "Europe TOTAL"
    else:
        subtotal_label = "OFFICE TOTAL"
        grand_total_label = "US TOTAL"

    df = normalize_excel_data_with_totals(
        file_path, sheet_name,
        subtotal_label=subtotal_label,
        grand_total_label=grand_total_label,
        header_row=header_row
    )

    # Case 1 — looking for a specific office subtotal
    # e.g. Boston in 202608_US, Brussels in 202608_Europe
    df_office = df[
        (df['Office'] == office_name) &
        (df['row_type'] == 'subtotal')
    ]
    if not df_office.empty:
        return {col: df_office[col].sum() for col in NUMERIC_COLS}

    # Case 2 — looking for US TOTAL or Europe TOTAL in Total sheet
    # e.g. office_name = 'US' in 202608_Total
    df_office = df[
        (df['Office'] == office_name) &
        (df['row_type'] == 'subtotal')
    ]
    if not df_office.empty:
        return {col: df_office[col].sum() for col in NUMERIC_COLS}

    # Case 3 — looking for grand total of a sheet
    # e.g. when detail_is_summary=True comparing Total vs US sheet
    df_grand = df[df['row_type'] == 'grand_total']
    if not df_grand.empty:
        return {col: df_grand[col].sum() for col in NUMERIC_COLS}

    return None

def _get_office_total_from_detail(file_path, sheet_name):
    """
    Reads a detail sheet (e.g. Month - CRB) and returns
    the OFFICE TOTAL row as a dict of numeric values.
    """
    # Try OFFICE TOTAL first
    try:
        df = normalize_excel_data_by_employee(
            file_path, sheet_name,
            grand_total_label="OFFICE TOTAL"
        )
        df_grand = df[df['row_type'] == 'grand_total']
        if not df_grand.empty:
            return {col: df_grand[col].sum() for col in NUMERIC_COLS}
    except Exception:
        pass

    # Try TOTAL for Data Science / Applied Research sheets
    try:
        df = normalize_excel_data_by_employee(
            file_path, sheet_name,
            grand_total_label="TOTAL"
        )
        df_grand = df[df['row_type'] == 'grand_total']
        if not df_grand.empty:
            return {col: df_grand[col].sum() for col in NUMERIC_COLS}
    except Exception:
        pass

    return None


def _run_cross_sheet_check(comparisons, period_label, header_row=1):
    """
    Core cross-sheet validation logic.

    For each comparison:
      - Gets the office total from the summary sheet
      - Gets the office total from the detail sheet
      - Compares all four numeric columns
      - Flags any mismatches

    Returns list of error strings.
    """
    errors = []
    tolerance = 1.0  # Allow $1 rounding difference

    for comp in comparisons:
        summary_sheet  = comp['summary_sheet']
        summary_office = comp['summary_office']
        detail_sheet   = comp['detail_sheet']
        label          = comp['label']
        is_summary     = comp.get('detail_is_summary', False)

        # Get summary totals
        # Total sheet has header_row=2
        s_header = 2 if 'Total' in summary_sheet else 1
        summary_totals = _get_office_total_from_summary(
            FIXTURE_FILE, summary_sheet,
            summary_office, header_row=s_header
        )

        if summary_totals is None:
            errors.append(
                f"[CROSS-SHEET] [{label}] "
                f"Could not find office '{summary_office}' "
                f"in sheet '{summary_sheet}'"
            )
            continue

        # Get detail totals
        if is_summary:
            # Detail sheet is also a summary sheet
            # We want the grand total of that sheet
            # e.g. US TOTAL from 202608_US
            # or Europe TOTAL from 202608_Europe
            d_header = 2 if 'Total' in detail_sheet else 1
            if '_Europe' in detail_sheet:
                grand_label = "Europe TOTAL"
                sub_label = "OFFICE TOTAL"
            else:
                grand_label = "US TOTAL"
                sub_label = "OFFICE TOTAL"

            df_detail = normalize_excel_data_with_totals(
                FIXTURE_FILE, detail_sheet,
                subtotal_label=sub_label,
                grand_total_label=grand_label,
                header_row=d_header
            )
            df_grand = df_detail[df_detail['row_type'] == 'grand_total']
            if df_grand.empty:
                detail_totals = None
            else:
                detail_totals = {
                    col: df_grand[col].sum()
                    for col in NUMERIC_COLS
                }
        else:
            detail_totals = _get_office_total_from_detail(
                FIXTURE_FILE, detail_sheet
            )

        if detail_totals is None:
            errors.append(
                f"[CROSS-SHEET] [{label}] "
                f"Could not find total in sheet '{detail_sheet}'"
            )
            continue

        # Compare each numeric column
        for col in NUMERIC_COLS:
            summary_val = summary_totals.get(col, 0) or 0
            detail_val  = detail_totals.get(col, 0) or 0
            diff        = abs(summary_val - detail_val)

            if diff > tolerance:
                pct_diff = (diff / summary_val * 100) if summary_val != 0 else 0
                errors.append(
                    f"[CROSS-SHEET MISMATCH] [{label}] "
                    f"Column={col} | "
                    f"Summary({summary_sheet})={summary_val:,.2f} | "
                    f"Detail({detail_sheet})={detail_val:,.2f} | "
                    f"Diff={diff:,.2f} | "
                    f"Diff%={pct_diff:.2f}%"
                )
        
        if not any(label in e for e in errors):
            print(f"    [OK] {label} — all columns match")

    return errors


def test_cross_sheet_consistency_monthly():
    """
    Validates that summary sheet office totals match
    corresponding detail sheet office totals for the report month.

    Checks:
      - Each US office total in 202608_US matches its Month - XX sheet
      - Each Europe office total in 202608_Europe matches its Month - XX sheet
      - US TOTAL in 202608_Total matches 202608_US US TOTAL
      - Europe TOTAL in 202608_Total matches 202608_Europe Europe TOTAL

    Catches cases where a detail sheet was updated but the summary
    sheet was not regenerated — or vice versa.
    """
    assert os.path.exists(FIXTURE_FILE), f"File not found: {FIXTURE_FILE}"

    errors = _run_cross_sheet_check(
        CONFIG_CROSS_SHEET["monthly"],
        f"Monthly {REPORT_MONTH}"
    )

    if errors:
        print(f"\n  CROSS-SHEET MISMATCHES FOUND ({len(errors)}):")
        for e in errors:
            print(f"    {e}")
        pytest.fail(
            f"{len(errors)} cross-sheet mismatch(es) found "
            f"in monthly report."
        )
    else:
        print(
            f"\n[PASS] All monthly summary sheet totals match "
            f"their corresponding detail sheets."
        )


def test_cross_sheet_consistency_ytd():
    """
    Validates that summary sheet office totals match
    corresponding detail sheet office totals for YTD.

    Checks:
      - Each US office total in 2026_US matches its YTD - XX sheet
      - Each Europe office total in 2026_Europe matches its YTD - XX sheet
      - US TOTAL in 2026_Total matches 2026_US US TOTAL
      - Europe TOTAL in 2026_Total matches 2026_Europe Europe TOTAL

    Catches cases where a detail sheet was updated but the summary
    sheet was not regenerated — or vice versa.
    """
    assert os.path.exists(FIXTURE_FILE), f"File not found: {FIXTURE_FILE}"

    errors = _run_cross_sheet_check(
        CONFIG_CROSS_SHEET["ytd"],
        f"YTD {REPORT_YEAR}"
    )

    if errors:
        print(f"\n  CROSS-SHEET MISMATCHES FOUND ({len(errors)}):")
        for e in errors:
            print(f"    {e}")
        pytest.fail(
            f"{len(errors)} cross-sheet mismatch(es) found "
            f"in YTD report."
        )
    else:
        print(
            f"\n[PASS] All YTD summary sheet totals match "
            f"their corresponding detail sheets."
        )

def test_cornerstone_research_equals_us_plus_europe_monthly():
    """
    Validates that each Cornerstone Research title row in the monthly
    Total sheet equals the sum of the corresponding US and Europe rows.

    For each title (Officer, Principal, Manager, Associate, Analyst):
        Cornerstone Research = US + Europe

    Catches cases where the Total sheet is not properly aggregating
    US and Europe data.
    """
    assert os.path.exists(FIXTURE_FILE), f"File not found: {FIXTURE_FILE}"

    df = normalize_excel_data_by_location(
        FIXTURE_FILE, f"{REPORT_MONTH}_Total", header_row=2
    )

    errors = []
    tolerance = 1.0

    # Get data rows only for each section
    df_us     = df[df['Office'] == 'US'].copy()
    df_europe = df[df['Office'] == 'Europe'].copy()
    df_cr     = df[df['Office'] == 'Cornerstone Research'].copy()

    if df_us.empty or df_europe.empty or df_cr.empty:
        errors.append(
            "[MISSING SECTION] Could not find US, Europe or "
            "Cornerstone Research section in Total sheet"
        )
        pytest.fail("\n".join(errors))
        return

    # For each title in Cornerstone Research verify = US + Europe
    for _, cr_row in df_cr.iterrows():
        title = cr_row['Title'].strip()

        us_row     = df_us[df_us['Title'].str.strip() == title]
        europe_row = df_europe[df_europe['Title'].str.strip() == title]

        if us_row.empty or europe_row.empty:
            # Title may not exist in both regions — skip
            continue

        for col in NUMERIC_COLS:
            cr_val     = cr_row[col] or 0
            us_val     = us_row[col].values[0] or 0
            europe_val = europe_row[col].values[0] or 0
            expected   = us_val + europe_val
            diff       = abs(cr_val - expected)

            if diff > tolerance:
                errors.append(
                    f"[CR MISMATCH] Title={title} | Col={col} | "
                    f"CR={cr_val:,.2f} | "
                    f"US={us_val:,.2f} + Europe={europe_val:,.2f} "
                    f"= {expected:,.2f} | Diff={diff:,.2f}"
                )

    if errors:
        print(f"\n  CORNERSTONE RESEARCH MISMATCH — MONTHLY ({len(errors)}):")
        for e in errors:
            print(f"    {e}")
        pytest.fail(
            f"{len(errors)} Cornerstone Research row(s) don't equal "
            f"US + Europe in monthly Total sheet."
        )
    else:
        print(
            f"\n[PASS] All Cornerstone Research rows in monthly Total "
            f"sheet equal US + Europe correctly."
        )


def test_cornerstone_research_equals_us_plus_europe_ytd():
    """
    Validates that each Cornerstone Research title row in the YTD
    Total sheet equals the sum of the corresponding US and Europe rows.

    For each title (Officer, Principal, Manager, Associate, Analyst):
        Cornerstone Research = US + Europe

    Catches cases where the Total sheet is not properly aggregating
    US and Europe data for the YTD period.
    """
    assert os.path.exists(FIXTURE_FILE), f"File not found: {FIXTURE_FILE}"

    df = normalize_excel_data_by_location(
        FIXTURE_FILE, f"{REPORT_YEAR}_Total", header_row=2
    )

    errors = []
    tolerance = 1.0

    df_us     = df[df['Office'] == 'US'].copy()
    df_europe = df[df['Office'] == 'Europe'].copy()
    df_cr     = df[df['Office'] == 'Cornerstone Research'].copy()

    if df_us.empty or df_europe.empty or df_cr.empty:
        errors.append(
            "[MISSING SECTION] Could not find US, Europe or "
            "Cornerstone Research section in YTD Total sheet"
        )
        pytest.fail("\n".join(errors))
        return

    for _, cr_row in df_cr.iterrows():
        title = cr_row['Title'].strip()

        us_row     = df_us[df_us['Title'].str.strip() == title]
        europe_row = df_europe[df_europe['Title'].str.strip() == title]

        if us_row.empty or europe_row.empty:
            continue

        for col in NUMERIC_COLS:
            cr_val     = cr_row[col] or 0
            us_val     = us_row[col].values[0] or 0
            europe_val = europe_row[col].values[0] or 0
            expected   = us_val + europe_val
            diff       = abs(cr_val - expected)

            if diff > tolerance:
                errors.append(
                    f"[CR MISMATCH] Title={title} | Col={col} | "
                    f"CR={cr_val:,.2f} | "
                    f"US={us_val:,.2f} + Europe={europe_val:,.2f} "
                    f"= {expected:,.2f} | Diff={diff:,.2f}"
                )

    if errors:
        print(f"\n  CORNERSTONE RESEARCH MISMATCH — YTD ({len(errors)}):")
        for e in errors:
            print(f"    {e}")
        pytest.fail(
            f"{len(errors)} Cornerstone Research row(s) don't equal "
            f"US + Europe in YTD Total sheet."
        )
    else:
        print(
            f"\n[PASS] All Cornerstone Research rows in YTD Total "
            f"sheet equal US + Europe correctly."
        )

def test_associate_analyst_rollup_by_location():
    """
    Validates that Associate and Analyst totals in each summary sheet
    equal the sum of their Exp and 1st Yr components.

    For each office in each summary sheet:
        Associate = Associate-Exp + Associate-1st Yr
        Analyst   = Analyst-Exp  + Analyst-1st Yr

    Catches cases where the rollup formula in Excel is broken or
    the data rows don't add up to the summary row.
    """
    assert os.path.exists(FIXTURE_FILE), f"File not found: {FIXTURE_FILE}"

    errors  = []
    tolerance = 1.0

    # Check all summary sheets
    sheets_to_check = [
        {"sheet": f"{REPORT_MONTH}_US",     "header_row": 1,
         "offices": ['CRB','CRCH','CRDC','CRLA','CRNY','CRSF','CRSV']},
        {"sheet": f"{REPORT_MONTH}_Europe", "header_row": 1,
         "offices": ['Brussels','London']},
        {"sheet": f"{REPORT_YEAR}_US",      "header_row": 1,
         "offices": ['CRB','CRCH','CRDC','CRLA','CRNY','CRSF','CRSV']},
        {"sheet": f"{REPORT_YEAR}_Europe",  "header_row": 1,
         "offices": ['Brussels','London']},
    ]

    for sheet_config in sheets_to_check:
        sheet      = sheet_config["sheet"]
        header_row = sheet_config["header_row"]
        offices    = sheet_config["offices"]

        df = normalize_excel_data_by_location(
            FIXTURE_FILE, sheet, header_row=header_row
        )

        for office in offices:
            df_office = df[df['Office'] == office]

            if df_office.empty:
                continue

            # Check Associate rollup
            row_exp    = df_office[df_office['Title'] == 'Associate-Exp']
            row_1yr    = df_office[df_office['Title'] == 'Associate-1st Yr']
            row_total  = df_office[df_office['Title'] == 'Associate']

            if not row_exp.empty and not row_1yr.empty and not row_total.empty:
                for col in NUMERIC_COLS:
                    exp_val   = row_exp[col].values[0] or 0
                    yr_val    = row_1yr[col].values[0] or 0
                    total_val = row_total[col].values[0] or 0
                    expected  = exp_val + yr_val
                    diff      = abs(total_val - expected)

                    if diff > tolerance:
                        errors.append(
                            f"[ASSOCIATE ROLLUP MISMATCH] "
                            f"Sheet={sheet} | Office={office} | "
                            f"Col={col} | "
                            f"Associate={total_val:,.2f} | "
                            f"Exp={exp_val:,.2f} + "
                            f"1stYr={yr_val:,.2f} = {expected:,.2f} | "
                            f"Diff={diff:,.2f}"
                        )

            # Check Analyst rollup
            row_exp   = df_office[df_office['Title'] == 'Analyst-Exp']
            row_1yr   = df_office[df_office['Title'] == 'Analyst-1st Yr']
            row_total = df_office[df_office['Title'] == 'Analyst']

            if not row_exp.empty and not row_1yr.empty and not row_total.empty:
                for col in NUMERIC_COLS:
                    exp_val   = row_exp[col].values[0] or 0
                    yr_val    = row_1yr[col].values[0] or 0
                    total_val = row_total[col].values[0] or 0
                    expected  = exp_val + yr_val
                    diff      = abs(total_val - expected)

                    if diff > tolerance:
                        errors.append(
                            f"[ANALYST ROLLUP MISMATCH] "
                            f"Sheet={sheet} | Office={office} | "
                            f"Col={col} | "
                            f"Analyst={total_val:,.2f} | "
                            f"Exp={exp_val:,.2f} + "
                            f"1stYr={yr_val:,.2f} = {expected:,.2f} | "
                            f"Diff={diff:,.2f}"
                        )

    if errors:
        print(f"\n  ROLLUP MISMATCHES ({len(errors)}):")
        for e in errors:
            print(f"    {e}")
        pytest.fail(
            f"{len(errors)} Associate/Analyst rollup mismatch(es) found."
        )
    else:
        print(
            f"\n[PASS] All Associate and Analyst rollup rows correctly "
            f"equal Exp + 1st Yr across all summary sheets."
        )

def test_no_employee_in_multiple_office_sheets_monthly():
    """
    Validates that no employee appears in more than one office
    monthly detail sheet.

    An employee should belong to exactly one office. If they appear
    in multiple sheets it indicates a data classification error or
    an office transfer that wasn't handled correctly.

    Note: Europe offices (CRUK and CRBE) are excluded from this check
    since both feed into the same 'Europe' office code in the SP —
    employees could legitimately appear in only one of them.
    """
    assert os.path.exists(FIXTURE_FILE), f"File not found: {FIXTURE_FILE}"

    errors = []

    # Only US offices — Europe excluded as explained above
    us_sheets = [
        {"sheet": "Month - CRB ",  "office": "Boston"},
        {"sheet": "Month - CRCH",  "office": "Chicago"},
        {"sheet": "Month - CRDC",  "office": "DC"},
        {"sheet": "Month - CRLA",  "office": "LA"},
        {"sheet": "Month - CRNY",  "office": "NY"},
        {"sheet": "Month - CRSF",  "office": "SF"},
        {"sheet": "Month - CRSV",  "office": "SV"},
    ]

    # Build map: EmpNo → list of offices they appear in
    emp_office_map = {}

    for sheet_config in us_sheets:
        sheet  = sheet_config["sheet"]
        office = sheet_config["office"]

        df = normalize_excel_data_by_employee(
            FIXTURE_FILE, sheet,
            grand_total_label="OFFICE TOTAL"
        )
        df_data = df[df['row_type'] == 'data'].copy()

        for _, row in df_data.iterrows():
            emp_no = str(row['EmpNo']).strip()
            if emp_no and emp_no != 'nan':
                if emp_no not in emp_office_map:
                    emp_office_map[emp_no] = []
                emp_office_map[emp_no].append(
                    {"office": office, "name": str(row['Name'])}
                )

    # Flag any employee in more than one office
    for emp_no, offices in emp_office_map.items():
        if len(offices) > 1:
            office_list = ", ".join(
                f"{o['office']} ({o['name']})" for o in offices
            )
            errors.append(
                f"[EMPLOYEE IN MULTIPLE OFFICES] "
                f"EmpNo={emp_no} appears in {len(offices)} offices: "
                f"{office_list}"
            )

    if errors:
        print(f"\n  EMPLOYEES IN MULTIPLE OFFICES — MONTHLY ({len(errors)}):")
        for e in errors:
            print(f"    {e}")
        pytest.fail(
            f"{len(errors)} employee(s) appear in multiple "
            f"monthly office sheets."
        )
    else:
        print(
            f"\n[PASS] No employees appear in multiple monthly "
            f"office sheets."
        )


def test_no_employee_in_multiple_office_sheets_ytd():
    """
    Validates that no employee appears in more than one office
    YTD detail sheet.

    Same logic as monthly check but for YTD sheets.
    """
    assert os.path.exists(FIXTURE_FILE), f"File not found: {FIXTURE_FILE}"

    errors = []

    us_sheets = [
        {"sheet": "YTD - CRB ",  "office": "Boston"},
        {"sheet": "YTD - CRCH",  "office": "Chicago"},
        {"sheet": "YTD - CRDC",  "office": "DC"},
        {"sheet": "YTD - CRLA",  "office": "LA"},
        {"sheet": "YTD - CRNY",  "office": "NY"},
        {"sheet": "YTD - CRSF",  "office": "SF"},
        {"sheet": "YTD - CRSV",  "office": "SV"},
    ]

    emp_office_map = {}

    for sheet_config in us_sheets:
        sheet  = sheet_config["sheet"]
        office = sheet_config["office"]

        df = normalize_excel_data_by_employee(
            FIXTURE_FILE, sheet,
            grand_total_label="OFFICE TOTAL"
        )
        df_data = df[df['row_type'] == 'data'].copy()

        for _, row in df_data.iterrows():
            emp_no = str(row['EmpNo']).strip()
            if emp_no and emp_no != 'nan':
                if emp_no not in emp_office_map:
                    emp_office_map[emp_no] = []
                emp_office_map[emp_no].append(
                    {"office": office, "name": str(row['Name'])}
                )

    for emp_no, offices in emp_office_map.items():
        if len(offices) > 1:
            office_list = ", ".join(
                f"{o['office']} ({o['name']})" for o in offices
            )
            errors.append(
                f"[EMPLOYEE IN MULTIPLE OFFICES] "
                f"EmpNo={emp_no} appears in {len(offices)} offices: "
                f"{office_list}"
            )

    if errors:
        print(f"\n  EMPLOYEES IN MULTIPLE OFFICES — YTD ({len(errors)}):")
        for e in errors:
            print(f"    {e}")
        pytest.fail(
            f"{len(errors)} employee(s) appear in multiple "
            f"YTD office sheets."
        )
    else:
        print(
            f"\n[PASS] No employees appear in multiple YTD "
            f"office sheets."
        )

def test_ytd_hours_greater_than_or_equal_to_monthly():
    """
    Validates that each employee's YTD hours and revenue are greater
    than or equal to their monthly hours and revenue.

    Since YTD covers Jan through report month, YTD should always be
    >= monthly for any given employee. If YTD < Monthly it indicates:
      - Wrong date range used for one of the queries
      - Data was removed from YTD but not monthly
      - Report was generated with mismatched periods

    Checks all US and Europe office sheets.
    """
    assert os.path.exists(FIXTURE_FILE), f"File not found: {FIXTURE_FILE}"

    errors   = []
    tolerance = 0.1  # Allow tiny rounding differences

    office_configs = [
        {"monthly": "Month - CRB ",  "ytd": "YTD - CRB ",  "office": "Boston"},
        {"monthly": "Month - CRCH",  "ytd": "YTD - CRCH",  "office": "Chicago"},
        {"monthly": "Month - CRDC",  "ytd": "YTD - CRDC",  "office": "DC"},
        {"monthly": "Month - CRLA",  "ytd": "YTD - CRLA",  "office": "LA"},
        {"monthly": "Month - CRNY",  "ytd": "YTD - CRNY",  "office": "NY"},
        {"monthly": "Month - CRSF",  "ytd": "YTD - CRSF",  "office": "SF"},
        {"monthly": "Month - CRSV",  "ytd": "YTD - CRSV",  "office": "SV"},
        {"monthly": "Month - CRBE",  "ytd": "YTD - CRBE",  "office": "Brussels"},
        {"monthly": "Month - CRUK",  "ytd": "YTD - CRUK",  "office": "London"},
    ]

    for config in office_configs:
        monthly_sheet = config["monthly"]
        ytd_sheet     = config["ytd"]
        office        = config["office"]

        # Load monthly data
        df_monthly = normalize_excel_data_by_employee(
            FIXTURE_FILE, monthly_sheet,
            grand_total_label="OFFICE TOTAL"
        )
        df_monthly = df_monthly[
            df_monthly['row_type'] == 'data'
        ].copy()

        # Load YTD data
        df_ytd = normalize_excel_data_by_employee(
            FIXTURE_FILE, ytd_sheet,
            grand_total_label="OFFICE TOTAL"
        )
        df_ytd = df_ytd[df_ytd['row_type'] == 'data'].copy()

        # Build YTD lookup by EmpNo
        ytd_lookup = {}
        for _, row in df_ytd.iterrows():
            emp_no = str(row['EmpNo']).strip()
            if emp_no and emp_no != 'nan':
                ytd_lookup[emp_no] = row

        # Compare monthly vs YTD for each employee
        for _, m_row in df_monthly.iterrows():
            emp_no   = str(m_row['EmpNo']).strip()
            emp_name = str(m_row['Name'])

            if emp_no not in ytd_lookup:
                # Employee in monthly but not YTD — caught by other tests
                continue

            y_row = ytd_lookup[emp_no]

            for col in NUMERIC_COLS:
                monthly_val = m_row[col] or 0
                ytd_val     = y_row[col] or 0

                if ytd_val < monthly_val - tolerance:
                    errors.append(
                        f"[YTD LESS THAN MONTHLY] "
                        f"Office={office} | "
                        f"EmpNo={emp_no} | Name={emp_name} | "
                        f"Col={col} | "
                        f"Monthly={monthly_val:,.2f} | "
                        f"YTD={ytd_val:,.2f} | "
                        f"Diff={monthly_val - ytd_val:,.2f}"
                    )

    if errors:
        print(f"\n  YTD LESS THAN MONTHLY ({len(errors)}):")
        for e in errors:
            print(f"    {e}")
        pytest.fail(
            f"{len(errors)} employee(s) have YTD values less than "
            f"their monthly values."
        )
    else:
        print(
            f"\n[PASS] All employees have YTD values greater than or "
            f"equal to their monthly values."
        )
        
if __name__ == "__main__":
    pytest.main([__file__])
