"""Pandas utility functions for data manipulation and analysis."""

import pandas as pd
import numpy as np
from typing import List, Dict, Optional, Any, Union, Tuple
from pathlib import Path


def safe_to_numeric(series: pd.Series, remove_commas: bool = True) -> pd.Series:
    """
    Safely convert a pandas Series to numeric, handling strings with commas.
    
    :param series: The Series to convert.
    :param remove_commas: If True, remove commas before conversion (e.g., '1,234.5' -> 1234.5).
    :return: Series with numeric values (floats). Non-convertible values become NaN.
    """
    if remove_commas:
        series = series.astype(str).str.replace(',', '', regex=False)
    return pd.to_numeric(series, errors='coerce')


# ============================================================================
# READING EXCEL FILES
# ============================================================================

def read_excel_file(
    file_path: str,
    sheet_name: Optional[Union[str, int]] = 0,
    header: int = 0,
    dtype: Optional[Dict[str, type]] = None,
    na_values: Optional[List[str]] = None
) -> pd.DataFrame:
    """
    Read an Excel file and return it as a pandas DataFrame.

    Validations:
      - Verifies the workbook file exists.
      - If `sheet_name` is provided, validates the sheet exists:
        * if sheet_name is str -> must be one of the workbook's sheet names
        * if sheet_name is int -> must be a valid sheet index (0-based)

    :param file_path: Path to the Excel file (.xlsx, .xls)
    :param sheet_name: Sheet name or index to read. Defaults to 0 (first sheet).
    :param header: Row number to use as column names. Defaults to 0.
    :param dtype: Dictionary specifying column data types.
    :param na_values: List of values to recognize as NaN.
    :return: DataFrame containing the Excel data.
    """
    p = Path(file_path)
    if not p.exists():
        raise FileNotFoundError(f"Excel file not found: {file_path}")

    # If a sheet_name was explicitly provided, validate it before reading for clearer errors
    if sheet_name is not None:
        # import here to avoid top-level dependency ordering issues (function defined below)
        sheets = get_excel_sheet_names(file_path)
        if isinstance(sheet_name, str):
            if sheet_name not in sheets:
                raise ValueError(f"Sheet '{sheet_name}' not found in workbook. Available sheets: {sheets}")
        else:
            # ensure int-like and in range
            try:
                idx = int(sheet_name)
            except Exception:
                raise ValueError(f"sheet_name must be a sheet name (str) or index (int); got: {sheet_name!r}")
            if idx < 0 or idx >= len(sheets):
                raise ValueError(f"Sheet index {idx} out of range (0..{max(0, len(sheets)-1)}). Available sheets: {sheets}")

    try:
        df = pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            header=header,
            dtype=dtype,
            na_values=na_values
        )
        return df
    except FileNotFoundError:
        # raise a consistent message (should be unreachable due to earlier check, but keep for completeness)
        raise FileNotFoundError(f"Excel file not found: {file_path}")
    except Exception as e:
        raise Exception(f"Error reading Excel file: {e}")


def read_multiple_sheets(
    file_path: str,
    sheet_names: Optional[List[Union[str, int]]] = None
) -> Dict[str, pd.DataFrame]:
    """
    Read multiple sheets from an Excel file.

    :param file_path: Path to the Excel file.
    :param sheet_names: List of sheet names/indices to read. If None, reads all sheets.
    :return: Dictionary with sheet names as keys and DataFrames as values.
    """
    try:
        if sheet_names is None:
            dfs = pd.read_excel(file_path, sheet_name=None)
        else:
            dfs = {name: pd.read_excel(file_path, sheet_name=name) for name in sheet_names}
        return dfs
    except Exception as e:
        raise Exception(f"Error reading multiple sheets: {e}")


def get_excel_sheet_names(file_path: str) -> List[str]:
    """
    Get all sheet names from an Excel file.

    :param file_path: Path to the Excel file.
    :return: List of sheet names.
    """
    try:
        xls = pd.ExcelFile(file_path)
        return xls.sheet_names
    except Exception as e:
        raise Exception(f"Error reading sheet names: {e}")


# ============================================================================
# DATA FILTERING
# ============================================================================

def filter_by_column(
    df: pd.DataFrame,
    column: str,
    value: Any,
    operator: str = "=="
) -> pd.DataFrame:
    """
    Filter DataFrame rows by column value using specified operator.

    :param df: Input DataFrame.
    :param column: Column name to filter by.
    :param value: Value to compare against.
    :param operator: Comparison operator ("==", "!=", ">", "<", ">=", "<=", "in", "not in")
    :return: Filtered DataFrame.
    """
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found in DataFrame")

    if operator == "==":
        return df[df[column] == value]
    elif operator == "!=":
        return df[df[column] != value]
    elif operator == ">":
        return df[df[column] > value]
    elif operator == "<":
        return df[df[column] < value]
    elif operator == ">=":
        return df[df[column] >= value]
    elif operator == "<=":
        return df[df[column] <= value]
    elif operator == "in":
        return df[df[column].isin(value if isinstance(value, list) else [value])]
    elif operator == "not in":
        return df[~df[column].isin(value if isinstance(value, list) else [value])]
    else:
        raise ValueError(f"Unsupported operator: {operator}")


def filter_by_multiple_conditions(
    df: pd.DataFrame,
    conditions: Dict[str, Tuple[str, Any]]
) -> pd.DataFrame:
    """
    Filter DataFrame by multiple conditions.

    :param df: Input DataFrame.
    :param conditions: Dictionary with column names as keys and tuples of (operator, value).
                      Example: {"Age": (">", 25), "Status": ("==", "Active")}
    :return: Filtered DataFrame.
    """
    result = df.copy()
    for column, (operator, value) in conditions.items():
        result = filter_by_column(result, column, value, operator)
    return result


def filter_by_string_contains(
    df: pd.DataFrame,
    column: str,
    pattern: str,
    case_sensitive: bool = False
) -> pd.DataFrame:
    """
    Filter DataFrame rows where column contains a string pattern.

    :param df: Input DataFrame.
    :param column: Column name to filter by.
    :param pattern: String pattern to search for.
    :param case_sensitive: Whether the search is case-sensitive.
    :return: Filtered DataFrame.
    """
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found in DataFrame")
    return df[df[column].str.contains(pattern, case=case_sensitive, na=False)]


def filter_non_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remove rows that are completely empty or contain only NaN values.

    :param df: Input DataFrame.
    :return: DataFrame with non-empty rows only.
    """
    return df.dropna(how='all')


def filter_rows_with_missing_values(
    df: pd.DataFrame,
    columns: Optional[List[str]] = None,
    keep_na: bool = False
) -> pd.DataFrame:
    """
    Filter rows based on missing values.

    :param df: Input DataFrame.
    :param columns: Specific columns to check. If None, checks all columns.
    :param keep_na: If True, keeps rows with NaN. If False, removes rows with NaN.
    :return: Filtered DataFrame.
    """
    if columns:
        if keep_na:
            return df[df[columns].isna().any(axis=1)]
        else:
            return df[df[columns].notna().all(axis=1)]
    else:
        if keep_na:
            return df[df.isna().any(axis=1)]
        else:
            return df.dropna()


# ============================================================================
# AGGREGATION & CALCULATIONS
# ============================================================================

def calculate_sum(
    df: pd.DataFrame,
    column: str,
    group_by: Optional[List[str]] = None
) -> Union[float, pd.Series]:
    """
    Calculate the sum of a numeric column.

    :param df: Input DataFrame.
    :param column: Column name to sum.
    :param group_by: Optional list of columns to group by.
    :return: Sum value or Series if grouped.
    """
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found in DataFrame")

    if group_by:
        for col in group_by:
            if col not in df.columns:
                raise ValueError(f"Column '{col}' not found in DataFrame")
        return df.groupby(group_by)[column].sum()
    return df[column].sum()


def sum_rows_for_value(
    df: pd.DataFrame,
    match_column: str,
    match_value: Any,
    sum_columns: Optional[Union[str, List[str]]] = None,
    ignore_case: bool = False,
    return_filtered_rows: bool = False
) -> Union[float, pd.Series, Tuple[Union[float, pd.Series], pd.DataFrame]]:
    """
    Find all rows where `match_column` matches `match_value` and return the sum(s).

    Reuses filter_by_column for exact matches and calculate_sum for single-column sums
    when possible to avoid duplicating logic present elsewhere in this module.
    """
    if match_column not in df.columns:
        raise ValueError(f"Match column '{match_column}' not found in DataFrame")

    # Use existing filter_by_column for exact (case-sensitive) matches
    if not ignore_case and pd.api.types.is_string_dtype(df[match_column].dtype):
        filtered = filter_by_column(df, match_column, match_value, operator="==")
    else:
        # case-insensitive or non-string dtype -> build mask manually
        if ignore_case and pd.api.types.is_string_dtype(df[match_column].dtype):
            mask = df[match_column].astype(str).str.strip().str.lower() == str(match_value).strip().lower()
        else:
            mask = df[match_column] == match_value
        filtered = df.loc[mask]

    # No matches -> return zeros consistent with requested shape
    if filtered.empty:
        if isinstance(sum_columns, str):
            result = 0.0
        else:
            if sum_columns is None:
                numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            else:
                cols = sum_columns if isinstance(sum_columns, list) else list(sum_columns)
                for c in cols:
                    if c not in df.columns:
                        raise ValueError(f"Column '{c}' not found in DataFrame")
                numeric_cols = cols
            result = pd.Series({c: 0.0 for c in numeric_cols})
        if return_filtered_rows:
            return result, filtered
        return result

    # determine target columns
    if sum_columns is None:
        target_cols = filtered.select_dtypes(include=[np.number]).columns.tolist()
    elif isinstance(sum_columns, str):
        if sum_columns not in df.columns:
            raise ValueError(f"Column '{sum_columns}' not found in DataFrame")
        target_cols = [sum_columns]
    else:
        missing = [c for c in sum_columns if c not in df.columns]
        if missing:
            raise ValueError(f"Columns not found: {missing}")
        target_cols = list(sum_columns)

    # If single column requested and calculate_sum exists, reuse it
    if isinstance(sum_columns, str):
        scalar = calculate_sum(filtered, sum_columns)
        if return_filtered_rows:
            return float(scalar), filtered
        return float(scalar)

    # multiple or all numeric columns -> use pandas sum
    sums = filtered[target_cols].sum()
    if return_filtered_rows:
        return sums, filtered
    return sums


def calculate_average(
    df: pd.DataFrame,
    column: str,
    group_by: Optional[List[str]] = None
) -> Union[float, pd.Series]:
    """
    Calculate the average of a numeric column.

    :param df: Input DataFrame.
    :param column: Column name to average.
    :param group_by: Optional list of columns to group by.
    :return: Average value or Series if grouped.
    """
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found in DataFrame")

    if group_by:
        for col in group_by:
            if col not in df.columns:
                raise ValueError(f"Column '{col}' not found in DataFrame")
        return df.groupby(group_by)[column].mean()
    return df[column].mean()


def calculate_count(
    df: pd.DataFrame,
    column: Optional[str] = None,
    group_by: Optional[List[str]] = None
) -> Union[int, pd.Series]:
    """
    Count non-null values in a column or rows.

    :param df: Input DataFrame.
    :param column: Column name to count. If None, counts all rows.
    :param group_by: Optional list of columns to group by.
    :return: Count value or Series if grouped.
    """
    if column and column not in df.columns:
        raise ValueError(f"Column '{column}' not found in DataFrame")

    if group_by:
        for col in group_by:
            if col not in df.columns:
                raise ValueError(f"Column '{col}' not found in DataFrame")
        if column:
            return df.groupby(group_by)[column].count()
        else:
            return df.groupby(group_by).size()
    
    if column:
        return df[column].count()
    return len(df)


def calculate_min_max(

    df: pd.DataFrame,
    column: str,
    group_by: Optional[List[str]] = None
) -> Dict[str, Union[float, pd.Series]]:
    """
    Calculate minimum and maximum values of a column.

    :param df: Input DataFrame.
    :param column: Column name.
    :param group_by: Optional list of columns to group by.
    :return: Dictionary with 'min' and 'max' keys.
    """
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found in DataFrame")

    if group_by:
        for col in group_by:
            if col not in df.columns:
                raise ValueError(f"Column '{col}' not found in DataFrame")
        return {
            "min": df.groupby(group_by)[column].min(),
            "max": df.groupby(group_by)[column].max()
        }
    return {
        "min": df[column].min(),
        "max": df[column].max()
    }


def aggregate_data(
    df: pd.DataFrame,
    group_by: List[str],
    aggregations: Dict[str, Union[str, List[str]]]
) -> pd.DataFrame:
    """
    Perform multiple aggregations on grouped data.

    :param df: Input DataFrame.
    :param group_by: List of columns to group by.
    :param aggregations: Dictionary with column names as keys and aggregation functions as values.
                        Example: {"Amount": ["sum", "mean"], "Count": "count"}
    :return: Aggregated DataFrame.
    """
    for col in group_by:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found in DataFrame")
    
    return df.groupby(group_by).agg(aggregations)


# ============================================================================
# DATA MANIPULATION
# ============================================================================

def add_calculated_column(

    df: pd.DataFrame,
    column_name: str,
    formula: callable
) -> pd.DataFrame:
    """
    Add a new calculated column to the DataFrame.

    :param df: Input DataFrame.
    :param column_name: Name of the new column.
    :param formula: Function that takes a row and returns a value.
    :return: DataFrame with new calculated column.
    """
    df[column_name] = df.apply(formula, axis=1)
    return df


def rename_columns(df: pd.DataFrame, mapping: Dict[str, str]) -> pd.DataFrame:
    """
    Rename columns in the DataFrame.

    :param df: Input DataFrame.
    :param mapping: Dictionary with old names as keys and new names as values.
    :return: DataFrame with renamed columns.
    """
    return df.rename(columns=mapping)


def drop_columns(df: pd.DataFrame, columns: List[str]) -> pd.DataFrame:
    """
    Drop specified columns from the DataFrame.

    :param df: Input DataFrame.
    :param columns: List of column names to drop.
    :return: DataFrame with dropped columns.
    """
    existing_cols = [col for col in columns if col in df.columns]
    return df.drop(columns=existing_cols)


def select_columns(df: pd.DataFrame, columns: List[str]) -> pd.DataFrame:
    """
    Select specific columns from the DataFrame.

    :param df: Input DataFrame.
    :param columns: List of column names to select.
    :return: DataFrame with selected columns only.
    """
    missing_cols = [col for col in columns if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Columns not found: {missing_cols}")
    return df[columns]


def sort_dataframe(
    df: pd.DataFrame,
    by: List[str],
    ascending: Union[bool, List[bool]] = True
) -> pd.DataFrame:
    """
    Sort the DataFrame by one or more columns.

    :param df: Input DataFrame.
    :param by: List of column names to sort by.
    :param ascending: Sort order (True for ascending, False for descending).
    :return: Sorted DataFrame.
    """
    for col in by:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found in DataFrame")
    return df.sort_values(by=by, ascending=ascending)


def remove_duplicates(
    df: pd.DataFrame,
    subset: Optional[List[str]] = None,
    keep: str = "first"
) -> pd.DataFrame:
    """
    Remove duplicate rows from the DataFrame.

    :param subset: List of columns to consider for identifying duplicates. If None, uses all columns.
    :param keep: Which duplicates to keep ('first', 'last', or False for none).
    :return: DataFrame with duplicates removed.
    """
    if subset:
        for col in subset:
            if col not in df.columns:
                raise ValueError(f"Column '{col}' not found in DataFrame")
    return df.drop_duplicates(subset=subset, keep=keep)


def fill_missing_values(
    df: pd.DataFrame,
    value: Any = 0,
    method: Optional[str] = None,
    columns: Optional[List[str]] = None
) -> pd.DataFrame:
    """
    Fill missing values in the DataFrame.

    :param df: Input DataFrame.
    :param value: Value to fill with (used if method is None).
    :param method: Interpolation method ('ffill' for forward fill, 'bfill' for backward fill).
    :param columns: Specific columns to fill. If None, fills all columns.
    :return: DataFrame with filled values.
    """
    if method:
        if columns:
            return df[columns].fillna(method=method)
        return df.fillna(method=method)
    else:
        if columns:
            df[columns] = df[columns].fillna(value)
            return df
        return df.fillna(value)


def convert_column_type(
    df: pd.DataFrame,
    column: str,
    dtype: Union[str, type]
) -> pd.DataFrame:
    """
    Convert a column to a specific data type.

    :param df: Input DataFrame.
    :param column: Column name to convert.
    :param dtype: Target data type (e.g., 'int64', 'float', 'str', 'datetime64').
    :return: DataFrame with converted column.
    """
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found in DataFrame")
    
    try:
        df[column] = df[column].astype(dtype)
    except Exception as e:
        raise Exception(f"Error converting column '{column}' to {dtype}: {e}")
    
    return df


# ============================================================================
# DATA EXPORT
# ============================================================================

def write_to_excel(
    df: pd.DataFrame,
    file_path: str,
    sheet_name: str = "Sheet1",
    index: bool = False,
    header: bool = True
) -> None:
    """
    Write DataFrame to an Excel file.

    :param df: DataFrame to write.
    :param file_path: Output Excel file path.
    :param sheet_name: Name of the sheet.
    :param index: Whether to write row indices.
    :param header: Whether to write column headers.
    """
    try:
        df.to_excel(file_path, sheet_name=sheet_name, index=index, header=header)
    except Exception as e:
        raise Exception(f"Error writing to Excel file: {e}")


def write_to_csv(

    df: pd.DataFrame,
    file_path: str,
    index: bool = False,
    header: bool = True
) -> None:
    """
    Write DataFrame to a CSV file.

    :param df: DataFrame to write.
    :param file_path: Output CSV file path.
    :param index: Whether to write row indices.
    :param header: Whether to write column headers.
    """
    try:
        df.to_csv(file_path, index=index, header=header)
    except Exception as e:
        raise Exception(f"Error writing to CSV file: {e}")


# ============================================================================
# DATA ANALYSIS
# ============================================================================

def get_dataframe_info(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Get summary information about the DataFrame.

    :param df: Input DataFrame.
    :return: Dictionary containing DataFrame information.
    """
    return {
        "rows": len(df),
        "columns": len(df.columns),
        "column_names": list(df.columns),
        "dtypes": df.dtypes.to_dict(),
        "memory_usage": df.memory_usage(deep=True).sum(),
        "null_counts": df.isnull().sum().to_dict()
    }


def find_column_by_keywords(df: pd.DataFrame, keywords: List[str]) -> Optional[str]:
    """
    Find the first column whose name contains all given keywords (case-insensitive).

    :param df: Input DataFrame.
    :param keywords: List of substrings that must all appear in the column name.
    :return: Matching column name, or None if not found.
    """
    for col in df.columns:
        col_lower = str(col).strip().lower()
        if all(k.lower() in col_lower for k in keywords):
            return col
    return None


def check_totals_match(
    df_data: pd.DataFrame,
    df_total_row: pd.DataFrame,
    numeric_cols: List[str],
    label: str,
    tolerance: float = 0.01
) -> List[str]:
    """
    Compare the sum of df_data against a single total row for each numeric column.

    :param df_data: DataFrame of detail rows to sum.
    :param df_total_row: Single-row DataFrame containing the expected total values.
    :param numeric_cols: List of numeric column names to check.
    :param label: Descriptive label used in mismatch messages (e.g. office or sheet name).
    :param tolerance: Allowed absolute difference before flagging a mismatch.
    :return: List of mismatch error messages (empty if all match).
    """
    errors = []
    for col in numeric_cols:
        calculated = df_data[col].sum()
        expected = df_total_row[col].iloc[0]
        if abs(calculated - expected) > tolerance:
            errors.append(
                f"[TOTAL MISMATCH] {label} | {col}: "
                f"calculated={calculated:,.2f} vs sheet={expected:,.2f} "
                f"(diff={abs(calculated - expected):,.2f})"
            )
    return errors


def compare_db_to_excel(
    merged: pd.DataFrame,
    key_cols: List[str],
    numeric_cols: List[str],
    tolerance: float = 0.1
) -> List[str]:
    """
    Inspect an outer-merged DataFrame (Excel vs DB) and return a list of error strings.

    Expects the merged DataFrame to have been produced with suffixes=('_XLS', '_DB').
    Checks for:
      - Rows present in Excel but missing in DB   → [MISSING IN DB]
      - Rows present in DB but missing in Excel   → [MISSING IN XLS]
      - Rows present in both but values differ    → [VALUE MISMATCH]

    :param merged:       Outer-merged DataFrame with _XLS / _DB suffixed numeric columns.
    :param key_cols:     Columns used as join keys (e.g. ['Office', 'Title'] or ['EmpNo']).
                         Used to identify rows in error messages.
    :param numeric_cols: List of base numeric column names (without suffixes).
    :param tolerance:    Allowed absolute difference before flagging a value mismatch.
    :return: List of error message strings (empty if all match).
    """
    errors = []
    first_num = f"{numeric_cols[0]}_DB"

    def row_label(row):
        return " | ".join(f"{k}={row[k]}" for k in key_cols)

    # Missing in DB
    for _, row in merged[merged[first_num].isna()].iterrows():
        errors.append(f"[MISSING IN DB] {row_label(row)}")

    # Missing in Excel
    first_xls = f"{numeric_cols[0]}_XLS"
    for _, row in merged[merged[first_xls].isna()].iterrows():
        errors.append(f"[MISSING IN XLS] {row_label(row)}")

    # Value mismatches on matched rows
    matched = merged.dropna(subset=[first_xls, first_num])
    for col in numeric_cols:
        xls_col, db_col = f"{col}_XLS", f"{col}_DB"
        # DB columns should already be float (normalized upfront), but convert just in case
        matched = matched.copy()
        matched[db_col] = pd.to_numeric(matched[db_col], errors='coerce')
        # Use percentage-based tolerance: diff > (tolerance% of XLS value)
        mismatches = matched[abs(matched[xls_col] - matched[db_col]) > (tolerance / 100 * abs(matched[xls_col]))]
        for _, row in mismatches.iterrows():
            pct_diff = (abs(row[xls_col] - row[db_col]) / abs(row[xls_col]) * 100) if row[xls_col] != 0 else float('inf')
            errors.append(
                f"[VALUE MISMATCH] {row_label(row)} | {col}: "
                f"XLS={row[xls_col]:,.2f} vs DB={row[db_col]:,.2f} "
                f"(diff={abs(row[xls_col] - row[db_col]):,.2f}, {pct_diff:.2f}%)"
            )

    return errors


def get_duplicate_count(df: pd.DataFrame, subset: Optional[List[str]] = None) -> int:
    """
    Get the count of duplicate rows.

    :param df: Input DataFrame.
    :param subset: List of columns to consider for duplicates. If None, uses all columns.
    :return: Number of duplicate rows.
    """
    return int(df.duplicated(subset=subset).sum())