"""Database utility functions."""

import pyodbc
import pandas as pd
from typing import Any, List, Tuple, Optional, Dict, Union
from dotenv import load_dotenv
import os

# Load environment variables from config/creds/personal.env
load_dotenv(dotenv_path=os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))), 'config', 'creds', 'personal.env'))


def get_db_connection_from_env(server: str, database: str, trusted_connection: bool = False):
    """
    Create and return a database connection using the provided server and database.

    :param server:             SQL Server hostname or IP (e.g. "SQLT1COFIN")
    :param database:           Database name (e.g. "OfficersContribution")
    :param trusted_connection: If True, use Windows Authentication (Trusted_Connection=yes).
                               If False, reads DB_USERNAME and DB_PASSWORD from config/creds/personal.env.
    """
    if trusted_connection:
        return connect_to_database(server, database, trusted_connection=True)
    username = os.getenv("DB_USERNAME")
    password = os.getenv("DB_PASSWORD")
    if not all([username, password]):
        raise ValueError("DB_USERNAME and DB_PASSWORD must be set in config/creds/personal.env.")
    return connect_to_database(server, database, username=username, password=password)


def connect_to_database(server: Optional[str], database: Optional[str],
                        username: Optional[str] = None, password: Optional[str] = None,
                        trusted_connection: bool = False):
    """
    Connect to a Microsoft SQL Server database using pyodbc.
    Supports both Windows Authentication (trusted_connection=True) and SQL Server auth.
    Returns a connection object if successful, or raises an exception.
    """
    if not all([server, database]):
        raise ValueError("Server and database must be provided.")
    if trusted_connection:
        connection_string = (
            f"Driver={{ODBC Driver 17 for SQL Server}};"
            f"Server={server};"
            f"Database={database};"
            f"Trusted_Connection=yes;"
        )
    else:
        if not all([username, password]):
            raise ValueError("Username and password must be provided for SQL Server authentication.")
        connection_string = (
            f"Driver={{ODBC Driver 17 for SQL Server}};"
            f"Server={server};"
            f"Database={database};"
            f"UID={username};"
            f"PWD={password};"
        )
    try:
        conn = pyodbc.connect(connection_string)
        print(f"Connected to SQL Server database '{database}' on server '{server}'")
        return conn
    except Exception as e:
        print(f"Failed to connect to database: {e}")
        raise


def select_query(
    conn: pyodbc.Connection,
    query: str,
    params: Optional[Tuple[Any, ...]] = None,
    as_dataframe: bool = False
) -> Union[List[Tuple], Any]:
    """
    Execute a SELECT query and return the results.

    :param conn: pyodbc connection
    :param query: SQL SELECT query string
    :param params: Optional tuple of query parameters
    :param as_dataframe: If True, return results as a pandas DataFrame
    :return: List of tuples by default, or a pandas DataFrame if as_dataframe=True
    """
    cursor = conn.cursor()
    try:
        cursor.execute(query, params or ())
        rows = cursor.fetchall()
        if as_dataframe:
            cols = [col[0] for col in cursor.description]
            return pd.DataFrame([tuple(row) for row in rows], columns=cols)
        return rows
    except Exception as e:
        print(f"Error executing SELECT query: {e}")
        raise
    finally:
        cursor.close()


def insert_query(conn: pyodbc.Connection, query: str, params: Optional[Tuple[Any, ...]] = None) -> int:
    """
    Execute an INSERT query. Returns the number of affected rows.
    """
    try:
        cursor = conn.cursor()
        cursor.execute(query, params or ())
        conn.commit()
        rowcount = cursor.rowcount
        cursor.close()
        return rowcount
    except Exception as e:
        print(f"Error executing INSERT query: {e}")
        conn.rollback()
        raise


def update_query(conn: pyodbc.Connection, query: str, params: Optional[Tuple[Any, ...]] = None) -> int:
    """
    Execute an UPDATE query. Returns the number of affected rows.
    """
    try:
        cursor = conn.cursor()
        cursor.execute(query, params or ())
        conn.commit()
        rowcount = cursor.rowcount
        cursor.close()
        return rowcount
    except Exception as e:
        print(f"Error executing UPDATE query: {e}")
        conn.rollback()
        raise


def delete_query(conn: pyodbc.Connection, query: str, params: Optional[Tuple[Any, ...]] = None) -> int:
    """
    Execute a DELETE query. Returns the number of affected rows.
    """
    try:
        cursor = conn.cursor()
        cursor.execute(query, params or ())
        conn.commit()
        rowcount = cursor.rowcount
        cursor.close()
        return rowcount
    except Exception as e:
        print(f"Error executing DELETE query: {e}")
        conn.rollback()
        raise 


def call_stored_procedure(
    conn: pyodbc.Connection,
    proc_name: str,
    params: Optional[Tuple[Any, ...]] = None,
    named_params: Optional[Dict[str, Any]] = None,
    fetch_as_dict: bool = False,
    as_dataframe: bool = False,
    commit_on_success: bool = False
) -> List[Any]:
    """
    Call a stored procedure and collect one or more result sets it returns.

    Behavior:
      - Supports positional params (params=) or named params (named_params=), not both.
      - Iterates all result sets returned by the procedure (uses cursor.nextset()).
      - Returns a list where each item is one result set:
          - as_dataframe=True  -> each result set is a pandas DataFrame
          - fetch_as_dict=True -> each result set is a list of dicts keyed by column name
          - default            -> each result set is a list of tuples
      - If the procedure returns no tabular result sets, an empty list is returned.
      - Optionally commits the connection if commit_on_success is True.

    Limitations:
      - Does not handle output parameters or return codes.
      - Non-tabular results (PRINT, RAISERROR, etc.) are ignored.

    Examples:
      # Positional params
      result_sets = call_stored_procedure(conn, "dbo.MyProc", params=(123, "abc"))

      # Named params -> builds: EXEC dbo.MyProc @Year=?, @Month=?
      result_sets = call_stored_procedure(conn, "dbo.MyProc",
                                          named_params={"Year": 2026, "Month": 2},
                                          as_dataframe=True)
      df = result_sets[0]  # first result set as DataFrame

    :param conn:             pyodbc connection
    :param proc_name:        Stored procedure name e.g. "dbo.usp_MyProc"
    :param params:           Tuple of positional parameters
    :param named_params:     Dict of named parameters e.g. {"StartDate": "2026-01-01"}
    :param fetch_as_dict:    If True, return rows as dicts keyed by column name
    :param as_dataframe:     If True, return each result set as a pandas DataFrame
    :param commit_on_success: If True, commit after a successful call
    :return: List of result sets (tuples, dicts, or DataFrames depending on flags)
    """
    if params and named_params:
        raise ValueError("Provide either 'params' or 'named_params', not both.")

    cursor = conn.cursor()
    result_sets: List[Any] = []
    try:
        # Build and execute EXEC statement
        if named_params:
            placeholders = ", ".join(f"@{k}=?" for k in named_params)
            cursor.execute(f"EXEC {proc_name} {placeholders}", tuple(named_params.values()))
        elif params:
            placeholders = ", ".join("?" for _ in params)
            cursor.execute(f"EXEC {proc_name} {placeholders}", params)
        else:
            cursor.execute(f"EXEC {proc_name}")

        # Iterate over all result sets
        while True:
            if cursor.description:
                cols = [col[0] for col in cursor.description]
                rows = cursor.fetchall()
                if as_dataframe:
                    result_sets.append(pd.DataFrame([tuple(row) for row in rows], columns=cols))
                elif fetch_as_dict:
                    result_sets.append([dict(zip(cols, row)) for row in rows])
                else:
                    result_sets.append(rows)
            if not cursor.nextset():
                break

        if commit_on_success:
            conn.commit()

        return result_sets

    except Exception as e:
        if commit_on_success:
            try:
                conn.rollback()
            except Exception:
                pass
        print(f"Error calling stored procedure '{proc_name}': {e}")
        raise
    finally:
        try:
            cursor.close()
        except Exception:
            pass