import re
from pathlib import Path


def load_query(file_name: str, query_name: str) -> str:
    """
    Load a named SQL query from a .sql file in the sqls directory.

    Usage:
        sql = load_query("ardent_queries", "employee_billable_hours_by_office")

    Queries are delimited by '-- query: <name>' markers inside the .sql file.
    """
    sql_path = Path(__file__).parent / f"{file_name}.sql"
    content  = sql_path.read_text(encoding="utf-8")

    pattern = rf"--\s*query:\s*{re.escape(query_name)}\s*\n(.*?)(?=--\s*query:|\Z)"
    match   = re.search(pattern, content, re.DOTALL | re.IGNORECASE)

    if not match:
        raise ValueError(f"Query '{query_name}' not found in '{file_name}.sql'")

    return match.group(1).strip()
