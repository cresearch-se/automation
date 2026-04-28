"""JSON utility functions."""

import json
from typing import Any

def read_json(file_path: str) -> Any:
    """Read a JSON file and return its contents."""
    with open(file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data 