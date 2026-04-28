"""Main module for the Python project."""

import argparse
import sys
from typing import Optional


def hello_world(name: Optional[str] = None) -> str:
    """
    Return a greeting message.
    
    Args:
        name: Optional name to include in the greeting.
        
    Returns:
        A greeting message string.
    """
    if name is None:
        name = "World"
    return f"Hello, {name}!"


def main() -> int:
    """Main entry point for the application."""
    parser = argparse.ArgumentParser(description="Cornerstone Automation")
    parser.add_argument(
        "--name", 
        type=str, 
        help="Name to greet (default: World)"
    )
    
    args = parser.parse_args()
    
    try:
        message = hello_world(args.name)
        print(message)
        return 0
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main()) 