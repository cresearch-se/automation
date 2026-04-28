#!/usr/bin/env python3
"""Example usage of the cornerstone-automation package."""

import sys
import os

# Add the src directory to the Python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from cornerstone_automation.main import hello_world


def main() -> None:
    """Demonstrate the package functionality."""
    print("=== Cornerstone Automation Example ===\n")
    
    # Example 1: Default greeting
    print("1. Default greeting:")
    print(hello_world())
    print()
    
    # Example 2: Custom name
    print("2. Custom name greeting:")
    print(hello_world("Python Developer"))
    print()
    
    # Example 3: Multiple greetings
    print("3. Multiple greetings:")
    names = ["Alice", "Bob", "Charlie"]
    for name in names:
        print(hello_world(name))
    print()
    
    print("=== End of Example ===")


if __name__ == "__main__":
    main() 