"""Shared fixtures and path setup for the test suite."""

import os
import sys

# Allow tests to import from src/ without installing the package.
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
