"""Pytest path fix — ensures the xlide package under python/ is importable."""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
