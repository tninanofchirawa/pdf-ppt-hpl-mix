"""
Text Processing & Sanitization Utilities
========================================
Functions for cleaning statistical values, normalizing player names, and fuzzy scoring.
"""

from difflib import SequenceMatcher
import pandas as pd


def clean_val(val, default: str = "-") -> str:
    """
    Sanitizes values, converting nan, None, or empty values to a clean default.
    Strips trailing '.0' from float integers.
    """
    if pd.isna(val) or val is None or str(val).strip().lower() in ["nan", "none", ""]:
        return default
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def clean_player_name(name: str) -> str:
    """
    Normalizes player names for accurate matching (removes symbols and role tags).
    """
    if not isinstance(name, str):
        return ""
    clean = "".join(c for c in name.upper() if c.isalnum() or c.isspace())
    return clean.replace("WK", "").strip()


def name_similarity(name1: str, name2: str) -> float:
    """
    Computes fuzzy matching similarity ratio between two player names.
    """
    return SequenceMatcher(None, clean_player_name(name1), clean_player_name(name2)).ratio()
