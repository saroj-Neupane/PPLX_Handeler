"""Shared utility functions for PPLX processing."""

import re
from typing import List


def safe_filename_part(s: str, extra_chars: str = "") -> str:
    """Strip to alphanumeric and -_ (and optional extra chars) for filename."""
    allowed = "-_" + extra_chars
    return "".join(c for c in str(s) if c.isalnum() or c in allowed)


def parse_keywords(csv_string, *, uppercase: bool = False) -> List[str]:
    """Parse comma-separated string (or list) into stripped non-empty keywords.

    Accepts either a csv string or a list of strings.
    When *uppercase* is True, each keyword is upper-cased (useful for matching).
    """
    if not csv_string:
        return []
    if isinstance(csv_string, str):
        items = csv_string.split(",")
    else:
        items = csv_string
    result = [kw.strip() for kw in items if kw and str(kw).strip()]
    if uppercase:
        result = [kw.upper() for kw in result]
    return result


def leading_int(val) -> int:
    """Extract leading integer from a string for numeric sorting."""
    m = re.match(r"(\d+)", str(val or ""))
    return int(m.group(1)) if m else float("inf")
