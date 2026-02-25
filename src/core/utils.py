"""Shared utility functions for PPLX processing."""

from typing import List


def safe_filename_part(s: str, extra_chars: str = "") -> str:
    """Strip to alphanumeric and -_ (and optional extra chars) for filename."""
    allowed = "-_" + extra_chars
    return "".join(c for c in str(s) if c.isalnum() or c in allowed)


def parse_keywords(csv_string: str) -> List[str]:
    """Parse comma-separated string into stripped non-empty keywords."""
    return [kw.strip() for kw in (csv_string or "").split(",") if kw.strip()]
