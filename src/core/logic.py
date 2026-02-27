"""
PPLX business logic - Aux Data analysis, SCID extraction, keyword handling.
"""

import re
from typing import List, Optional, Tuple

from src.core.utils import parse_keywords

# Default fallback when no pole tag found
POLE_TAG_BLANK = "NO TAG"

# Default AUX Data values
DEFAULT_AUX_VALUES = {
    "Aux Data 1": "XCEL",
    "Aux Data 2": "NO TAG",
    "Aux Data 3": "EXISTING",
    "Aux Data 4": "NO MAKE READY",
    "Aux Data 5": "NO",
}


def analyze_mr_note_for_aux_data(
    mr_note: str,
    comm_keywords: Optional[List[str]] = None,
    power_keywords: Optional[List[str]] = None,
    pco_keywords: Optional[List[str]] = None,
    aux5_keywords: Optional[List[str]] = None,
    power_label: str = "POWER",
) -> Tuple[str, str]:
    """
    Analyze mr_note to determine Aux Data 4 and 5 values.
    Returns (aux_data_4, aux_data_5) - both in ALL CAPS format.
    power_label replaces "POWER" in output (e.g. "OPPD" -> "OPPD MAKE READY").
    """
    if not mr_note or mr_note.strip() == "":
        return "NO MAKE READY", "NO"

    mr_note_upper = mr_note.upper()
    comm_kw = parse_keywords(comm_keywords, uppercase=True)
    power_kw = parse_keywords(power_keywords, uppercase=True)
    pco_kw = parse_keywords(pco_keywords, uppercase=True)
    aux5_kw = parse_keywords(aux5_keywords, uppercase=True)
    label = power_label.upper()

    if any(keyword in mr_note_upper for keyword in pco_kw):
        aux_data_4 = "PCO"
    else:
        has_comm = any(keyword in mr_note_upper for keyword in comm_kw)
        has_power = any(keyword in mr_note_upper for keyword in power_kw)
        if has_comm and has_power:
            aux_data_4 = f"{label} & COMM MAKE READY"
        elif has_comm:
            aux_data_4 = "COMM MAKE READY"
        elif has_power:
            aux_data_4 = f"{label} MAKE READY"
        else:
            aux_data_4 = "NO MAKE READY"

    aux_data_5 = (
        "YES"
        if aux5_kw and any(keyword in mr_note_upper for keyword in aux5_kw)
        else "NO"
    )
    return aux_data_4, aux_data_5


def extract_scid_from_filename(filename: str) -> str:
    """Extract SCID from PPLX filename. Example: '001_Ocalc.pplx' -> '001'."""
    if "_Ocalc.pplx" in filename:
        return filename.split("_Ocalc.pplx")[0]
    elif "_" in filename and filename.endswith(".pplx"):
        return filename.split("_")[0]
    return filename.replace(".pplx", "")


def clean_scid_keywords(scid: str, ignore_keywords: str = "") -> str:
    """Remove ignore keywords from SCID."""
    keywords = [kw.strip() for kw in ignore_keywords.split(",") if kw.strip()]
    cleaned_scid = scid
    for keyword in keywords:
        cleaned_scid = re.sub(
            re.escape(keyword), "", cleaned_scid, flags=re.IGNORECASE
        ).strip()
    return " ".join(cleaned_scid.split())
