"""
PPLX business logic - Aux Data analysis, SCID extraction, keyword handling.
"""

import re
from typing import List, Optional, Tuple

# Default AUX Data values
DEFAULT_AUX_VALUES = {
    "Aux Data 1": "XCEL",
    "Aux Data 2": "NO TAG",
    "Aux Data 3": "EXISTING",
    "Aux Data 4": "NO MAKE READY",
    "Aux Data 5": "NO",
}


def _normalize_keywords(keywords) -> List[str]:
    """Normalize keyword inputs to uppercase lists."""
    if not keywords:
        return []
    if isinstance(keywords, str):
        keywords = keywords.split(",")
    return [keyword.strip().upper() for keyword in keywords if keyword and keyword.strip()]


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
    comm_keywords = _normalize_keywords(comm_keywords)
    power_keywords = _normalize_keywords(power_keywords)
    pco_keywords = _normalize_keywords(pco_keywords)
    aux5_keywords = _normalize_keywords(aux5_keywords)
    label = power_label.upper()

    if any(keyword in mr_note_upper for keyword in pco_keywords):
        aux_data_4 = "PCO"
    else:
        has_comm = any(keyword in mr_note_upper for keyword in comm_keywords)
        has_power = any(keyword in mr_note_upper for keyword in power_keywords)
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
        if aux5_keywords and any(keyword in mr_note_upper for keyword in aux5_keywords)
        else "NO"
    )
    return aux_data_4, aux_data_5


def determine_aux_data_values(
    scid: str,
    mr_note: str,
    excel_data: Optional[dict] = None,
    comm_keywords: Optional[List[str]] = None,
    power_keywords: Optional[List[str]] = None,
    pco_keywords: Optional[List[str]] = None,
    aux5_keywords: Optional[List[str]] = None,
    power_label: str = "POWER",
) -> dict:
    """Determine Aux Data values based on SCID, mr_note, and Excel data."""
    aux_updates = {}
    row_data = excel_data.get(scid, {}) if excel_data else {}

    aux_updates["Aux Data 1"] = row_data.get("pole_tag_company", DEFAULT_AUX_VALUES["Aux Data 1"])
    pole_tag = row_data.get("pole_tag_tagtext", "")
    if not pole_tag or pole_tag.strip() == "" or str(pole_tag).lower() == "nan":
        aux_updates["Aux Data 2"] = DEFAULT_AUX_VALUES["Aux Data 2"]
    else:
        aux_updates["Aux Data 2"] = str(pole_tag).strip()

    aux_updates["Aux Data 3"] = DEFAULT_AUX_VALUES["Aux Data 3"]
    aux_data_4, aux_data_5 = analyze_mr_note_for_aux_data(
        mr_note,
        comm_keywords=comm_keywords,
        power_keywords=power_keywords,
        pco_keywords=pco_keywords,
        aux5_keywords=aux5_keywords,
        power_label=power_label,
    )
    aux_updates["Aux Data 4"] = aux_data_4
    aux_updates["Aux Data 5"] = aux_data_5
    return aux_updates


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


def normalize_scid_for_excel_lookup(scid: str) -> str:
    """Normalize SCID for Excel lookup by removing periods and spaces."""
    return scid.replace(".", "").replace(" ", "")
