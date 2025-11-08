#!/usr/bin/env python3
"""
PPLX Configuration Module - Shared constants and logic for PPLX processing

This module consolidates all the shared configuration, keywords, and business logic
used across the PPLX processing applications to eliminate redundancy.
"""

# =============================================================================
# DEFAULT VALUES
# =============================================================================

# Default AUX Data values
DEFAULT_AUX_VALUES = {
    'Aux Data 1': 'XCEL',           # Default pole owner
    'Aux Data 2': 'NO TAG',         # Default pole tag
    'Aux Data 3': 'EXISTING',       # Default condition
    'Aux Data 4': 'NO MAKE READY',  # Default make ready type
    'Aux Data 5': 'NO'              # Default proposed riser
}

# Default owner configurations
# =============================================================================
# AUX DATA ANALYSIS LOGIC
# =============================================================================

def _normalize_keywords(keywords):
    """
    Normalize keyword inputs to uppercase lists.
    """
    if not keywords:
        return []
    if isinstance(keywords, str):
        keywords = keywords.split(',')
    return [keyword.strip().upper() for keyword in keywords if keyword and keyword.strip()]


def analyze_mr_note_for_aux_data(
    mr_note: str,
    comm_keywords: list = None,
    power_keywords: list = None,
    pco_keywords: list = None,
    aux5_keywords: list = None
) -> tuple:
    """
    Analyze mr_note to determine Aux Data 4 and 5 values.
    
    Args:
        mr_note (str): The mr_note text to analyze
        comm_keywords (list, optional): List of communication owner keywords
        power_keywords (list, optional): List of power owner keywords
        pco_keywords (list, optional): List of PCO keywords
        
    Returns:
        tuple: (aux_data_4, aux_data_5) - both in ALL CAPS format
    """
    if not mr_note or mr_note.strip() == '':
        return "NO MAKE READY", "NO"
    
    mr_note_upper = mr_note.upper()
    
    # Normalize provided keyword lists
    comm_keywords = _normalize_keywords(comm_keywords)
    power_keywords = _normalize_keywords(power_keywords)
    pco_keywords = _normalize_keywords(pco_keywords)
    aux5_keywords = _normalize_keywords(aux5_keywords)
    
    # Check for PCO keywords (highest priority for Aux Data 4)
    if any(keyword in mr_note_upper for keyword in pco_keywords):
        aux_data_4 = "PCO"
    else:
        # Check for owner mentions
        has_comm = any(keyword in mr_note_upper for keyword in comm_keywords)
        has_power = any(keyword in mr_note_upper for keyword in power_keywords)
        
        if has_comm and has_power:
            aux_data_4 = "POWER & COMM MAKE READY"
        elif has_comm:
            aux_data_4 = "COMM MAKE READY"
        elif has_power:
            aux_data_4 = "POWER MAKE READY"
        else:
            aux_data_4 = "NO MAKE READY"
    
    # Check for Aux Data 5 (Proposed Riser)
    aux_data_5 = "YES" if aux5_keywords and any(keyword in mr_note_upper for keyword in aux5_keywords) else "NO"
    
    return aux_data_4, aux_data_5

def determine_aux_data_values(
    scid: str,
    mr_note: str,
    excel_data: dict = None,
    comm_keywords: list = None,
    power_keywords: list = None,
    pco_keywords: list = None,
    aux5_keywords: list = None
) -> dict:
    """
    Determine Aux Data values based on SCID, mr_note, and Excel data.
    
    Args:
        scid (str): The SCID from filename
        mr_note (str): The mr_note from Excel
        excel_data (dict, optional): Full Excel data mapping
        
    Returns:
        dict: Dictionary of aux data field names to values (ALL CAPS format)
    """
    aux_updates = {}
    
    # Get Excel row data for this SCID
    row_data = excel_data.get(scid, {}) if excel_data else {}
    
    # Aux Data 1 (Pole Owner) - from pole_tag_company or default to MVEC
    aux_updates['Aux Data 1'] = row_data.get('pole_tag_company', DEFAULT_AUX_VALUES['Aux Data 1'])
    
    # Aux Data 2 (Pole Tag) - from pole_tag_tagtext or "NO TAG"
    pole_tag = row_data.get('pole_tag_tagtext', '')
    if not pole_tag or pole_tag.strip() == '' or str(pole_tag).lower() == 'nan':
        aux_updates['Aux Data 2'] = DEFAULT_AUX_VALUES['Aux Data 2']
    else:
        aux_updates['Aux Data 2'] = str(pole_tag).strip()
    
    # Aux Data 3 (Condition) - default to EXISTING
    aux_updates['Aux Data 3'] = DEFAULT_AUX_VALUES['Aux Data 3']
    
    # Aux Data 4 (Make Ready Type) - based on mr_note content using configurable keywords
    aux_data_4, aux_data_5 = analyze_mr_note_for_aux_data(
        mr_note,
        comm_keywords=comm_keywords,
        power_keywords=power_keywords,
        pco_keywords=pco_keywords,
        aux5_keywords=aux5_keywords
    )
    aux_updates['Aux Data 4'] = aux_data_4
    
    # Aux Data 5 (Proposed) - based on mr_note content
    aux_updates['Aux Data 5'] = aux_data_5
    
    return aux_updates

# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def extract_scid_from_filename(filename: str) -> str:
    """
    Extract SCID from PPLX filename. Example: '001_Ocalc.pplx' -> '001'
    
    Args:
        filename (str): The PPLX filename
        
    Returns:
        str: The extracted SCID
    """
    if '_Ocalc.pplx' in filename:
        return filename.split('_Ocalc.pplx')[0]
    elif '_' in filename and filename.endswith('.pplx'):
        return filename.split('_')[0]
    else:
        # Fallback - remove .pplx extension
        return filename.replace('.pplx', '')

def clean_scid_keywords(scid: str, ignore_keywords: str = "") -> str:
    """
    Remove ignore keywords from SCID. Example: '005 Foreign Pole' -> '005'
    
    Args:
        scid (str): The SCID to clean
        ignore_keywords (str, optional): Comma-separated keywords to ignore
        
    Returns:
        str: The cleaned SCID
    """
    # Split keywords by comma and clean each one
    keywords = [keyword.strip() for keyword in ignore_keywords.split(',') if keyword.strip()]
    
    cleaned_scid = scid
    for keyword in keywords:
        # Remove the keyword from the SCID (case-insensitive)
        cleaned_scid = cleaned_scid.replace(keyword, '').strip()
        # Also try case-insensitive replacement
        import re
        cleaned_scid = re.sub(re.escape(keyword), '', cleaned_scid, flags=re.IGNORECASE).strip()
    
    # Clean up any extra spaces
    cleaned_scid = ' '.join(cleaned_scid.split())
    
    return cleaned_scid

def normalize_scid_for_excel_lookup(scid: str) -> str:
    """
    Normalize SCID for Excel lookup by removing periods and spaces.
    
    Args:
        scid (str): The SCID to normalize
        
    Returns:
        str: The normalized SCID
    """
    # Remove periods and spaces to match Excel format
    normalized = scid.replace('.', '').replace(' ', '')
    return normalized
