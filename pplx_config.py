#!/usr/bin/env python3
"""
PPLX Configuration Module - Shared constants and logic for PPLX processing

This module consolidates all the shared configuration, keywords, and business logic
used across the PPLX processing applications to eliminate redundancy.
"""

# =============================================================================
# CONFIGURATION - OWNER KEYWORDS
# =============================================================================
# Users can modify these lists to add or remove keywords for classification
# Keywords are case-insensitive and searched within the mr_note field
# 
# To add a keyword: Add it to the appropriate list as a string
# To remove a keyword: Delete it from the list
# 
# Examples:
#   - Add 'spectrum' to COMM_OWNER_KEYWORDS: 'spectrum'
#   - Add 'electric' to POWER_OWNER_KEYWORDS: 'electric'
#   - Add 'pole change' to PCO_KEYWORDS: 'pole change'

# Comm Owner Keywords - triggers "COMM MAKE READY"
COMM_OWNER_KEYWORDS = [
    'comm',           # Communication companies
    'catv',           # Cable TV
    'fiber',          # Fiber optic
    'cable',          # Cable companies
    'metronet'        # MetroNet specifically
]

# Power Owner Keywords - triggers "POWER MAKE READY"
POWER_OWNER_KEYWORDS = [
    'power',          # Power companies
    'mvec',           # MVEC (power company)
    'primary',        # Primary power lines
    'secondary'       # Secondary power lines
]

# PCO Keywords - triggers "PCO" (Pole Change Out)
# Note: Currently only "REPLACING EXISTING" is used for PCO classification
# Other PCO-related keywords can be added here for future use
PCO_KEYWORDS = [
    'pco',            # Pole Change Out
    'pole replacement', # Pole replacement work
    'replacing existing' # Specific phrase for PCO classification
]

# =============================================================================
# DEFAULT VALUES
# =============================================================================

# Default AUX Data values
DEFAULT_AUX_VALUES = {
    'Aux Data 1': 'MVEC',           # Default pole owner
    'Aux Data 2': 'NO TAG',         # Default pole tag
    'Aux Data 3': 'EXISTING',       # Default condition
    'Aux Data 4': 'NO MAKE READY',  # Default make ready type
    'Aux Data 5': 'NO'              # Default proposed riser
}

# Default owner configurations
DEFAULT_COMM_OWNERS = "MetroNet, CATV, TELCO"
DEFAULT_POWER_OWNERS = "MVEC, Xcel, Consumers"

# Default SCID keywords to ignore
DEFAULT_IGNORE_SCID_KEYWORDS = "Foreign Pole, AFOREIGNPOLE"

# =============================================================================
# AUX DATA ANALYSIS LOGIC
# =============================================================================

def analyze_mr_note_for_aux_data(mr_note: str, comm_owners: list = None, power_owners: list = None) -> tuple:
    """
    Analyze mr_note to determine Aux Data 4 and 5 values.
    
    Args:
        mr_note (str): The mr_note text to analyze
        comm_owners (list, optional): List of communication owner keywords
        power_owners (list, optional): List of power owner keywords
        
    Returns:
        tuple: (aux_data_4, aux_data_5) - both in ALL CAPS format
    """
    if not mr_note or mr_note.strip() == '':
        return "NO MAKE READY", "NO"
    
    mr_note_upper = mr_note.upper()
    
    # Use provided owner lists or defaults
    if comm_owners is None:
        comm_owners = [owner.upper() for owner in COMM_OWNER_KEYWORDS]
    else:
        comm_owners = [owner.strip().upper() for owner in comm_owners if owner.strip()]
    
    if power_owners is None:
        power_owners = [owner.upper() for owner in POWER_OWNER_KEYWORDS]
    else:
        power_owners = [owner.strip().upper() for owner in power_owners if owner.strip()]
    
    # Check for REPLACING EXISTING (highest priority for Aux Data 4)
    if 'REPLACING EXISTING' in mr_note_upper:
        aux_data_4 = "PCO"
    else:
        # Check for owner mentions
        has_comm = any(owner in mr_note_upper for owner in comm_owners)
        has_power = any(owner in mr_note_upper for owner in power_owners)
        
        if has_comm and has_power:
            aux_data_4 = "POWER & COMM MAKE READY"
        elif has_comm:
            aux_data_4 = "COMM MAKE READY"
        elif has_power:
            aux_data_4 = "POWER MAKE READY"
        else:
            aux_data_4 = "NO MAKE READY"
    
    # Check for Aux Data 5 (Proposed Riser)
    aux_data_5 = "YES" if 'INSTALL METRONET RISER' in mr_note_upper else "NO"
    
    return aux_data_4, aux_data_5

def determine_aux_data_values(scid: str, mr_note: str, excel_data: dict = None) -> dict:
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
    aux_data_4, aux_data_5 = analyze_mr_note_for_aux_data(mr_note)
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

def clean_scid_keywords(scid: str, ignore_keywords: str = None) -> str:
    """
    Remove ignore keywords from SCID. Example: '005 Foreign Pole' -> '005'
    
    Args:
        scid (str): The SCID to clean
        ignore_keywords (str, optional): Comma-separated keywords to ignore
        
    Returns:
        str: The cleaned SCID
    """
    if not ignore_keywords:
        ignore_keywords = DEFAULT_IGNORE_SCID_KEYWORDS
    
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
