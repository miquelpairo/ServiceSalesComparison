"""
Column Mappings for Sales Comparison Dashboard
Supports multiple input formats from Power BI exports
"""

# =============================================================================
# FORMAT DEFINITIONS
# =============================================================================

# Original format (backwards compatibility)
ORIGINAL_FORMAT = {
    'Date': 'Date',
    'Business Partner Name': 'Business Partner Name',
    'ItemIdAndName': 'ItemIdAndName',
    'ProductType': 'ProductType',
    'Qty': 'Qty',
    'EUR': 'EUR',
    'SalesRepresentative': 'SalesRepresentative',
    'Set': 'Set',
    'Productline': 'Productline'
}

# New multi-currency format
NEW_FORMAT = {
    'Date': 'Date',
    'Business Partner Name': 'End User',  # Mapped from End User
    'ItemIdAndName': 'Id - Name.1',  # Second Id - Name column (product)
    'ProductType': 'Product Type',  # Space instead of no space
    'Qty': 'Qty',
    'EUR': 'LC',  # Local Currency
    'SalesRepresentative': 'Sales Representative',  # Space instead of no space
    'Set': 'Set',
    'Productline': 'Product Line'  # Space instead of no space
}

# Mixed format (actual current Power BI export format)
# This format has columns already standardized but with slight variations
MIXED_FORMAT = {
    'Date': 'Date',
    'Business Partner Name': 'End User',  # Maps from End User column
    'ItemIdAndName': 'Id - Name.1',  # Second Id - Name column (product/service) - pandas auto-renames to .1
    'ProductType': 'Product Type',  # Has space
    'Qty': 'Qty',
    'EUR': 'LC',  # Local Currency
    'SalesRepresentative': 'Sales Representative',  # Has space
    'Set': 'Set',
    'Productline': 'Product Line'  # Has space
}

# Additional columns to preserve from new/mixed format
ADDITIONAL_COLUMNS_NEW_FORMAT = [
    'Market Organization Name',
    'Sales Territory',
    'Country',
    'End User Segment',
    'SFDC Link',
    'Segment',
    'Postal Code',
    'Document Number',
    'Position',
    'FC',  # Foreign Currency
    'CHF'  # Swiss Francs
    # Note: 'Id - Name' (first one) is intentionally excluded as it's not needed
]

# =============================================================================
# REQUIRED COLUMNS (standardized internal names)
# =============================================================================

REQUIRED_COLUMNS = [
    'Date',
    'Business Partner Name',
    'ItemIdAndName',
    'ProductType',
    'Qty',
    'EUR',
    'SalesRepresentative',
    'Set',
    'Productline'
]

# =============================================================================
# DETECTION FUNCTIONS
# =============================================================================

def detect_format(columns):
    """
    Detect which format the input file uses
    
    Args:
        columns (list): List of column names from the DataFrame
        
    Returns:
        str: 'original', 'new', 'mixed', or 'unknown'
    """
    columns_set = set(columns)
    
    # Check for mixed format signature (current actual format)
    # Has: End User, LC, Product Type (with space), Sales Representative (with space), Product Line (with space), Id - Name.1
    mixed_format_signatures = ['End User', 'LC', 'Product Type', 'Sales Representative', 'Product Line', 'Id - Name.1']
    mixed_format_match = sum(1 for sig in mixed_format_signatures if sig in columns_set)
    
    # Check for new format signature columns (pure new format - not used yet)
    new_format_signatures = ['End User', 'LC', 'Id - Name.1']
    new_format_match = sum(1 for sig in new_format_signatures if sig in columns_set)
    
    # Check for original format signature columns
    original_format_signatures = ['Business Partner Name', 'EUR', 'ProductType', 'SalesRepresentative', 'Productline', 'ItemIdAndName']
    original_format_match = sum(1 for sig in original_format_signatures if sig in columns_set)
    
    # Detect based on matches
    if mixed_format_match >= 5:  # If at least 5 signatures match, it's mixed format
        return 'mixed'
    elif original_format_match >= 5:
        return 'original'
    elif new_format_match >= 2:
        return 'new'
    else:
        return 'unknown'

def get_mapping_for_format(format_type):
    """
    Get the column mapping dictionary for a given format
    
    Args:
        format_type (str): 'original', 'new', or 'mixed'
        
    Returns:
        dict: Mapping from standardized names to actual column names
    """
    if format_type == 'original':
        return ORIGINAL_FORMAT
    elif format_type == 'new':
        return NEW_FORMAT
    elif format_type == 'mixed':
        return MIXED_FORMAT
    else:
        return None

def get_additional_columns(format_type):
    """
    Get list of additional columns to preserve
    
    Args:
        format_type (str): 'original' or 'new'
        
    Returns:
        list: List of additional column names to preserve
    """
    if format_type == 'new':
        return ADDITIONAL_COLUMNS_NEW_FORMAT
    else:
        return []

# =============================================================================
# VALIDATION FUNCTIONS
# =============================================================================

def validate_format(df, format_type):
    """
    Validate that DataFrame has all required columns for the detected format
    
    Args:
        df (pd.DataFrame): Input DataFrame
        format_type (str): 'original' or 'new'
        
    Returns:
        tuple: (is_valid (bool), missing_columns (list))
    """
    mapping = get_mapping_for_format(format_type)
    if mapping is None:
        return False, []
    
    missing_columns = []
    for standard_name, actual_name in mapping.items():
        if actual_name not in df.columns:
            missing_columns.append(actual_name)
    
    is_valid = len(missing_columns) == 0
    return is_valid, missing_columns

def get_format_info(format_type):
    """
    Get human-readable information about a format
    
    Args:
        format_type (str): 'original', 'new', or 'mixed'
        
    Returns:
        dict: Information about the format
    """
    if format_type == 'original':
        return {
            'name': 'Original Format',
            'description': 'Power BI export with EUR currency and original column names',
            'currency': 'EUR',
            'version': 'v1.0-v2.0'
        }
    elif format_type == 'new':
        return {
            'name': 'Multi-Currency Format',
            'description': 'Power BI export with LC/FC/CHF currencies and extended data',
            'currency': 'LC (Local Currency)',
            'version': 'v3.0+'
        }
    elif format_type == 'mixed':
        return {
            'name': 'Current Format',
            'description': 'Power BI export with End User, LC currency, and column names with spaces',
            'currency': 'LC (Local Currency)',
            'version': 'v3.0 (Current)'
        }
    else:
        return {
            'name': 'Unknown Format',
            'description': 'Format not recognized',
            'currency': 'Unknown',
            'version': 'Unknown'
        }