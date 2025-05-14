# operations/fill_missing.py
import pandas as pd # Added import
import numpy as np # pandas uses numpy for NaN

def fill_missing(data, fill_value, column_name=None):
    """Fills missing values (NaN, None, empty strings)."""
    # If data is the column name itself, do not fill it.
    if column_name is not None and str(data) == str(column_name):
        return data
        
    # Check for pandas NaN, None, or empty string after stripping
    if pd.isna(data) or str(data).strip() == '' :
        return fill_value
    return data
