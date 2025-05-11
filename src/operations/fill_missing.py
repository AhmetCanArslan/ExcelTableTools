# operations/fill_missing.py
import pandas as pd # Added import
import numpy as np # pandas uses numpy for NaN

def fill_missing(data, fill_value):
    """Fills missing values (NaN, None, empty strings)."""
    # Check for pandas NaN, None, or empty string after stripping
    if pd.isna(data) or str(data).strip() == '' :
        return fill_value
    return data
