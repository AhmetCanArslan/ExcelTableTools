import pandas as pd
import numpy as np
from typing import Tuple, Dict, Any

def apply_distinct_group_encoding(df: pd.DataFrame, column: str) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Apply distinct group encoding to a column, creating a new column with numeric labels.
    Optimized for large datasets using pandas factorize.
    
    Args:
        df: Input DataFrame
        column: Name of the column to encode
        
    Returns:
        Tuple of (modified DataFrame, operation metadata)
    """
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found in DataFrame")
        
    # Generate new column name
    new_column = f"{column}_distinct_group"
    counter = 1
    while new_column in df.columns:
        new_column = f"{column}_distinct_group_{counter}"
        counter += 1
    
    # Use factorize which is more memory efficient than map or replace
    # as it doesn't create an intermediate Series
    codes, uniques = pd.factorize(df[column], sort=True)
    
    # Add 1 to codes to start from 1 instead of 0
    df[new_column] = codes + 1
    
    # Create mapping dictionary for preview/info
    value_mapping = {val: idx + 1 for idx, val in enumerate(uniques)}
    
    # Return metadata for operation tracking
    metadata = {
        'new_column': new_column,
        'value_mapping': value_mapping,
        'unique_values': len(value_mapping)
    }
    
    return df, metadata

def preview_distinct_group(df: pd.DataFrame, column: str, preview_rows: int = 1000) -> Tuple[pd.DataFrame, bool, str]:
    """
    Generate a preview of the distinct group encoding operation.
    
    Args:
        df: Input DataFrame
        column: Name of the column to encode
        preview_rows: Number of rows to preview
        
    Returns:
        Tuple of (preview DataFrame, success boolean, message string)
    """
    try:
        preview_df = df.head(preview_rows).copy()
        modified_df, metadata = apply_distinct_group_encoding(preview_df, column)
        
        # Create success message with mapping info
        mapping_str = "\n".join([f"{val} → {num}" for val, num in metadata['value_mapping'].items()])
        message = f"Created new column: {metadata['new_column']}\n"
        message += f"Found {metadata['unique_values']} unique values.\n"
        message += f"Mapping:\n{mapping_str}"
        
        return modified_df, True, message
        
    except Exception as e:
        return df, False, f"Error in preview: {str(e)}" 