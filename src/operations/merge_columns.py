import pandas as pd

def _safe_str(val):
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val)

def apply_merge_columns(dataframe, cols_to_merge, new_col_name, separator, fill_missing, texts):
    """Merges multiple columns into a new one, handling missing values."""
    missing_cols = [col for col in cols_to_merge if col not in dataframe.columns]
    if missing_cols:
        return dataframe, ('error', texts['column_not_found'].format(col=", ".join(missing_cols)))

    new_df = dataframe.copy()

    # If fill_missing is True, fill missing values with an empty string
    if fill_missing:
        new_df[cols_to_merge] = new_df[cols_to_merge].fillna('')
    
    # Use _safe_str for each value to avoid .0 for integers
    new_df[new_col_name] = new_df[cols_to_merge].applymap(_safe_str).agg(separator.join, axis=1)

    return new_df, ('success', texts['merge_success'].format(new_col=new_col_name, count=len(cols_to_merge)))
