# operations/extract_pattern.py
import pandas as pd
import re

def apply_extract_pattern(dataframe, col, new_col_name, pattern, texts):
    """Extracts data using regex pattern into a new column."""
    if col not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=col))

    try:
        # Attempt to compile regex to catch errors early
        compiled_pattern = re.compile(pattern)
    except re.error as e:
        return dataframe, ('error', texts['regex_error'].format(error=e))

    new_df = dataframe.copy()
    # Extract first match, fill non-matches with empty string
    new_df[new_col_name] = new_df[col].astype(str).str.extract(compiled_pattern, expand=False).fillna('')

    # Convert to safe string (remove .0 for ints)
    def _safe_str(val):
        try:
            f = float(val)
            if f.is_integer():
                return str(int(f))
        except Exception:
            pass
        return str(val)
    new_df[new_col_name] = new_df[new_col_name].apply(_safe_str)

    original_col_index = new_df.columns.get_loc(col)
    # Insert the new column right after the original column (if it still exists)
    # Need to reorder columns if insert isn't sufficient or original is dropped
    if new_col_name in new_df.columns: # Check if extraction created the column
        new_col_series = new_df.pop(new_col_name)
        new_df.insert(original_col_index + 1, new_col_name, new_col_series)

    return new_df, ('success', texts['extract_success'].format(col=col, new_col=new_col_name))
