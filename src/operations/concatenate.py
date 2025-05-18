# operations/concatenate.py
import pandas as pd

def _safe_str(val):
    # Prevent float->int conversion like 123.0 -> '123'
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val)

def apply_concatenate(dataframe, cols_to_concat, new_col_name, separator, texts):
    """Concatenates multiple columns into a new one."""
    missing_cols = [col for col in cols_to_concat if col not in dataframe.columns]
    if missing_cols:
        return dataframe, ('error', texts['column_not_found'].format(col=", ".join(missing_cols)))

    new_df = dataframe.copy()
    # Use _safe_str for each value to avoid .0 for integers
    new_df[new_col_name] = new_df[cols_to_concat].apply(lambda row: separator.join(_safe_str(val) for val in row), axis=1)

    return new_df, ('success', texts['concatenate_success'].format(new_col=new_col_name, count=len(cols_to_concat)))
