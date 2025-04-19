# operations/concatenate.py
import pandas as pd

def apply_concatenate(dataframe, cols_to_concat, new_col_name, separator, texts):
    """Concatenates multiple columns into a new one."""
    missing_cols = [col for col in cols_to_concat if col not in dataframe.columns]
    if missing_cols:
        return dataframe, ('error', texts['column_not_found'].format(col=", ".join(missing_cols)))

    new_df = dataframe.copy()
    # Ensure all columns are string type before concatenation
    new_df[new_col_name] = new_df[cols_to_concat].astype(str).agg(separator.join, axis=1)

    # Optionally drop original columns? For now, keep them.

    return new_df, ('success', texts['concatenate_success'].format(new_col=new_col_name, count=len(cols_to_concat)))
