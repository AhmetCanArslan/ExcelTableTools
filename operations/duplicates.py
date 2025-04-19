# operations/duplicates.py
import pandas as pd

def apply_mark_duplicates(dataframe, col, new_col_name, texts):
    """Adds a boolean column marking duplicate values in the specified column."""
    if col not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=col))

    new_df = dataframe.copy()
    # keep=False marks all duplicates as True
    new_df[new_col_name] = new_df.duplicated(subset=[col], keep=False)

    return new_df, ('success', texts['duplicates_marked_success'].format(col=col, new_col=new_col_name))

def apply_remove_duplicates(dataframe, col, texts):
    """Removes duplicate rows based on the specified column, keeping the first occurrence."""
    if col not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=col))

    original_row_count = len(dataframe)
    new_df = dataframe.drop_duplicates(subset=[col], keep='first').copy()
    rows_removed = original_row_count - len(new_df)

    return new_df, ('success', texts['duplicates_removed_success'].format(col=col, count=rows_removed))
