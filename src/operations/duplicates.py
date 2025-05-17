# operations/duplicates.py
import pandas as pd

def apply_mark_duplicates(dataframe, col, new_col_name, texts):
    """Highlights duplicate values in the specified column."""
    if col not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=col))

    new_df = dataframe.copy()
    
    # Mark duplicates (keep=False marks all duplicates as True)
    duplicated_mask = new_df.duplicated(subset=[col], keep=False)
    
    # Set up the styling for highlighting duplicates in red
    if not hasattr(new_df, '_styled_columns'):
        # Use object.__setattr__ to avoid pandas warning
        object.__setattr__(new_df, '_styled_columns', {})
    
    # Save which cells should be highlighted
    new_df._styled_columns[col] = duplicated_mask
    
    # Count duplicates for the message
    duplicate_count = duplicated_mask.sum()
    
    return new_df, ('success', texts['duplicates_marked_success'].format(
        col=col, count=duplicate_count))

def apply_remove_duplicates(dataframe, col, texts):
    """Removes duplicate rows based on the specified column, keeping the first occurrence."""
    if col not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=col))

    original_row_count = len(dataframe)
    new_df = dataframe.drop_duplicates(subset=[col], keep='first').copy()
    rows_removed = original_row_count - len(new_df)

    return new_df, ('success', texts['duplicates_removed_success'].format(col=col, count=rows_removed))
