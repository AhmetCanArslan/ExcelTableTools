# operations/duplicates.py
import pandas as pd

def apply_mark_duplicates(dataframe, col, new_col_name, texts, selected_columns=None):
    """Highlights duplicate values across the specified columns."""
    # If selected_columns is provided, use those; otherwise use the single column
    if selected_columns:
        columns_to_check = selected_columns
    else:
        columns_to_check = [col]
    
    # Check if all columns exist
    missing_cols = [c for c in columns_to_check if c not in dataframe.columns]
    if missing_cols:
        return dataframe, ('error', texts['column_not_found'].format(col=", ".join(missing_cols)))

    new_df = dataframe.copy()
    
    # Collect all values from selected columns
    all_values = []
    for column in columns_to_check:
        all_values.extend(new_df[column].astype(str).tolist())
    
    # Find values that appear more than once
    value_counts = pd.Series(all_values).value_counts()
    duplicate_values = set(value_counts[value_counts > 1].index)
    
    # Set up the styling for highlighting duplicates in red
    if not hasattr(new_df, '_styled_columns'):
        # Use object.__setattr__ to avoid pandas warning
        object.__setattr__(new_df, '_styled_columns', {})
    
    # Mark cells that contain duplicate values in each selected column
    total_duplicate_cells = 0
    for column in columns_to_check:
        # Create mask for cells containing duplicate values
        duplicated_mask = new_df[column].astype(str).isin(duplicate_values)
        
        # Save which cells should be highlighted
        new_df._styled_columns[column] = duplicated_mask
        
        # Count duplicate cells in this column
        total_duplicate_cells += duplicated_mask.sum()
    
    if len(columns_to_check) == 1:
        message = texts['duplicates_marked_success'].format(
            col=columns_to_check[0], count=total_duplicate_cells)
    else:
        message = texts['duplicates_marked_success_multiple'].format(
            cols=", ".join(columns_to_check), count=total_duplicate_cells)
    
    return new_df, ('success', message)

def apply_remove_duplicates(dataframe, col, texts):
    """Removes duplicate rows based on the specified column, keeping the first occurrence."""
    if col not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=col))

    original_row_count = len(dataframe)
    new_df = dataframe.drop_duplicates(subset=[col], keep='first').copy()
    rows_removed = original_row_count - len(new_df)

    return new_df, ('success', texts['duplicates_removed_success'].format(col=col, count=rows_removed))
