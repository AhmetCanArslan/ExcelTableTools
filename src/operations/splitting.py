# operations/splitting.py
import pandas as pd

def split_surname(full_name, column_name=None):
    """Splits the last word (assumed surname) from the full name."""
    if column_name is not None and str(full_name) == str(column_name):
        return full_name, "" # Return original name and empty surname
        
    name_str = str(full_name).strip()
    parts = name_str.split()
    if len(parts) > 1:
        surname = parts[-1]
        name_part = " ".join(parts[:-1])
        return name_part, surname
    else:
        # Handle single names or empty strings - return name as is, empty surname
        return name_str, ""

def apply_split_surname(dataframe, col, texts):
    """Applies surname splitting. Returns modified dataframe and status message info."""
    if col not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=col))

    split_results = dataframe[col].apply(split_surname, column_name=col)
    name_series = split_results.apply(lambda x: x[0])
    surname_series = split_results.apply(lambda x: x[1])

    new_surname_col_name = f"{col}_Surname"
    counter = 1
    base_name = new_surname_col_name
    while new_surname_col_name in dataframe.columns:
        new_surname_col_name = f"{base_name}_{counter}"
        counter += 1

    original_col_index = dataframe.columns.get_loc(col)
    # Create a copy to avoid modifying the original DataFrame directly before returning
    new_df = dataframe.copy()
    new_df[col] = name_series
    new_df.insert(original_col_index + 1, new_surname_col_name, surname_series)

    return new_df, ('success', texts['surname_split_success'].format(col=col, new_col=new_surname_col_name))

def apply_split_by_delimiter(dataframe, col, delimiter, texts):
    """Applies delimiter splitting. Returns modified dataframe and status message info."""
    if col not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=col))

    col_data = dataframe[col].astype(str) # Keep as astype(str) for general case

    # Identify rows to skip special handling for (value equals column name)
    # These rows' first split part will be the original value, others empty.
    skip_mask = (dataframe[col].astype(str) == str(col))
    
    if not col_data.str.contains(delimiter, regex=False).any() and not skip_mask.any():
        # If delimiter not found anywhere AND no rows match column name (which would be "split" differently)
        # then it's a true "delimiter not found" scenario.
        # If skip_mask.any() is true, some rows will be "split" into [value, '', ...], so proceed.
        is_delimiter_present_in_non_skipped = col_data[~skip_mask].str.contains(delimiter, regex=False).any()
        if not is_delimiter_present_in_non_skipped:
             return dataframe, ('warning', texts['split_warning_delimiter_not_found'].format(delimiter=delimiter, col=col))

    max_splits = col_data.str.split(delimiter).str.len().max()
    # If all values are like column_name, max_splits could be 1. Ensure at least 1 for the column itself.
    if skip_mask.all(): # If all rows are to be skipped
        max_splits = 1 
    
    new_cols_base = [f"{col}_part{i+1}" for i in range(max_splits)]

    # Ensure new column names are unique
    existing_cols = set(dataframe.columns)
    final_new_cols = []
    for new_col_base in new_cols_base:
        new_col = new_col_base
        counter = 1
        while new_col in existing_cols or new_col in final_new_cols:
            new_col = f"{new_col_base}_{counter}"
            counter += 1
        final_new_cols.append(new_col)

    split_data = col_data.str.split(delimiter, expand=True, n=max_splits - 1 if max_splits > 0 else 0)

    # Pad with empty columns if split result has fewer columns than max_splits
    if split_data.shape[1] < len(final_new_cols):
        for i in range(split_data.shape[1], len(final_new_cols)):
            split_data[i] = '' # Add empty columns
    
    split_data.columns = final_new_cols

    # Convert all values to safe string (remove .0 for ints)
    def _safe_str(val):
        try:
            f = float(val)
            if f.is_integer():
                return str(int(f))
        except Exception:
            pass
        return str(val)
    for c in split_data.columns:
        split_data[c] = split_data[c].apply(_safe_str)

    # For rows where original value was column name, adjust the split_data
    # The original column `col` is effectively replaced by `final_new_cols[0]`
    if skip_mask.any():
        rows_to_adjust_indices = dataframe.index[skip_mask]
        if not rows_to_adjust_indices.empty:
            # Ensure split_data has these indices
            valid_indices_for_adjustment = rows_to_adjust_indices.intersection(split_data.index)
            if not valid_indices_for_adjustment.empty:
                split_data.loc[valid_indices_for_adjustment, final_new_cols[0]] = dataframe.loc[valid_indices_for_adjustment, col]
                for k_idx in range(1, len(final_new_cols)):
                    split_data.loc[valid_indices_for_adjustment, final_new_cols[k_idx]] = ''

    original_col_index = dataframe.columns.get_loc(col)

    # Create a new DataFrame by concatenating parts
    df_before = dataframe.iloc[:, :original_col_index]
    df_after = dataframe.iloc[:, original_col_index+1:]
    new_df = pd.concat([df_before, split_data, df_after], axis=1)

    return new_df, ('success', texts['split_success'].format(col=col, delimiter=delimiter, count=len(final_new_cols)))
