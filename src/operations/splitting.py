# operations/splitting.py
import pandas as pd

def split_surname(full_name, column_name=None):
    """Splits the last word (assumed surname) from the full name."""
    if column_name is not None and str(full_name) == str(column_name):
        return full_name, "" # Return original name and empty surname
        
    name_str = str(full_name).strip()
    if pd.isna(full_name) or name_str == '' or name_str.lower() == 'nan':
        return "", ""
        
    parts = name_str.split()
    if len(parts) > 1:
        surname = parts[-1]
        name_part = " ".join(parts[:-1])
        return name_part, surname
    else:
        # Handle single names - return name as is, empty surname
        return name_str, ""

def apply_split_surname(dataframe, col, texts):
    """Applies surname splitting. Returns modified dataframe and status message info."""
    if col not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=col))

    # Create a copy to avoid modifying the original DataFrame
    new_df = dataframe.copy()
    
    split_results = new_df[col].apply(split_surname, column_name=col)
    name_series = split_results.apply(lambda x: x[0])
    surname_series = split_results.apply(lambda x: x[1])

    new_surname_col_name = f"{col}_Surname"
    counter = 1
    base_name = new_surname_col_name
    while new_surname_col_name in new_df.columns:
        new_surname_col_name = f"{base_name}_{counter}"
        counter += 1

    original_col_index = new_df.columns.get_loc(col)
    new_df[col] = name_series
    new_df.insert(original_col_index + 1, new_surname_col_name, surname_series)

    return new_df, ('success', texts['surname_split_success'].format(col=col, new_col=new_surname_col_name))

def apply_split_by_delimiter(dataframe, col, delimiter, texts):
    """Applies delimiter splitting. Returns modified dataframe and status message info."""
    if col not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=col))

    if delimiter is None or delimiter == '':
        return dataframe, ('error', "Delimiter cannot be empty. If you intended to use a space, please ensure the delimiter is preserved during the operation.")

    # Create a copy to avoid modifying the original DataFrame
    new_df = dataframe.copy()
    col_data = new_df[col].astype(str)

    # Check if delimiter exists in the data
    contains_delimiter = col_data.str.contains(delimiter, regex=False, na=False)
    if not contains_delimiter.any():
        return dataframe, ('warning', texts['split_warning_delimiter_not_found'].format(delimiter=delimiter, col=col))

    # Split the data
    split_data = col_data.str.split(delimiter, expand=True)
    
    # Generate column names
    num_parts = split_data.shape[1]
    new_cols_base = [f"{col}_part{i+1}" for i in range(num_parts)]

    # Ensure new column names are unique
    existing_cols = set(new_df.columns)
    final_new_cols = []
    for new_col_base in new_cols_base:
        new_col = new_col_base
        counter = 1
        while new_col in existing_cols or new_col in final_new_cols:
            new_col = f"{new_col_base}_{counter}"
            counter += 1
        final_new_cols.append(new_col)

    split_data.columns = final_new_cols

    # Clean up the split data - handle NaN and convert to appropriate strings
    for c in split_data.columns:
        split_data[c] = split_data[c].fillna('').astype(str)

    # Insert the new columns at the position of the original column
    original_col_index = new_df.columns.get_loc(col)
    
    # Remove the original column
    new_df = new_df.drop(columns=[col])
    
    # Insert the split columns
    for i, new_col in enumerate(final_new_cols):
        new_df.insert(original_col_index + i, new_col, split_data[new_col])

    return new_df, ('success', texts['split_success'].format(col=col, delimiter=delimiter, count=len(final_new_cols)))
