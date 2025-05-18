import pandas as pd
import numpy as np

def _to_numeric_coerce(series):
    return pd.to_numeric(series, errors='coerce')

def apply_round_numbers(dataframe, col, decimals, texts):
    """Rounds numbers in a column to specified decimal places."""
    if col not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=col))
    
    new_df = dataframe.copy()
    
    # Identify rows where value equals column name (as string)
    skip_mask = (new_df[col].astype(str) == str(col))
    
    # Process only non-skipped rows
    col_to_process = new_df.loc[~skip_mask, col]
    
    if col_to_process.empty and skip_mask.all(): # All rows skipped
        return new_df, ('success', texts['rounding_success'].format(col=col, decimals=decimals) + " (all rows matched column name and were skipped)")

    numeric_col_processed = _to_numeric_coerce(col_to_process)
    
    if numeric_col_processed.isnull().all() and not col_to_process.empty:
        # This means all rows that were attempted to be processed became NaN
        return dataframe, ('warning', texts['column_not_numeric'].format(col=col) + " (for non-skipped rows)")

    new_df.loc[~skip_mask, col] = numeric_col_processed.round(decimals)
    return new_df, ('success', texts['rounding_success'].format(col=col, decimals=decimals))

def apply_calculate_column_constant(dataframe, col, operation, value, texts):
    """Performs a calculation (add, subtract, multiply, divide) on a column with a constant."""
    if col not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=col))

    new_df = dataframe.copy()

    # Identify rows where value equals column name (as string)
    skip_mask = (new_df[col].astype(str) == str(col))
    
    # Process only non-skipped rows
    col_to_process = new_df.loc[~skip_mask, col]

    if col_to_process.empty and skip_mask.all(): # All rows skipped
        return new_df, ('success', texts['calculation_success'].format(col=col) + " (all rows matched column name and were skipped)")

    numeric_col_processed = _to_numeric_coerce(col_to_process)

    if numeric_col_processed.isnull().all() and not col_to_process.empty:
        return dataframe, ('warning', texts['column_not_numeric'].format(col=col) + " (for non-skipped rows)")

    calculated_values = None
    if operation == '+':
        calculated_values = numeric_col_processed + value
    elif operation == '-':
        calculated_values = numeric_col_processed - value
    elif operation == '*':
        calculated_values = numeric_col_processed * value
    elif operation == '/':
        if value == 0:
            # Apply division by zero only to relevant part
            # Check if any non-NaN in numeric_col_processed would be divided by zero
            if not numeric_col_processed[numeric_col_processed.notna()].empty:
                 # If there are actual numbers to be divided
                return dataframe, ('error', texts['division_by_zero'].format(col=col))
            # If all are NaN or skipped, division by zero might not "happen" to a value
            # but it's still an invalid operation setup.
            # However, if col_to_process was empty, this path might not be hit.
            # Safest to prevent division by zero if value is 0.
            return dataframe, ('error', texts['division_by_zero'].format(col=col))

        calculated_values = numeric_col_processed / value
        # Replace inf/-inf with NaN for cleaner output if desired, or handle as per requirements
        # calculated_values.replace([np.inf, -np.inf], np.nan, inplace=True)
    else:
        return dataframe, ('error', f"Unknown operation: {operation}")
    
    # Explicitly cast to float to avoid dtype warning
    calculated_values = calculated_values.astype(float)
    new_df.loc[~skip_mask, col] = calculated_values
        
    return new_df, ('success', texts['calculation_success'].format(col=col))

def apply_create_calculated_column(dataframe, col1_name, col2_name, operation, new_col_name, texts):
    """Creates a new column by performing an operation on two existing columns."""
    if col1_name not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=col1_name))
    if col2_name not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=col2_name))
    if new_col_name in dataframe.columns: # Should be checked by caller, but good to have
        return dataframe, ('error', texts['column_already_exists'].format(name=new_col_name))

    new_df = dataframe.copy()
    col1_numeric = _to_numeric_coerce(new_df[col1_name])
    col2_numeric = _to_numeric_coerce(new_df[col2_name])

    if col1_numeric.isnull().all():
        return dataframe, ('warning', texts['column_not_numeric'].format(col=col1_name))
    if col2_numeric.isnull().all():
        return dataframe, ('warning', texts['column_not_numeric'].format(col=col2_name))

    # Determine decimal precision based on operands
    def get_decimal_places(series):
        # Extract non-null values that are numeric
        numeric_values = series.dropna().apply(lambda x: str(x) if pd.notna(x) else '')
        # Get decimal places for each value
        decimals = numeric_values.apply(lambda x: len(x.split('.')[-1]) if '.' in x else 0)
        # Return max decimal places found
        return decimals.max() if not decimals.empty else 0
    
    col1_decimals = get_decimal_places(col1_numeric)
    col2_decimals = get_decimal_places(col2_numeric)
    max_decimals = max(col1_decimals, col2_decimals)

    result_col = None
    if operation == '+':
        result_col = col1_numeric + col2_numeric
    elif operation == '-':
        result_col = col1_numeric - col2_numeric
    elif operation == '*':
        result_col = col1_numeric * col2_numeric
    elif operation == '/':
        # Handle division by zero: result will be inf or -inf, or NaN if 0/0.
        if (col2_numeric == 0).any():
             # Informative warning, actual division by zero will result in inf/NaN handled by pandas
            pass # texts['division_by_zero'] could be used if we want to pre-emptively stop
        result_col = col1_numeric / col2_numeric
        # Replace inf with NaN for cleaner output, if desired
        result_col.replace([np.inf, -np.inf], np.nan, inplace=True)
    else:
        return dataframe, ('error', f"Unknown operation: {operation}")

    # Round to the same precision as the operands if needed
    if max_decimals > 0 and not result_col.isnull().all():
        result_col = result_col.round(max_decimals)

    new_df[new_col_name] = result_col
    return new_df, ('success', texts['create_column_success'].format(new_col=new_col_name, col1=col1_name, col2=col2_name))
