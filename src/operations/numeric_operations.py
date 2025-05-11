import pandas as pd
import numpy as np

def _to_numeric_coerce(series):
    return pd.to_numeric(series, errors='coerce')

def apply_round_numbers(dataframe, col, decimals, texts):
    """Rounds numbers in a column to specified decimal places."""
    if col not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=col))
    
    new_df = dataframe.copy()
    numeric_col = _to_numeric_coerce(new_df[col])
    
    if numeric_col.isnull().all(): # All values became NaN after coercion
        return dataframe, ('warning', texts['column_not_numeric'].format(col=col))

    new_df[col] = numeric_col.round(decimals)
    return new_df, ('success', texts['rounding_success'].format(col=col, decimals=decimals))

def apply_calculate_column_constant(dataframe, col, operation, value, texts):
    """Performs a calculation (add, subtract, multiply, divide) on a column with a constant."""
    if col not in dataframe.columns:
        return dataframe, ('error', texts['column_not_found'].format(col=col))

    new_df = dataframe.copy()
    numeric_col = _to_numeric_coerce(new_df[col])

    if numeric_col.isnull().all():
        return dataframe, ('warning', texts['column_not_numeric'].format(col=col))

    if operation == '+':
        new_df[col] = numeric_col + value
    elif operation == '-':
        new_df[col] = numeric_col - value
    elif operation == '*':
        new_df[col] = numeric_col * value
    elif operation == '/':
        if value == 0:
            return dataframe, ('error', texts['division_by_zero'].format(col=col))
        new_df[col] = numeric_col / value
    else:
        return dataframe, ('error', f"Unknown operation: {operation}")
        
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

    result_col = None
    if operation == '+':
        result_col = col1_numeric + col2_numeric
    elif operation == '-':
        result_col = col1_numeric - col2_numeric
    elif operation == '*':
        result_col = col1_numeric * col2_numeric
    elif operation == '/':
        # Handle division by zero: result will be inf or -inf, or NaN if 0/0.
        # Replace inf/-inf with NaN, or let pandas handle it. For now, let it be.
        if (col2_numeric == 0).any():
             # Informative warning, actual division by zero will result in inf/NaN handled by pandas
            pass # texts['division_by_zero'] could be used if we want to pre-emptively stop
        result_col = col1_numeric / col2_numeric
        # Replace inf with NaN for cleaner output, if desired
        result_col.replace([np.inf, -np.inf], np.nan, inplace=True)
    else:
        return dataframe, ('error', f"Unknown operation: {operation}")

    new_df[new_col_name] = result_col
    return new_df, ('success', texts['create_column_success'].format(new_col=new_col_name, col1=col1_name, col2=col2_name))
