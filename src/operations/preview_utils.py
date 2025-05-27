import re
import pandas as pd
from tkinter import simpledialog, messagebox

from .masking import mask_data, mask_words
from .trimming import trim_spaces
from .splitting import apply_split_surname, apply_split_by_delimiter
from .case_change import change_case
from .find_replace import find_replace
from .remove_chars import remove_chars
from .extract_pattern import apply_extract_pattern
from .fill_missing import fill_missing
from .duplicates import apply_mark_duplicates, apply_remove_duplicates
from .concatenate import apply_concatenate
from .merge_columns import apply_merge_columns
from .rename_column import apply_rename_column   
from .numeric_operations import apply_round_numbers, apply_calculate_column_constant, apply_create_calculated_column
from .validate_inputs import apply_validation

# Minimal texts dictionary for preview operations
PREVIEW_TEXTS = {
    'column_not_found': "Column '{col}' not found.",
    'check_valid_inputs_success': "Checked validity in column '{col}' with type '{type}'.",
    'validation_color_applied': "Invalid values in column '{col}' will be highlighted in red when saved.",
    'validation_email': "Email addresses",
    'validation_phone': "Phone numbers",
    'validation_date': "Date format",
    'validation_numeric': "Numeric values",
    'validation_alphanumeric': "Alphanumeric text",
    'validation_url': "URL addresses"
}

def apply_operation_to_partition(df, operation_type, operation_params):
    """Helper function to apply operation to a partition."""
    if operation_type == 'column_operation':
        column = operation_params.get('column')
        op_key = operation_params.get('key')
        
        # Create a copy of the dataframe to avoid modifying the original
        df = df.copy()
        
        # Import necessary functions based on operation type
        if op_key == "op_mask":
            from operations.masking import mask_data
            df[column] = df[column].astype(str).apply(mask_data, column_name=column)
        elif op_key == "op_mask_email":
            from operations.masking import mask_data
            invalid_mask = pd.Series(False, index=df.index)
            result_series = df[column].astype(str).apply(
                lambda x: mask_data(x, mode='email', column_name=column, track_invalid=True)
            )
            df[column] = result_series.apply(lambda x: x[0] if isinstance(x, tuple) else x)
            invalid_mask = result_series.apply(lambda x: isinstance(x, tuple) and not x[1])
            if not hasattr(df, '_styled_columns'):
                object.__setattr__(df, '_styled_columns', {})
            df._styled_columns[column] = invalid_mask
        elif op_key == "op_mask_words":
            from operations.masking import mask_words
            df[column] = df[column].astype(str).apply(mask_words, column_name=column)
        elif op_key == "op_trim":
            from operations.trimming import trim_spaces
            orig = df[column].astype(str)
            df[column] = orig.apply(trim_spaces, column_name=column)
            # Track changes for highlighting
            changed = orig != df[column]
            if not hasattr(df, '_styled_columns'):
                object.__setattr__(df, '_styled_columns', {})
            df._styled_columns[column] = changed
        elif op_key == "op_upper":
            from operations.case_change import change_case
            df[column] = df[column].astype(str).apply(change_case, case_type='upper', column_name=column)
        elif op_key == "op_lower":
            from operations.case_change import change_case
            df[column] = df[column].astype(str).apply(change_case, case_type='lower', column_name=column)
        elif op_key == "op_title":
            from operations.case_change import change_case
            df[column] = df[column].astype(str).apply(change_case, case_type='title', column_name=column)
        elif op_key == "op_remove_non_numeric":
            from operations.remove_chars import remove_chars
            orig = df[column].astype(str)
            df[column] = orig.apply(remove_chars, mode='non_numeric', column_name=column)
            # Track changes for highlighting
            changed = orig != df[column]
            if not hasattr(df, '_styled_columns'):
                object.__setattr__(df, '_styled_columns', {})
            df._styled_columns[column] = changed
        elif op_key == "op_remove_non_alpha":
            from operations.remove_chars import remove_chars
            orig = df[column].astype(str)
            df[column] = orig.apply(remove_chars, mode='non_alphabetic', column_name=column)
            # Track changes for highlighting
            changed = orig != df[column]
            if not hasattr(df, '_styled_columns'):
                object.__setattr__(df, '_styled_columns', {})
            df._styled_columns[column] = changed
        elif op_key.startswith("op_validate_"):
            from operations.validate_inputs import apply_validation
            validation_type = op_key.replace("op_validate_", "")
            df, result = apply_validation(df, column, validation_type, PREVIEW_TEXTS)
            if result[0] != 'success':
                raise Exception(result[1])
            
    return df

def generate_preview(app, op_key, selected_col, current_preview_df, PREVIEW_ROWS):
    """Generate a preview of the operation on the data."""
    try:
        # Create a copy of the preview data
        preview_df = current_preview_df.copy()
        
        # Create operation parameters
        operation = {
            'type': 'column_operation',
            'key': op_key,
            'column': selected_col
        }
        
        # Apply the operation
        preview_df = apply_operation_to_partition(preview_df, operation['type'], operation)
        
        # Check if any styling information was added
        has_styling = hasattr(preview_df, '_styled_columns') and selected_col in preview_df._styled_columns
        
        if has_styling:
            return preview_df, True, "Preview generated with highlighted changes"
        return preview_df, True, "Preview generated successfully"
    except Exception as e:
        return current_preview_df, False, str(e)
