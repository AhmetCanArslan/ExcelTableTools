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
from .distinct_group import apply_distinct_group_encoding, preview_distinct_group

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
    'validation_url': "URL addresses",
    'split_success': "Split column '{col}' by '{delimiter}' into {count} new columns.",
    'split_warning_delimiter_not_found': "The delimiter '{delimiter}' was not found in column '{col}'. No changes made.",
    'surname_split_success': "Split surname from column '{col}' into new column '{new_col}'.",
    'extract_success': "Extracted pattern from '{col}' into new column '{new_col}'.",
    'merge_success': "Merged {count} columns into new column '{new_col}'.",
    'concatenate_success': "Concatenated {count} columns into new column '{new_col}'."
}

def apply_operation_to_partition(df, operation_type, operation_params):
    """Helper function to apply operation to a partition."""
    try:
        print(f"DEBUG: apply_operation_to_partition called with operation_type='{operation_type}'")
        print(f"DEBUG: operation_params: {operation_params}")
        print(f"DEBUG: DataFrame shape: {df.shape}")
        
        if operation_type == 'column_operation':
            column = operation_params.get('column')
            op_key = operation_params.get('key')
            
            print(f"DEBUG: Processing column '{column}' with operation '{op_key}'")
            
            # Check if column exists in DataFrame
            if column not in df.columns:
                raise KeyError(f"Column '{column}' not found in the DataFrame.")
            
            # Create a copy of the dataframe to avoid modifying the original
            print(f"DEBUG: Creating DataFrame copy...")
            df = df.copy()
            print(f"DEBUG: DataFrame copy created successfully")
            
            # Import necessary functions based on operation type
            if op_key == "op_mask":
                print(f"DEBUG: Applying mask operation")
                from operations.masking import mask_data
                df[column] = df[column].astype(str).apply(mask_data, column_name=column)
            elif op_key == "op_mask_email":
                print(f"DEBUG: Applying email mask operation")
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
                print(f"DEBUG: Applying word mask operation")
                from operations.masking import mask_words
                df[column] = df[column].astype(str).apply(mask_words, column_name=column)
            elif op_key == "op_trim":
                print(f"DEBUG: Applying trim operation")
                from operations.trimming import trim_spaces
                orig = df[column].astype(str)
                df[column] = orig.apply(trim_spaces, column_name=column)
                # Track changes for highlighting
                changed = orig != df[column]
                if not hasattr(df, '_styled_columns'):
                    object.__setattr__(df, '_styled_columns', {})
                df._styled_columns[column] = changed
            elif op_key == "op_upper":
                print(f"DEBUG: Applying upper case operation")
                from operations.case_change import change_case
                df[column] = df[column].astype(str).apply(change_case, case_type='upper', column_name=column)
            elif op_key == "op_lower":
                print(f"DEBUG: Applying lower case operation")
                from operations.case_change import change_case
                df[column] = df[column].astype(str).apply(change_case, case_type='lower', column_name=column)
            elif op_key == "op_title":
                print(f"DEBUG: Applying title case operation")
                from operations.case_change import change_case
                df[column] = df[column].astype(str).apply(change_case, case_type='title', column_name=column)
            elif op_key == "op_remove_non_numeric":
                print(f"DEBUG: Applying remove non-numeric operation")
                from operations.remove_chars import remove_chars
                orig = df[column].astype(str)
                df[column] = orig.apply(remove_chars, mode='non_numeric', column_name=column)
                # Track changes for highlighting
                changed = orig != df[column]
                if not hasattr(df, '_styled_columns'):
                    object.__setattr__(df, '_styled_columns', {})
                df._styled_columns[column] = changed
            elif op_key == "op_remove_non_alpha":
                print(f"DEBUG: Applying remove non-alphabetic operation")
                from operations.remove_chars import remove_chars
                orig = df[column].astype(str)
                df[column] = orig.apply(remove_chars, mode='non_alphabetic', column_name=column)
                # Track changes for highlighting
                changed = orig != df[column]
                if not hasattr(df, '_styled_columns'):
                    object.__setattr__(df, '_styled_columns', {})
                df._styled_columns[column] = changed
            elif op_key == "op_find_replace":
                print(f"DEBUG: Applying find/replace operation")
                from operations.find_replace import find_replace
                find_text = operation_params.get('find_text', '')
                replace_text = operation_params.get('replace_text', '')
                
                # Apply find_replace and track changes
                orig = df[column].astype(str)
                result_series = orig.apply(find_replace, find_text=find_text, replace_text=replace_text, column_name=column)
                
                # Extract the modified values and change tracking
                df[column] = result_series.apply(lambda x: x[0] if isinstance(x, tuple) else x)
                changed = result_series.apply(lambda x: x[1] if isinstance(x, tuple) else False)
                
                # Track changes for highlighting
                if not hasattr(df, '_modified_columns'):
                    object.__setattr__(df, '_modified_columns', {})
                df._modified_columns[column] = changed
                
                # Also add to styled columns for preview highlighting
                if not hasattr(df, '_styled_columns'):
                    object.__setattr__(df, '_styled_columns', {})
                df._styled_columns[column] = changed
            elif op_key == "op_split_delimiter":
                print(f"DEBUG: Applying split by delimiter operation")
                delimiter = operation_params.get('delimiter', '')
                print(f"DEBUG: Delimiter parameter: '{delimiter}'")
                from operations.splitting import apply_split_by_delimiter
                modified_df, result = apply_split_by_delimiter(df, column, delimiter, PREVIEW_TEXTS)
                if result[0] == 'success':
                    df = modified_df
                    print(f"DEBUG: Split operation successful")
                else:
                    print(f"DEBUG: Split operation failed: {result[1]}")
                    raise Exception(result[1])
            elif op_key == "op_split_surname":
                print(f"DEBUG: Applying split surname operation")
                from operations.splitting import apply_split_surname
                modified_df, result = apply_split_surname(df, column, PREVIEW_TEXTS)
                if result[0] == 'success':
                    df = modified_df
                else:
                    raise Exception(result[1])
            elif op_key == "op_remove_specific":
                print(f"DEBUG: Applying remove specific characters operation")
                from operations.remove_chars import remove_chars
                chars_to_remove = operation_params.get('chars_to_remove', '')
                df[column] = df[column].astype(str).apply(remove_chars, mode='specific', chars_to_remove=chars_to_remove, column_name=column)
            elif op_key == "op_fill_missing":
                print(f"DEBUG: Applying fill missing operation")
                from operations.fill_missing import fill_missing
                fill_value = operation_params.get('fill_value', '')
                df[column] = df[column].apply(fill_missing, fill_value=fill_value, column_name=column)
            elif op_key == "op_extract_pattern":
                print(f"DEBUG: Applying extract pattern operation")
                from operations.extract_pattern import apply_extract_pattern
                pattern = operation_params.get('pattern', '')
                new_col_name = operation_params.get('new_col_name', '')
                df, result = apply_extract_pattern(df, column, new_col_name, pattern, PREVIEW_TEXTS)
                if result[0] != 'success':
                    raise Exception(result[1])
            elif op_key.startswith("op_validate_"):
                print(f"DEBUG: Applying validation operation")
                from operations.validate_inputs import apply_validation
                validation_type = op_key.replace("op_validate_", "")
                df, result = apply_validation(df, column, validation_type, PREVIEW_TEXTS)
                if result[0] != 'success':
                    raise Exception(result[1])
            elif op_key == "op_distinct_group":
                print(f"DEBUG: Applying distinct group operation")
                from operations.distinct_group import apply_distinct_group_encoding
                df, metadata = apply_distinct_group_encoding(df, column)
            else:
                print(f"WARNING: Unknown operation key: {op_key}")
                
        print(f"DEBUG: Operation completed successfully, returning DataFrame with shape: {df.shape}")
        return df
        
    except Exception as e:
        print(f"FATAL ERROR in apply_operation_to_partition: {e}")
        import traceback
        traceback.print_exc()
        raise Exception(f"Failed to apply operation: {e}")

def generate_preview(app, op_key, selected_col, current_preview_df, PREVIEW_ROWS, operation_params=None):
    """Generate a preview of the operation on the data."""
    try:
        if current_preview_df is None:
            return None, False, "No data available for preview"
            
        # Create a deep copy of the preview data
        preview_df = current_preview_df.copy(deep=True)
        
        # Create operation parameters if not provided
        if operation_params is None:
            operation_params = {
                'type': 'column_operation',
                'key': op_key,
                'column': selected_col
            }
        
        # Apply the operation
        preview_df = apply_operation_to_partition(preview_df, operation_params['type'], operation_params)
        
        # Check if any styling information was added
        has_styling = hasattr(preview_df, '_styled_columns') and selected_col in preview_df._styled_columns
        
        if has_styling:
            return preview_df, True, "Preview generated with highlighted changes"
        return preview_df, True, "Preview generated successfully"
    except Exception as e:
        # Return a copy of the original preview data on error
        return current_preview_df.copy(deep=True) if current_preview_df is not None else None, False, str(e)
