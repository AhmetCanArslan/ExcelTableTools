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

def generate_preview(app, op_key, selected_col, current_preview_df, PREVIEW_ROWS):
    """
    Returns (modified_df, success:bool, status_message:str).
    Centralizes all preview logic, including input dialogs.
    """
    texts = app.texts
    root = app.root
    df = current_preview_df.copy()

    try:
        if op_key == "op_mask":
            df[selected_col] = df[selected_col].astype(str).apply(mask_data, column_name=selected_col)
        elif op_key == "op_mask_email":
            df[selected_col] = df[selected_col].astype(str).apply(mask_data, mode='email', column_name=selected_col)
        elif op_key == "op_mask_words":
            df[selected_col] = df[selected_col].astype(str).apply(mask_words, column_name=selected_col)
        elif op_key == "op_trim":
            df[selected_col] = df[selected_col].astype(str).apply(trim_spaces, column_name=selected_col)
        elif op_key in ("op_upper","op_lower","op_title"):
            m = {"op_upper":"upper","op_lower":"lower","op_title":"title"}
            df[selected_col] = df[selected_col].astype(str).apply(change_case, case_type=m[op_key], column_name=selected_col)
        elif op_key == "op_find_replace":
            ft = simpledialog.askstring(texts['input_needed'], texts['enter_find_text']+" (preview)", parent=root)
            rt = simpledialog.askstring(texts['input_needed'], texts['enter_replace_text']+" (preview)", parent=root) if ft else None
            if ft is None or rt is None:
                return df, False, "Find/replace cancelled"
            df[selected_col] = df[selected_col].astype(str).apply(find_replace, find_text=ft, replace_text=rt, column_name=selected_col)
        elif op_key == "op_remove_specific":
            chars = simpledialog.askstring(texts['input_needed'], texts['enter_chars_to_remove']+" (preview)", parent=root)
            if chars is None:
                return df, False, "Cancel remove-specific"
            df[selected_col] = df[selected_col].astype(str).apply(remove_chars, mode='specific', chars_to_remove=chars, column_name=selected_col)
        elif op_key == "op_remove_non_numeric":
            df[selected_col] = df[selected_col].astype(str).apply(remove_chars, mode='non_numeric', column_name=selected_col)
        elif op_key == "op_remove_non_alpha":
            df[selected_col] = df[selected_col].astype(str).apply(remove_chars, mode='non_alphabetic', column_name=selected_col)
        elif op_key == "op_fill_missing":
            fv = simpledialog.askstring(texts['input_needed'], texts['enter_fill_value']+" (preview)", parent=root)
            if fv is None:
                return df, False, "Cancel fill-missing"
            df[selected_col] = df[selected_col].apply(fill_missing, fill_value=fv, column_name=selected_col)
        elif op_key == "op_split_space":
            df, (st, msg) = apply_split_by_delimiter(df, selected_col, " ", texts)
            if st!="success": return df, st=="success", msg
        elif op_key == "op_split_colon":
            df, (st, msg) = apply_split_by_delimiter(df, selected_col, ":", texts)
            if st!="success": return df, st=="success", msg
        elif op_key == "op_split_surname":
            df, (st, msg) = apply_split_surname(df, selected_col, texts)
            if st!="success": return df, st=="success", msg
        elif op_key == "op_extract_pattern":
            pat = simpledialog.askstring(texts['input_needed'], texts['enter_regex_pattern']+" (preview)", parent=root)
            if pat is None:
                return df, False, "Cancel regex"
            try: re.compile(pat)
            except re.error as e:
                return df, False, texts['regex_error'].format(error=e)
            name = app.get_unique_col_name(f"{selected_col}_ext_prev", df.columns)
            df, (st, msg) = apply_extract_pattern(df, selected_col, name, pat, texts)
            if st!="success": return df, st=="success", msg
        elif op_key == "op_mark_duplicates":
            name = app.get_unique_col_name(f"{selected_col}_dup_prev", df.columns)
            df, (st, msg) = apply_mark_duplicates(df, selected_col, name, texts)
            if st!="success": return df, st=="success", msg
        elif op_key == "op_remove_duplicates":
            df, (st, msg) = apply_remove_duplicates(df, selected_col, texts)
            if st!="success": return df, st=="success", msg
        elif op_key == "op_concatenate":
            cols = app.get_multiple_columns('input_needed','select_columns_concat')
            if not cols or len(cols)<2:
                return df, False, "Select ≥2 cols"
            sep = simpledialog.askstring(texts['input_needed'], texts['enter_separator']+" (preview)", parent=root)
            if sep is None:
                return df, False, "Cancel concat"
            name = app.get_unique_col_name("_".join(cols)+"_c_prev", df.columns)
            df, (st, msg) = apply_concatenate(df, cols, name, sep, texts)
            if st!="success": return df, st=="success", msg
        elif op_key == "op_merge_columns":
            cols = app.get_multiple_columns('input_needed','select_columns_merge')
            if not cols or len(cols)<2:
                return df, False, "Select ≥2 cols"
            sep = simpledialog.askstring(texts['input_needed'], texts['enter_separator']+" (preview)", parent=root)
            if sep is None:
                return df, False, "Cancel merge"
            fl = messagebox.askyesno(texts['input_needed'], texts['fill_missing_merge']+" (preview)", parent=root)
            name = app.get_unique_col_name("_".join(cols)+"_m_prev", df.columns)
            df, (st, msg) = apply_merge_columns(df, cols, name, sep, fl, texts)
            if st!="success": return df, st=="success", msg
        elif op_key == "op_rename_column":
            new_name = simpledialog.askstring(texts['input_needed'],
                                              texts['enter_new_col_name']+" (preview)",
                                              initialvalue=selected_col,
                                              parent=root)
            if new_name is None:
                return df, False, "Rename cancelled"
            df2, (st, msg) = apply_rename_column(df, selected_col, new_name, texts)
            if st!="success":
                return df2, st=="success", msg
            return df2, True, msg
        elif op_key == "op_round_numbers":
            try:
                decimals_str = simpledialog.askstring(texts['input_needed'], texts['enter_decimal_places'] + " (preview)", parent=root)
                if decimals_str is None: return df, False, "Rounding cancelled"
                decimals = int(decimals_str)
                df, (st, msg) = apply_round_numbers(df, selected_col, decimals, texts)
                if st != "success": return df, False, msg
            except ValueError:
                return df, False, texts['invalid_input_numeric']
        elif op_key == "op_calculate_column_constant":
            op_type = simpledialog.askstring(texts['input_needed'], texts['select_calculation_operation'] + " (+, -, *, /) (preview)", parent=root)
            if op_type not in ['+', '-', '*', '/']: return df, False, "Invalid operation type"
            try:
                constant_str = simpledialog.askstring(texts['input_needed'], texts['enter_constant_value'] + " (preview)", parent=root)
                if constant_str is None: return df, False, "Calculation cancelled"
                constant = float(constant_str)
                df, (st, msg) = apply_calculate_column_constant(df, selected_col, op_type, constant, texts)
                if st != "success": return df, False, msg
            except ValueError:
                return df, False, texts['invalid_input_numeric']
        elif op_key == "op_create_calculated_column":
            # For preview, this is complex as it needs two columns.
            # We might simplify or show a general message.
            # For now, let's assume selected_col is the first, and prompt for a second (mocked for preview)
            # This operation is better handled by the main apply_operation for full functionality.
            # A true preview would need a more complex dialog.
            # Let's show a message that preview is limited.
            messagebox.showinfo(texts['info'], texts['preview_not_available_complex'], parent=root)
            return df, False, texts['preview_not_available_complex']
        elif op_key == "op_check_valid_inputs":
            # Create a simple dialog for validation type selection for preview
            validation_types = {
                'email': texts.get('validation_email', "Email addresses"),
                'phone': texts.get('validation_phone', "Phone numbers"),
                'date': texts.get('validation_date', "Date format"),
                'numeric': texts.get('validation_numeric', "Numeric values"),
                'alphanumeric': texts.get('validation_alphanumeric', "Alphanumeric text"),
                'url': texts.get('validation_url', "URL addresses")
            }
            # For preview, just use a simple dialog instead of a complex UI
            validation_type_keys = list(validation_types.keys())
            validation_type_values = list(validation_types.values())
            val_type_idx = simpledialog.askinteger(
                texts['input_needed'],
                texts['select_validation_type'] + " (preview)\n\n" + 
                "\n".join([f"{i+1}. {val}" for i, val in enumerate(validation_type_values)]),
                parent=root,
                minvalue=1,
                maxvalue=len(validation_types)
            )
            if val_type_idx is None:
                return df, False, "Validation cancelled"
            
            validation_type = validation_type_keys[val_type_idx-1]
            result = apply_validation(df, selected_col, validation_type, texts)
            df, status = result[0], result[1]
            
            if status[0] != "success": 
                return df, status[0]=="success", status[1]
                
            # Return success but don't show the validation results window in preview mode
            return df, True, f"Preview: {status[1]}"
            
        else:
            return df, False, texts['not_implemented'].format(op=texts.get(op_key,op_key))

        return df, True, ""
    except Exception as e:
        return current_preview_df, False, str(e)
