import dask.dataframe as dd
import pandas as pd
import numpy as np
from typing import List, Dict, Any, Tuple, Optional
import threading
from queue import Queue
import time
import os
import pickle

def apply_operation_to_partition(df, operation_type, operation_params):
    """Helper function to apply operation to a partition."""
    if operation_type == 'column_operation':
        column = operation_params.get('column')
        op_key = operation_params.get('key')
        
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
            df[column] = df[column].astype(str).apply(trim_spaces, column_name=column)
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
            df[column] = df[column].astype(str).apply(remove_chars, mode='non_numeric', column_name=column)
        elif op_key == "op_remove_non_alpha":
            from operations.remove_chars import remove_chars
            df[column] = df[column].astype(str).apply(remove_chars, mode='non_alphabetic', column_name=column)
        elif op_key.startswith("op_validate_"):
            from operations.validate_inputs import apply_validation
            validation_type = op_key.replace("op_validate_", "")
            df, _ = apply_validation(df, column, validation_type, None)  # None for texts as it's not needed here
            
    return df

class DelayedOperationManager:
    def __init__(self):
        self.operations = []
        self.preview_df = None
        self.full_file_path = None
        self._cancel_flag = False
        self._progress_queue = Queue()
        self.input_file_type = None

    def _get_file_type(self, file_path: str) -> str:
        """Determine file type from extension."""
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.csv':
            return 'csv'
        elif ext in ['.xls', '.xlsx']:
            return 'excel'
        else:
            raise ValueError(f"Unsupported file type: {ext}")

    def load_preview(self, file_path: str, size: str, position: str) -> pd.DataFrame:
        """Load a preview of the file based on size and position."""
        self.full_file_path = file_path
        self.input_file_type = self._get_file_type(file_path)
        
        # Determine number of rows to read
        nrows = 1000 if size == "1k" else 10000
        
        try:
            if self.input_file_type == 'csv':
                if position == "head":
                    return pd.read_csv(file_path, nrows=nrows)
                elif position == "tail":
                    full_df = pd.read_csv(file_path)
                    return full_df.tail(nrows)
                else:  # middle - random sample
                    full_df = pd.read_csv(file_path)
                    if len(full_df) <= nrows:
                        return full_df
                    middle_start = max(0, (len(full_df) - nrows) // 2)
                    return full_df.iloc[middle_start:middle_start + nrows]
            else:  # Excel files
                if position == "head":
                    return pd.read_excel(file_path, nrows=nrows)
                elif position == "tail":
                    full_df = pd.read_excel(file_path)
                    return full_df.tail(nrows)
                else:  # middle - random sample
                    full_df = pd.read_excel(file_path)
                    if len(full_df) <= nrows:
                        return full_df
                    middle_start = max(0, (len(full_df) - nrows) // 2)
                    return full_df.iloc[middle_start:middle_start + nrows]
        except Exception as e:
            raise Exception(f"Error loading file: {str(e)}")

    def add_operation(self, operation: Dict[str, Any]):
        """Add a new operation to the queue."""
        # Store only the necessary operation parameters
        if operation['type'] == 'column_operation':
            self.operations.append({
                'type': operation['type'],
                'key': operation['key'],
                'column': operation['column']
            })
        
    def clear_operations(self):
        """Clear all pending operations."""
        self.operations = []
        
    def simulate_operations(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply operations to a preview DataFrame without modifying original."""
        result = df.copy()
        for op in self.operations:
            try:
                result = apply_operation_to_partition(result, op['type'], op)
            except Exception as e:
                print(f"Error simulating operation: {e}")
        return result

    def cancel_processing(self):
        """Cancel the current processing operation."""
        self._cancel_flag = True

    def _get_dask_meta(self, df: pd.DataFrame) -> Dict:
        """Get metadata for dask operations based on pandas DataFrame."""
        return df.dtypes.to_dict()

    def save_with_operations(self, output_path: str, progress_callback=None) -> bool:
        """
        Apply all operations to the full file and save the result.
        Returns True if successful, False if cancelled.
        """
        self._cancel_flag = False
        total_ops = len(self.operations) + 2  # +2 for loading and saving
        current_op = 0

        try:
            # Load the full file
            if progress_callback:
                progress_callback(current_op / total_ops, "Loading file...")

            # First load a sample to get metadata
            sample_size = 1000
            if self.input_file_type == 'csv':
                sample_df = pd.read_csv(self.full_file_path, nrows=sample_size)
                ddf = dd.read_csv(self.full_file_path)
            else:
                sample_df = pd.read_excel(self.full_file_path, nrows=sample_size)
                full_df = pd.read_excel(self.full_file_path)
                ddf = dd.from_pandas(full_df, npartitions=max(1, len(full_df) // 100000))

            current_op += 1

            # Apply each operation
            for i, op in enumerate(self.operations):
                if self._cancel_flag:
                    return False
                    
                if progress_callback:
                    progress_callback((current_op + i) / total_ops, 
                                   f"Applying operation {i+1} of {len(self.operations)}...")
                
                try:
                    # Get metadata from sample
                    sample_result = apply_operation_to_partition(sample_df, op['type'], op)
                    meta = self._get_dask_meta(sample_result)
                    
                    # Apply operation with proper metadata
                    ddf = ddf.map_partitions(
                        apply_operation_to_partition,
                        op['type'],
                        op,
                        meta=meta
                    )
                except Exception as e:
                    print(f"Error applying operation: {e}")
                    continue

            # Save the result
            if self._cancel_flag:
                return False
                
            if progress_callback:
                progress_callback((total_ops - 1) / total_ops, "Saving file...")
            
            # Compute and save the final result
            result_df = ddf.compute()
            
            # Save using the same format as input
            if self.input_file_type == 'csv':
                result_df.to_csv(output_path, index=False)
            else:
                result_df.to_excel(output_path, index=False)
            
            if progress_callback:
                progress_callback(1.0, "Complete!")
            
            return True

        except Exception as e:
            print(f"Error during save operation: {e}")
            raise  # Re-raise the exception to see the full traceback
            return False 