import dask.dataframe as dd
import pandas as pd
import numpy as np
from typing import List, Dict, Any, Tuple, Optional
import threading
from queue import Queue
import time
import os
import pickle
from concurrent.futures import ThreadPoolExecutor
import math
import multiprocessing
from .preview_utils import apply_operation_to_partition

def get_optimal_workers():
    """Calculate optimal number of workers based on CPU cores."""
    cpu_count = multiprocessing.cpu_count()
    if cpu_count <= 4:
        return cpu_count
    elif cpu_count <= 8:
        return min(cpu_count + 4, 12)
    else:
        return min(cpu_count + 4, 16)

class DelayedOperationManager:
    def __init__(self):
        self.operations = []
        self.preview_df = None
        self.full_file_path = None
        self._cancel_flag = False
        self._progress_queue = Queue()
        self.input_file_type = None
        self.chunk_size = 1000000
        self.max_workers = get_optimal_workers()  # Adaptive worker count

    def _get_file_type(self, file_path: str) -> str:
        """Determine file type from extension."""
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.csv':
            return 'csv'
        elif ext in ['.xls', '.xlsx']:
            return 'excel'
        else:
            raise ValueError(f"Unsupported file type: {ext}")

    def _get_total_rows(self, file_path: str) -> int:
        """Get total number of rows in the file."""
        if self.input_file_type == 'csv':
            # Count lines in CSV file efficiently
            with open(file_path, 'rb') as f:
                return sum(1 for _ in f) - 1  # Subtract 1 for header
        else:
            # For Excel, use pandas to get info
            return len(pd.read_excel(file_path, nrows=None))

    def load_preview(self, file_path: str, position: str) -> pd.DataFrame:
        """Load a preview of the file based on position."""
        self.full_file_path = file_path
        self.input_file_type = self._get_file_type(file_path)
        
        # Always use 1000 rows for preview
        nrows = 1000
        
        try:
            if self.input_file_type == 'csv':
                total_rows = self._get_total_rows(file_path)
                if position == "head":
                    return pd.read_csv(file_path, nrows=nrows)
                elif position == "tail":
                    if total_rows <= nrows:
                        return pd.read_csv(file_path)
                    skiprows = max(0, total_rows - nrows)
                    return pd.read_csv(file_path, skiprows=skiprows)
                else:  # middle
                    if total_rows <= nrows:
                        return pd.read_csv(file_path)
                    skiprows = max(0, (total_rows - nrows) // 2)
                    return pd.read_csv(file_path, skiprows=skiprows, nrows=nrows)
            else:  # Excel files
                if position == "head":
                    return pd.read_excel(file_path, nrows=nrows)
                elif position == "tail":
                    full_df = pd.read_excel(file_path)
                    return full_df.tail(nrows)
                else:  # middle
                    full_df = pd.read_excel(file_path)
                    if len(full_df) <= nrows:
                        return full_df
                    middle_start = max(0, (len(full_df) - nrows) // 2)
                    return full_df.iloc[middle_start:middle_start + nrows]
        except Exception as e:
            raise Exception(f"Error loading file: {str(e)}")

    def add_operation(self, operation: Dict[str, Any]):
        """Add a new operation to the queue."""
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

    def _process_chunk(self, chunk_df: pd.DataFrame, operations: List[Dict]) -> pd.DataFrame:
        """Process a single chunk of data with all operations."""
        for op in operations:
            if self._cancel_flag:
                return None
            chunk_df = apply_operation_to_partition(chunk_df, op['type'], op)
        return chunk_df

    def _save_chunk(self, chunk_df: pd.DataFrame, output_path: str, mode: str):
        """Save a chunk to file with appropriate mode."""
        if self.input_file_type == 'csv':
            chunk_df.to_csv(output_path, mode=mode, header=(mode == 'w'), index=False)
        else:
            # For Excel, we'll collect chunks and save at once
            return chunk_df

    def save_with_operations(self, output_path: str, progress_callback=None) -> bool:
        """Apply all operations to the full file and save the result."""
        self._cancel_flag = False
        
        try:
            # Calculate total steps for accurate progress
            total_rows = self._get_total_rows(self.full_file_path)
            total_chunks = math.ceil(total_rows / self.chunk_size)
            total_steps = total_chunks + 1  # +1 for final saving
            current_step = 0

            if progress_callback:
                progress_callback(0, "Starting file processing...")

            # Process the file in chunks
            if self.input_file_type == 'csv':
                # For CSV, we'll process and save chunks directly
                chunks = pd.read_csv(self.full_file_path, chunksize=self.chunk_size)
                first_chunk = True

                with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                    futures = []
                    
                    for chunk_idx, chunk in enumerate(chunks):
                        if self._cancel_flag:
                            return False

                        # Process chunk
                        future = executor.submit(self._process_chunk, chunk, self.operations)
                        futures.append((future, chunk_idx))
                        
                        # Update progress more frequently
                        if progress_callback:
                            current_step += 1
                            progress = (current_step - 1) / total_steps
                            progress_callback(progress, f"Processing chunk {chunk_idx + 1} of {total_chunks}...")

                        # Save chunks as they complete to avoid memory buildup
                        completed_futures = []
                        for future, idx in futures:
                            if future.done():
                                chunk_result = future.result()
                                if chunk_result is not None:
                                    mode = 'w' if first_chunk else 'a'
                                    self._save_chunk(chunk_result, output_path, mode)
                                    first_chunk = False
                                completed_futures.append((future, idx))
                                
                        # Remove completed futures
                        futures = [f for f in futures if f not in completed_futures]

                    # Process any remaining futures
                    for future, idx in futures:
                        if self._cancel_flag:
                            return False
                        chunk_result = future.result()
                        if chunk_result is not None:
                            mode = 'w' if first_chunk else 'a'
                            self._save_chunk(chunk_result, output_path, mode)
                            first_chunk = False

            else:
                # For Excel files, we need to collect all chunks and save at once
                # but we'll process them in smaller batches
                full_df = pd.read_excel(self.full_file_path)
                chunks = [full_df[i:i + self.chunk_size] for i in range(0, len(full_df), self.chunk_size)]
                del full_df  # Free up memory
                
                processed_chunks = []
                
                with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                    futures = []
                    
                    for chunk_idx, chunk in enumerate(chunks):
                        if self._cancel_flag:
                            return False
                        
                        future = executor.submit(self._process_chunk, chunk, self.operations)
                        futures.append(future)
                        
                        if progress_callback:
                            current_step += 1
                            progress = (current_step - 1) / total_steps
                            progress_callback(progress, f"Processing chunk {chunk_idx + 1} of {total_chunks}...")
                        
                        # Process completed chunks immediately
                        completed_futures = [f for f in futures if f.done()]
                        for completed_future in completed_futures:
                            chunk_result = completed_future.result()
                            if chunk_result is not None:
                                processed_chunks.append(chunk_result)
                        futures = [f for f in futures if f not in completed_futures]
                        
                    # Process any remaining futures
                    for future in futures:
                        if self._cancel_flag:
                            return False
                        chunk_result = future.result()
                        if chunk_result is not None:
                            processed_chunks.append(chunk_result)

                if self._cancel_flag:
                    return False

                # Combine chunks and save
                if progress_callback:
                    progress_callback(0.95, "Saving file...")

                # Combine chunks efficiently
                final_df = pd.concat(processed_chunks, ignore_index=True, copy=False)
                final_df.to_excel(output_path, index=False)
                
                # Clear memory
                del processed_chunks
                del final_df

            if progress_callback:
                progress_callback(1.0, "Complete!")

            return True

        except Exception as e:
            print(f"Error during save operation: {e}")
            raise 