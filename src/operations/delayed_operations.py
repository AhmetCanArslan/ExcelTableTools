import pandas as pd
import numpy as np
from typing import List, Dict, Any, Tuple, Optional
import threading
from queue import Queue
import time
import os
import gc
import psutil
from functools import lru_cache
import openpyxl
from concurrent.futures import ThreadPoolExecutor
import math
from .preview_utils import apply_operation_to_partition

def calculate_optimal_chunk_size(file_size: int) -> int:
    """Calculate optimal chunk size based on file size and available memory."""
    available_memory = psutil.virtual_memory().available
    base_chunk_size = min(1000000, max(1000, file_size // 100))
    return min(base_chunk_size, available_memory // 10)  # Use at most 10% of available memory

class ChunkIterator:
    """Memory-efficient iterator for processing file chunks."""
    def __init__(self, file_path: str, chunk_size: int):
        self.file_path = file_path
        self.chunk_size = chunk_size
        self.total_rows = 0
        self._count_rows()
    
    def _count_rows(self):
        """Count total rows efficiently."""
        if self.file_path.lower().endswith('.csv'):
            with open(self.file_path, 'rb') as f:
                self.total_rows = sum(1 for _ in f) - 1
        else:
            with openpyxl.load_workbook(self.file_path, read_only=True) as wb:
                self.total_rows = wb.active.max_row - 1

    def __iter__(self):
        if self.file_path.lower().endswith('.csv'):
            yield from pd.read_csv(
                self.file_path,
                chunksize=self.chunk_size,
                low_memory=False,
                dtype_backend='numpy_nullable',
                engine='c'
            )
        else:
            wb = openpyxl.load_workbook(self.file_path, read_only=True, data_only=True)
            sheet = wb.active
            headers = [cell.value for cell in sheet[1]]
            
            current_chunk = []
            for row in sheet.iter_rows(min_row=2, values_only=True):
                current_chunk.append(row)
                if len(current_chunk) >= self.chunk_size:
                    df = pd.DataFrame(current_chunk, columns=headers)
                    yield df
                    del df, current_chunk
                    current_chunk = []
                    gc.collect()
            
            if current_chunk:
                yield pd.DataFrame(current_chunk, columns=headers)
            
            wb.close()

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

    @lru_cache(maxsize=32)
    def _get_column_metadata(self, column_name: str) -> Dict:
        """Cache column metadata to avoid recomputing."""
        return {
            'dtype': self.preview_df[column_name].dtype,
            'unique_count': len(self.preview_df[column_name].unique()),
            'has_nulls': self.preview_df[column_name].isnull().any()
        }

    def _optimize_dtypes(self, df: pd.DataFrame) -> pd.DataFrame:
        """Optimize DataFrame memory usage by choosing appropriate dtypes."""
        dtype_map = {
            'float64': 'float32',
            'int64': 'int32'
        }
        
        return df.astype({col: dtype_map.get(str(dtype), dtype) 
                         for col, dtype in df.dtypes.items()})

    def load_preview(self, file_path: str, position: str) -> pd.DataFrame:
        """Load a preview of the file based on position."""
        self.full_file_path = file_path
        self.input_file_type = self._get_file_type(file_path)
        
        nrows = 1000  # Fixed preview size
        
        try:
            if self.input_file_type == 'csv':
                total_rows = sum(1 for _ in open(file_path)) - 1
                if position == "head":
                    df = pd.read_csv(file_path, nrows=nrows, low_memory=False)
                elif position == "tail":
                    if total_rows <= nrows:
                        df = pd.read_csv(file_path, low_memory=False)
                    else:
                        skiprows = range(1, total_rows - nrows + 1)
                        df = pd.read_csv(file_path, skiprows=skiprows, low_memory=False)
                else:  # middle
                    if total_rows <= nrows:
                        df = pd.read_csv(file_path, low_memory=False)
                    else:
                        skiprows = range(1, (total_rows - nrows) // 2)
                        df = pd.read_csv(file_path, skiprows=skiprows, nrows=nrows, low_memory=False)
            else:
                wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
                sheet = wb.active
                total_rows = sheet.max_row - 1  # Exclude header

                if position == "head":
                    data = list(sheet.iter_rows(min_row=1, max_row=nrows+1, values_only=True))
                elif position == "tail":
                    if total_rows <= nrows:
                        data = list(sheet.iter_rows(values_only=True))
                    else:
                        start_row = total_rows - nrows + 1
                        data = list(sheet.iter_rows(min_row=1, max_row=2, values_only=True))  # Headers
                        data.extend(sheet.iter_rows(min_row=start_row, values_only=True))
                else:  # middle
                    if total_rows <= nrows:
                        data = list(sheet.iter_rows(values_only=True))
                    else:
                        middle_start = (total_rows - nrows) // 2
                        data = list(sheet.iter_rows(min_row=1, max_row=2, values_only=True))  # Headers
                        data.extend(sheet.iter_rows(min_row=middle_start, max_row=middle_start+nrows, values_only=True))

                headers = data[0]
                df = pd.DataFrame(data[1:], columns=headers)
                wb.close()

            # Optimize memory usage
            df = self._optimize_dtypes(df)
            return df

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
        """Clear all pending operations and reset state."""
        self.operations = []
        self._cancel_flag = False
        self.full_file_path = None
        
        # Clear any cached data
        if hasattr(self, '_get_column_metadata'):
            self._get_column_metadata.cache_clear()
        
        # Force garbage collection
        gc.collect()

    def cancel_processing(self):
        """Cancel the current processing operation."""
        self._cancel_flag = True

    def _process_chunk(self, chunk: pd.DataFrame) -> pd.DataFrame:
        """Process a single chunk with all operations."""
        try:
            for op in self.operations:
                if self._cancel_flag:
                    return None
                chunk = apply_operation_to_partition(chunk, op['type'], op)
            return chunk
        finally:
            gc.collect()

    def save_with_operations(self, output_path: str, progress_callback=None) -> bool:
        """Apply all operations to the full file and save the result."""
        self._cancel_flag = False
        file_size = os.path.getsize(self.full_file_path)
        chunk_size = calculate_optimal_chunk_size(file_size)
        
        try:
            chunk_iterator = ChunkIterator(self.full_file_path, chunk_size)
            total_rows = chunk_iterator.total_rows
            processed_rows = 0
            current_excel_row = 0  # Track current row position for Excel writing

            if progress_callback:
                progress_callback(0, f"Starting file processing ({total_rows:,} total rows)...")

            # For CSV output
            if output_path.lower().endswith('.csv'):
                first_chunk = True
                for chunk_idx, chunk in enumerate(chunk_iterator):
                    if self._cancel_flag:
                        return False

                    # Process chunk
                    processed_chunk = self._process_chunk(chunk)
                    if processed_chunk is None:
                        return False

                    # Write chunk to CSV
                    processed_chunk.to_csv(
                        output_path,
                        mode='w' if first_chunk else 'a',
                        header=first_chunk,
                        index=False
                    )

                    processed_rows += len(chunk)
                    if progress_callback:
                        progress = processed_rows / total_rows
                        progress_callback(
                            progress,
                            f"Processed {processed_rows:,} of {total_rows:,} rows ({progress*100:.1f}%)..."
                        )

                    first_chunk = False
                    del chunk, processed_chunk
                    gc.collect()

            # For Excel output
            else:
                # Create Excel writer with xlsxwriter engine
                with pd.ExcelWriter(
                    output_path,
                    engine='xlsxwriter',
                    engine_kwargs={'options': {
                        'strings_to_numbers': False,
                        'constant_memory': True
                    }}
                ) as writer:
                    
                    workbook = writer.book
                    
                    # Define formats
                    default_format = workbook.add_format({
                        'num_format': '@'  # Text format for all cells
                    })
                    invalid_format = workbook.add_format({
                        'bg_color': '#FFCCCC',  # Light red background
                        'font_color': '#000000',  # Black text
                        'num_format': '@'  # Text format
                    })
                    
                    # Dictionary to track maximum width of each column
                    max_width = {}
                    
                    # Process chunks and write to Excel
                    for chunk_idx, chunk in enumerate(chunk_iterator):
                        if self._cancel_flag:
                            return False

                        # Process chunk
                        processed_chunk = self._process_chunk(chunk)
                        if processed_chunk is None:
                            return False

                        if chunk_idx == 0:
                            # Write headers on first chunk
                            worksheet = workbook.add_worksheet('Sheet1')
                            for col_idx, col_name in enumerate(processed_chunk.columns):
                                worksheet.write(0, col_idx, col_name, default_format)
                                # Initialize max width with header length
                                max_width[col_idx] = len(str(col_name)) + 2

                            current_excel_row = 1

                        # Write data rows and track column widths
                        for row_idx, row in processed_chunk.iterrows():
                            for col_idx, value in enumerate(row):
                                # Check if cell should be marked as invalid
                                is_invalid = False
                                if hasattr(processed_chunk, '_styled_columns'):
                                    col_name = processed_chunk.columns[col_idx]
                                    if col_name in processed_chunk._styled_columns:
                                        is_invalid = processed_chunk._styled_columns[col_name].iloc[row_idx]
                                
                                # Convert value to string and get its length
                                str_value = '' if pd.isna(value) else str(value)
                                # Update maximum width if necessary
                                max_width[col_idx] = max(max_width[col_idx], len(str_value) + 2)
                                
                                # Write cell with appropriate format
                                cell_format = invalid_format if is_invalid else default_format
                                worksheet.write(
                                    current_excel_row,
                                    col_idx,
                                    str_value,
                                    cell_format
                                )
                            current_excel_row += 1

                        processed_rows += len(chunk)
                        if progress_callback:
                            progress = processed_rows / total_rows
                            progress_callback(
                                progress,
                                f"Processed {processed_rows:,} of {total_rows:,} rows ({progress*100:.1f}%)..."
                            )

                        del chunk, processed_chunk
                        gc.collect()

                    # Set column widths based on content
                    for col_idx, width in max_width.items():
                        # Cap maximum width at 100 characters to prevent extremely wide columns
                        worksheet.set_column(col_idx, col_idx, min(width, 100), default_format)

            if progress_callback:
                progress_callback(1.0, f"Complete! Processed {total_rows:,} rows.")

            return True

        except Exception as e:
            print(f"Error during save operation: {e}")
            raise 