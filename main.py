import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

# --- Processing Functions ---
def mask_data(data):
    """Masks data by keeping the first 2 and last 2 characters."""
    s_data = str(data)
    if len(s_data) <= 4:
        return s_data # Or return "****" if you want to mask short strings too
    else:
        return s_data[:2] + '*' * (len(s_data) - 4) + s_data[-2:]

def trim_spaces(data):
    """Removes leading/trailing spaces from data."""
    return str(data).strip()

# --- GUI Application ---
class ExcelEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Table Tools")
        self.root.geometry("600x400")

        self.file_path = tk.StringVar()
        self.selected_column = tk.StringVar()
        self.selected_operation = tk.StringVar()
        self.dataframe = None

        # --- File Selection ---
        file_frame = ttk.LabelFrame(root, text="File Selection")
        file_frame.pack(padx=10, pady=10, fill="x")

        ttk.Label(file_frame, text="Excel File:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(file_frame, textvariable=self.file_path, width=50, state="readonly").grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse...", command=self.browse_file).grid(row=0, column=2, padx=5, pady=5)

        # --- Operations ---
        ops_frame = ttk.LabelFrame(root, text="Operations")
        ops_frame.pack(padx=10, pady=10, fill="x")

        ttk.Label(ops_frame, text="Column:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.column_combobox = ttk.Combobox(ops_frame, textvariable=self.selected_column, state="disabled", width=30)
        self.column_combobox.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(ops_frame, text="Operation:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.operation_combobox = ttk.Combobox(ops_frame, textvariable=self.selected_operation, state="disabled", width=30)
        self.operation_combobox['values'] = ["Mask Column (Keep 2+2)", "Trim Spaces", "Split Column (Space)", "Split Column (:)"] # Add more operations here
        self.operation_combobox.grid(row=1, column=1, padx=5, pady=5)
        # Add entry for delimiter if needed later

        ttk.Button(ops_frame, text="Apply Operation", command=self.apply_operation).grid(row=2, column=0, columnspan=2, padx=5, pady=10)

        # --- Save ---
        save_frame = ttk.Frame(root)
        save_frame.pack(padx=10, pady=10, fill="x")
        ttk.Button(save_frame, text="Save Changes", command=self.save_file).pack(side="right", padx=5)


    def browse_file(self):
        """Opens a file dialog to select an Excel file."""
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if path:
            self.file_path.set(path)
            self.load_excel()

    def load_excel(self):
        """Loads the selected Excel file into a pandas DataFrame."""
        path = self.file_path.get()
        if not path:
            return
        try:
            # Try reading with openpyxl first for .xlsx
            if path.endswith('.xlsx'):
                self.dataframe = pd.read_excel(path, engine='openpyxl')
            else: # For .xls
                 self.dataframe = pd.read_excel(path) # pandas default engine handles .xls

            # Update column combobox
            self.column_combobox['values'] = list(self.dataframe.columns)
            self.column_combobox.config(state="readonly")
            if self.dataframe.columns.any():
                 self.selected_column.set(self.dataframe.columns[0]) # Default to first column
            self.operation_combobox.config(state="readonly") # Enable operations dropdown
            messagebox.showinfo("Success", f"Loaded '{os.path.basename(path)}' successfully.")
        except Exception as e:
            messagebox.showerror("Error Loading File", f"Failed to load Excel file.\nError: {e}")
            self.file_path.set("")
            self.dataframe = None
            self.column_combobox['values'] = []
            self.column_combobox.config(state="disabled")
            self.operation_combobox.config(state="disabled")
            self.selected_column.set("")
            self.selected_operation.set("")


    def apply_operation(self):
        """Applies the selected operation to the selected column."""
        if self.dataframe is None:
            messagebox.showwarning("No File", "Please load an Excel file first.")
            return

        col = self.selected_column.get()
        op = self.selected_operation.get()

        if not col:
            messagebox.showwarning("No Column", "Please select a column.")
            return
        if not op:
            messagebox.showwarning("No Operation", "Please select an operation.")
            return

        try:
            if op == "Mask Column (Keep 2+2)":
                # Ensure column is treated as string before applying mask
                self.dataframe[col] = self.dataframe[col].astype(str).apply(mask_data)
                messagebox.showinfo("Success", f"Masked data in column '{col}'.")
            elif op == "Trim Spaces":
                 # Ensure column is treated as string before applying trim
                 self.dataframe[col] = self.dataframe[col].astype(str).apply(trim_spaces)
                 messagebox.showinfo("Success", f"Trimmed spaces in column '{col}'.")
            elif op == "Split Column (Space)":
                self.split_column_operation(col, ' ')
            elif op == "Split Column (:)":
                 self.split_column_operation(col, ':')
            # Add more operations here
            else:
                messagebox.showwarning("Not Implemented", f"Operation '{op}' is not yet implemented.")
                return

            # Refresh column list in case new columns were added (by split)
            self.column_combobox['values'] = list(self.dataframe.columns)
            # Try to keep the original column selected if it still exists, otherwise select the first
            if col in self.dataframe.columns:
                self.selected_column.set(col)
            elif self.dataframe.columns.any():
                self.selected_column.set(self.dataframe.columns[0])
            else:
                self.selected_column.set("")


        except Exception as e:
            messagebox.showerror("Operation Error", f"An error occurred during the operation:\n{e}")

    def split_column_operation(self, col, delimiter):
        """Handles the split column operation."""
        if col not in self.dataframe.columns:
             messagebox.showerror("Error", f"Column '{col}' not found.")
             return

        # Ensure the column is string type before splitting
        col_data = self.dataframe[col].astype(str)

        # Check if delimiter exists in any cell before splitting
        if not col_data.str.contains(delimiter, regex=False).any():
             messagebox.showwarning("Split Warning", f"The delimiter '{delimiter}' was not found in column '{col}'. No changes made.")
             return

        # Create new column names based on the maximum number of splits
        max_splits = col_data.str.split(delimiter).str.len().max()
        new_cols = [f"{col}_part{i+1}" for i in range(max_splits)]

        # Perform the split
        split_data = col_data.str.split(delimiter, expand=True, n=max_splits - 1) # n=max_splits-1 ensures correct number of columns
        split_data.columns = new_cols # Assign new column names

        # Find the index of the original column to insert the new columns after it
        original_col_index = self.dataframe.columns.get_loc(col)

        # Drop the original column
        df_before = self.dataframe.iloc[:, :original_col_index]
        df_after = self.dataframe.iloc[:, original_col_index+1:]

        # Concatenate the parts: before original, new split columns, after original
        self.dataframe = pd.concat([df_before, split_data, df_after], axis=1)

        messagebox.showinfo("Success", f"Split column '{col}' by '{delimiter}' into {len(new_cols)} new columns.")


    def save_file(self):
        """Saves the modified DataFrame to a new Excel file."""
        if self.dataframe is None:
            messagebox.showwarning("No Data", "No data to save. Load and modify a file first.")
            return

        original_path = self.file_path.get()
        # Suggest a new filename based on the original
        if original_path:
            dir_name = os.path.dirname(original_path)
            base_name = os.path.basename(original_path)
            name, ext = os.path.splitext(base_name)
            suggested_name = os.path.join(dir_name, f"{name}_modified{ext}")
        else:
            suggested_name = "modified_excel.xlsx" # Default if no file was loaded

        save_path = filedialog.asksaveasfilename(
            title="Save Modified Excel File",
            initialfile=suggested_name,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )

        if save_path:
            try:
                # Use openpyxl engine for .xlsx to support newer features if needed
                if save_path.endswith('.xlsx'):
                    self.dataframe.to_excel(save_path, index=False, engine='openpyxl')
                else: # For .xls
                    self.dataframe.to_excel(save_path, index=False)
                messagebox.showinfo("Success", f"File saved successfully to:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Save Error", f"Failed to save the file.\nError: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelEditorApp(root)
    root.mainloop()
