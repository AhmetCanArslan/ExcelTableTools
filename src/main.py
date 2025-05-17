import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, scrolledtext
import pandas as pd
import os
import re
import json
import sys

# Add the project root to the Python path to allow imports from anywhere
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
if project_root not in sys.path:
    sys.path.insert(0, project_root)

# Import operations
from src.operations.masking import mask_data, mask_email, mask_words  
from src.operations.trimming import trim_spaces
from src.operations.splitting import apply_split_surname, apply_split_by_delimiter
from src.operations.case_change import change_case
from src.operations.find_replace import find_replace
from src.operations.remove_chars import remove_chars
from src.operations.concatenate import apply_concatenate
from src.operations.extract_pattern import apply_extract_pattern
from src.operations.fill_missing import fill_missing
from src.operations.duplicates import apply_mark_duplicates, apply_remove_duplicates
from src.operations.merge_columns import apply_merge_columns
from src.operations.rename_column import apply_rename_column   
from src.operations.preview_utils import generate_preview
from src.operations.numeric_operations import apply_round_numbers
from src.operations import numeric_operations
from src.operations.validate_inputs import apply_validation

# Import translations
from src.translations import LANGUAGES

# Constants
PREVIEW_ROWS = 1000  # Number of rows to show in preview
RESOURCES_DIR = os.path.join(project_root, 'resources')

# --- GUI Application ---
class ExcelEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.geometry("800x550")  
        self.file_path = tk.StringVar()
        self.selected_column = tk.StringVar()
        self.selected_operation = tk.StringVar()
        self.dataframe = None

        # --- language selection & persistence ---
        self.available_languages = list(LANGUAGES.keys())
        self.current_lang = tk.StringVar(value=self.load_last_language())
        self.texts = LANGUAGES[self.current_lang.get()]

        # --- Undo/Redo History ---
        self.undo_stack = []
        self.redo_stack = []

        # Remember last browsing directory - now uses persistent storage
        self.last_dir = self.load_last_directory()

        # load operations configuration instead of hard‐coding
        config_path = os.path.join(RESOURCES_DIR, 'operations_config.json')
        with open(config_path, "r") as f:
            self.ops_config = json.load(f)
        self.operation_keys = self.ops_config["operations"]

        # --- Main Content Frame ---
        main_content_frame = ttk.Frame(root)
        main_content_frame.pack(fill="both", expand=True, side=tk.TOP)

        # --- Top Frame for Language Selector ---
        top_frame = ttk.Frame(main_content_frame)
        top_frame.pack(fill="x", padx=10, pady=(5, 0))

        self.refresh_button = ttk.Button(top_frame, text=self.texts['refresh'], command=self.refresh_app)
        self.refresh_button.pack(side="left")

        self.lang_label = ttk.Label(top_frame, text=self.texts['language'] + ":")
        self.lang_label.pack(side="left", padx=(10,5))
        self.lang_combobox = ttk.Combobox(
            top_frame,
            textvariable=self.current_lang,
            values=self.available_languages,
            state="readonly",
            width=5
        )
        self.lang_combobox.pack(side="left")
        self.lang_combobox.bind("<<ComboboxSelected>>", lambda e: self.change_language())

        # --- File Selection ---
        self.file_frame = ttk.LabelFrame(main_content_frame, text=self.texts['file_selection'])
        self.file_frame.pack(padx=10, pady=10, fill="x")

        self.file_label = ttk.Label(self.file_frame, text=self.texts['excel_file'])
        self.file_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.file_entry = ttk.Entry(self.file_frame, textvariable=self.file_path, width=70, state="readonly")
        self.file_entry.grid(row=0, column=1, padx=5, pady=5)
        self.browse_button = ttk.Button(self.file_frame, text=self.texts['browse'], command=self.browse_file)
        self.browse_button.grid(row=0, column=2, padx=5, pady=5)

        # --- Operations ---
        self.ops_frame = ttk.LabelFrame(main_content_frame, text=self.texts['operations'])
        self.ops_frame.pack(padx=10, pady=10, fill="x")

        self.column_label = ttk.Label(self.ops_frame, text=self.texts['column'])
        self.column_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.column_combobox = ttk.Combobox(self.ops_frame, textvariable=self.selected_column, state="disabled", width=85)
        self.column_combobox.grid(row=0, column=1, padx=5, pady=5)

        self.operation_label = ttk.Label(self.ops_frame, text=self.texts['operation'])
        self.operation_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.operation_combobox = ttk.Combobox(self.ops_frame, textvariable=self.selected_operation, state="disabled", width=85)
        self.operation_combobox.grid(row=1, column=1, padx=5, pady=5)

        # Create a frame to hold buttons
        buttons_frame = ttk.Frame(self.ops_frame)
        buttons_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=10, sticky="ew")
        
        # Split the frame into a 3-column grid
        buttons_frame.columnconfigure(0, weight=1)
        buttons_frame.columnconfigure(1, weight=1)
        buttons_frame.columnconfigure(2, weight=1)

        self.apply_button = ttk.Button(buttons_frame, text=self.texts['apply_operation'], command=self.apply_operation)
        self.apply_button.grid(row=0, column=0, padx=2, sticky="ew")

        self.preview_button = ttk.Button(buttons_frame,
            text=self.texts.get('operation_preview_button', "Operation Preview"),
            command=self.preview_operation,
            state="disabled")  # start disabled
        self.preview_button.grid(row=0, column=1, padx=2, sticky="ew")
        
        self.output_preview_button = ttk.Button(buttons_frame,
            text=self.texts.get('output_preview_button', "Output File Preview"),
            command=self.preview_output_file,
            state="disabled")
        self.output_preview_button.grid(row=0, column=2, padx=2, sticky="ew")

        # toggle preview buttons when operation selection changes
        self.selected_operation.trace_add("write", self._on_operation_change)
        
        # When a file is loaded, enable output preview
        self.file_path.trace_add("write", self._on_file_loaded)

        self.ops_frame.columnconfigure(0, weight=1)
        self.ops_frame.columnconfigure(1, weight=1)

        # --- Save and Undo/Redo Frame ---
        save_frame = ttk.Frame(main_content_frame)
        save_frame.pack(padx=10, pady=10, fill="x")

        self.undo_button = ttk.Button(save_frame, text="Undo", command=self.undo_action, state="disabled")
        self.undo_button.pack(side="left", padx=5)

        self.redo_button = ttk.Button(save_frame, text="Redo", command=self.redo_action, state="disabled")
        self.redo_button.pack(side="left", padx=5)

        # Add a dropdown to choose output file extension
        self.output_extension = tk.StringVar()
        self.output_formats = ["xls", "xlsx", "csv", "json", "html", "md"]
        self.extension_dropdown = ttk.Combobox(save_frame, textvariable=self.output_extension,
                                               values=self.output_formats, state="readonly", width=5)
        self.extension_dropdown.pack(side="right", padx=5)

        self.save_button = ttk.Button(save_frame, text=self.texts['save_changes'], command=self.save_file)
        self.save_button.pack(side="right", padx=5)

        # Set default to 'xlsx' initially
        self.output_extension.set("xlsx")

        # --- Status Area (CLI-like) ---
        self.status_frame = ttk.LabelFrame(root, text=self.texts['status_log'])
        self.status_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(5, 10))

        self.status_text = scrolledtext.ScrolledText(self.status_frame, height=5, wrap=tk.WORD, state='disabled')
        self.status_text.pack(fill="both", expand=True, padx=5, pady=5)

        self.update_status("Ready.")

        self.update_ui_language()

    def get_unique_col_name(self, base_name, existing_columns):
        """Generates a unique column name based on existing ones."""
        new_name = base_name
        counter = 1
        while new_name in existing_columns:
            new_name = f"{base_name}_{counter}"
            counter += 1
        return new_name

    def update_status(self, message):
        """Appends a message to the status text area."""
        self.status_text.config(state='normal')
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state='disabled')

    def update_ui_language(self):
        self.texts = LANGUAGES[self.current_lang.get()]
        self.root.title(self.texts['title'])

        self.file_frame.config(text=self.texts['file_selection'])
        self.ops_frame.config(text=self.texts['operations'])

        self.file_label.config(text=self.texts['excel_file'])
        self.browse_button.config(text=self.texts['browse'])
        self.column_label.config(text=self.texts['column'])
        self.operation_label.config(text=self.texts['operation'])
        self.apply_button.config(text=self.texts['apply_operation'])
        self.preview_button.config(text=self.texts.get('operation_preview_button', "Operation Preview"))
        self.output_preview_button.config(text=self.texts.get('output_preview_button', "Output File Preview"))
        self.save_button.config(text=self.texts['save_changes'])
        self.refresh_button.config(text=self.texts['refresh'])
        self.lang_label.config(text=self.texts['language'] + ":")

        # Update Undo/Redo button texts
        self.undo_button.config(text=self.texts.get('undo', "Undo"))
        self.redo_button.config(text=self.texts.get('redo', "Redo"))

        # Update the Status Log frame text directly using the stored reference
        self.status_frame.config(text=self.texts['status_log'])

        translated_ops = [self.texts[key] for key in self.operation_keys]
        current_selection_text = self.selected_operation.get()
        self.operation_combobox['values'] = translated_ops

        if current_selection_text:
            try:
                current_key = None
                old_lang = 'tr' if self.current_lang.get() == 'en' else 'en'
                for key, text in LANGUAGES[old_lang].items():
                    if text == current_selection_text and key in self.operation_keys:
                        current_key = key
                        break
                if current_key:
                    self.selected_operation.set(self.texts[current_key])
                else:
                    self.selected_operation.set("")
            except Exception:
                self.selected_operation.set("")
        else:
            self.selected_operation.set("")

    def load_last_language(self):
        """Load last selected language or default to 'en'."""
        try:
            lang_file_path = os.path.join(RESOURCES_DIR, "last_language.txt")
            if os.path.exists(lang_file_path):
                with open(lang_file_path, "r") as f:
                    lang = f.read().strip()
                    return lang if lang in LANGUAGES else 'en'
            return 'en'
        except Exception:
            return 'en'

    def save_last_language(self):
        """Save current language to file."""
        try:
            lang_file_path = os.path.join(RESOURCES_DIR, "last_language.txt")
            with open(lang_file_path, "w") as f:
                f.write(self.current_lang.get())
        except Exception:
            pass
            
    def load_last_directory(self):
        """Load last browsing directory or default to current directory."""
        try:
            dir_file_path = os.path.join(RESOURCES_DIR, "last_directory.txt")
            if os.path.exists(dir_file_path):
                with open(dir_file_path, "r") as f:
                    directory = f.read().strip()
                    return directory if os.path.isdir(directory) else os.getcwd()
            return os.getcwd()
        except Exception:
            return os.getcwd()
            
    def save_last_directory(self):
        """Save current browsing directory to file."""
        try:
            dir_file_path = os.path.join(RESOURCES_DIR, "last_directory.txt")
            with open(dir_file_path, "w") as f:
                f.write(self.last_dir)
        except Exception:
            pass

    def change_language(self):
        """Apply and persist the selected language."""
        lang = self.current_lang.get()
        self.save_last_language()
        self.texts = LANGUAGES[lang]
        self.update_ui_language()

    def refresh_app(self):
        """Resets the application to its initial state."""
        # Clear file and data
        self.file_path.set("")
        self.dataframe = None

        # Disable & clear comboboxes
        self.column_combobox.set("")
        self.column_combobox['values'] = []
        self.column_combobox.config(state="disabled")
        self.operation_combobox.set("")
        self.operation_combobox['values'] = []
        self.operation_combobox.config(state="disabled")
        self.operation_combobox['values'] = [self.texts[key] for key in self.operation_keys]

        # Clear undo/redo history
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.update_undo_redo_buttons()

        # Clear status log
        self.status_text.config(state='normal')
        self.status_text.delete('1.0', tk.END)
        self.status_text.config(state='disabled')

        # Reset output extension dropdown
        self.output_extension.set("xlsx")

        # disable preview until user selects operation again
        self.preview_button.config(state="disabled")
        self.output_preview_button.config(state="disabled")

        # Inform user
        self.update_status("Application refreshed.")

    def get_operation_key(self, translated_op_text):
        for key in self.operation_keys:
            if self.texts[key] == translated_op_text:
                return key
        return None

    def _on_operation_change(self, *args):
        """Enable Preview only if a valid operation is selected."""
        op_text = self.selected_operation.get()
        if self.get_operation_key(op_text):
            self.preview_button.config(state="normal")
        else:
            self.preview_button.config(state="disabled")
            
    def _on_file_loaded(self, *args):
        """Enable output preview button when a file is loaded."""
        if self.file_path.get():
            self.output_preview_button.config(state="normal")
        else:
            self.output_preview_button.config(state="disabled")

    def browse_file(self):
        path = filedialog.askopenfilename(
            initialdir=self.last_dir,
            title=self.texts['select_excel_file'],
            filetypes=[(self.texts['excel_files'], "*.xlsx *.xls *.csv")]
        )
        if path:
            self.last_dir = os.path.dirname(path)
            self.save_last_directory()  # Save the directory for future sessions
            self.file_path.set(path)
            self.load_excel()
        else:
            self.update_status("File selection cancelled.")

    def load_excel(self):
        path = self.file_path.get()
        if not path:
            return
        try:
            if path.lower().endswith('.csv'):
                self.dataframe = pd.read_csv(path, low_memory=False)
            elif path.lower().endswith('.xlsx'):
                self.dataframe = pd.read_excel(path, engine='openpyxl')
            else:
                self.dataframe = pd.read_excel(path)

            # Reset history after loading
            self.undo_stack = []
            self.redo_stack = []

            # Commit the loaded DataFrame as a baseline state
            self._commit_undoable_action(self.dataframe.copy(deep=True))
            # Store original for later comparison in preview
            self.original_df = self.dataframe.copy(deep=True)

            self.column_combobox['values'] = list(self.dataframe.columns)
            self.column_combobox.config(state="readonly")
            if self.dataframe.columns.any():
                self.selected_column.set(self.dataframe.columns[0])
            self.operation_combobox.config(state="readonly")
            messagebox.showinfo(self.texts['success'], self.texts['loaded_successfully'].format(filename=os.path.basename(path)))
            self.update_status(f"Loaded '{os.path.basename(path)}'. Rows: {len(self.dataframe)}")

            # After loading the file successfully, set the extension dropdown
            ext = os.path.splitext(path)[1].lower()  # e.g. '.xlsx', '.csv'
            if ext in [".xls", ".xlsx", ".csv"]:
                self.output_extension.set(ext.lstrip('.'))

        except Exception as e:
            messagebox.showerror(self.texts['error'], self.texts['error_loading'].format(error=e))
            self.file_path.set("")
            self.dataframe = None
            self.column_combobox['values'] = []
            self.column_combobox.config(state="disabled")
            self.operation_combobox.config(state="disabled")
            self.selected_column.set("")
            self.selected_operation.set("")
            self.update_status(f"Error loading file: {e}")

    def get_input(self, title_key, prompt_key):
        return simpledialog.askstring(self.texts[title_key], self.texts[prompt_key], parent=self.root)

    def get_new_column_name(self, base_suggestion):
        while True:
            name = simpledialog.askstring(self.texts['input_needed'],
                                          self.texts['enter_new_col_name'],
                                          initialvalue=base_suggestion,
                                          parent=self.root)
            if name is None:
                return None
            name = name.strip()
            if not name:
                messagebox.showwarning(self.texts['warning'], self.texts['invalid_column_name'], parent=self.root)
                continue
            if name in self.dataframe.columns:
                messagebox.showwarning(self.texts['warning'], self.texts['column_already_exists'].format(name=name), parent=self.root)
                continue
            return name

    def get_multiple_columns(self, title_key, prompt_key):
        if self.dataframe is None or self.dataframe.columns.empty:
            return None

        dialog = tk.Toplevel(self.root)
        dialog.title(self.texts[title_key])
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.geometry("300x300")

        ttk.Label(dialog, text=self.texts[prompt_key]).pack(pady=5)

        listbox_frame = ttk.Frame(dialog)
        listbox_frame.pack(expand=True, fill="both", padx=10, pady=5)

        listbox = tk.Listbox(listbox_frame, selectmode="extended", exportselection=False)
        listbox.pack(side="left", expand=True, fill="both")

        scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=listbox.yview)
        scrollbar.pack(side="right", fill="y")
        listbox.config(yscrollcommand=scrollbar.set)

        for col_name in self.dataframe.columns:
            listbox.insert(tk.END, col_name)

        selected_columns = []

        def on_ok():
            nonlocal selected_columns
            selected_indices = listbox.curselection()
            selected_columns = [listbox.get(i) for i in selected_indices]
            if not selected_columns:
                messagebox.showwarning(self.texts['warning'], self.texts['no_columns_selected'], parent=dialog)
                return
            dialog.destroy()

        def on_cancel():
            nonlocal selected_columns
            selected_columns = None
            dialog.destroy()

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="OK", command=on_ok).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side="left", padx=5)

        self.root.wait_window(dialog)
        return selected_columns

    def show_preview_dialog(self, original_df_sample, modified_df_sample, op_text):
        preview_dialog = tk.Toplevel(self.root)
        preview_dialog.title(self.texts['preview_display_title'])
        preview_dialog.transient(self.root)
        preview_dialog.grab_set()

        width = 1000
        height = 675
        preview_dialog.geometry(f"{width}x{height}")
        preview_dialog.resizable(True, True)    

        main_frame = ttk.Frame(preview_dialog, padding="10")
        main_frame.pack(expand=True, fill="both")

        ttk.Label(main_frame, text=f"{op_text} - {self.texts['preview_display_title']}").pack(pady=5)

        notebook = ttk.Notebook(main_frame)
        notebook.pack(expand=True, fill="both", pady=5)

        # Format data as a well-aligned table string
        def format_dataframe_as_table(df):
            table_string = ""
            
            # Get maximum width for each column for proper alignment
            col_widths = {'row_num': 5}  # Start with row numbers column
            for col in df.columns:
                # Get max width of column name and values
                col_values = df[col].astype(str)
                max_value_width = max((col_values.str.len().max(), len(str(col))))
                col_widths[col] = max_value_width + 3  # Add padding
            
            # Create header
            header_row = f"{'#':<5}"  # Row number header
            separator_row = "-" * 5 + " "
            for col in df.columns:
                width = col_widths[col]
                header_row += f"{str(col):<{width}}"
                separator_row += "-" * width + " "
            
            table_string += header_row + "\n"
            table_string += separator_row + "\n"
            
            # Create data rows
            for idx, (_, row) in enumerate(df.iterrows(), 1):
                data_row = f"{idx:<5}"  # Add row number
                for col in df.columns:
                    width = col_widths[col]
                    value = str(row[col]) if not pd.isna(row[col]) else ""
                    data_row += f"{value:<{width}}"
                table_string += data_row + "\n"
                
            return table_string

        # Create the original data tab
        original_tab = ttk.Frame(notebook)
        notebook.add(original_tab, text=self.texts['preview_original_data'].format(n=PREVIEW_ROWS))
        
        # Create frame for original data with scrollbars
        original_frame = ttk.Frame(original_tab)
        original_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Add Text widget with both scrollbars
        original_text = tk.Text(original_frame, wrap=tk.NONE, font=("Courier", 15))
        original_text.insert(tk.END, format_dataframe_as_table(original_df_sample))
        original_text.config(state="disabled")
        
        # Create and position scrollbars
        original_vsb = ttk.Scrollbar(original_frame, orient="vertical", command=original_text.yview)
        original_hsb = ttk.Scrollbar(original_frame, orient="horizontal", command=original_text.xview)
        original_text.configure(yscrollcommand=original_vsb.set, xscrollcommand=original_hsb.set)
        
        # Grid layout for text and scrollbars
        original_frame.grid_rowconfigure(0, weight=1)
        original_frame.grid_columnconfigure(0, weight=1)
        original_text.grid(row=0, column=0, sticky="nsew")
        original_vsb.grid(row=0, column=1, sticky="ns")
        original_hsb.grid(row=1, column=0, sticky="ew")
        
        # Create the modified data tab
        modified_tab = ttk.Frame(notebook)
        notebook.add(modified_tab, text=self.texts['preview_modified_data'].format(n=PREVIEW_ROWS))
        
        # Create frame for modified data with scrollbars
        modified_frame = ttk.Frame(modified_tab)
        modified_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Add Text widget with both scrollbars
        modified_text = tk.Text(modified_frame, wrap=tk.NONE, font=("Courier", 15))
        
        # If there is styling information, apply it using tags
        has_styling = hasattr(modified_df_sample, '_styled_columns')
        if has_styling:
            # Configure a tag for invalid cells
            modified_text.tag_configure("invalid", background="#FFCCCC")
            
            # Format table with styling
            styled_table = format_dataframe_as_table(modified_df_sample)
            modified_text.insert(tk.END, styled_table)
            
            # Apply tags for invalid cells
            for col, mask in modified_df_sample._styled_columns.items():
                if col in modified_df_sample.columns:
                    # Find column index
                    col_idx = list(modified_df_sample.columns).index(col)
                    
                    # Apply styling to each invalid cell
                    for row_idx, is_invalid in enumerate(mask):
                        if is_invalid:
                            # Calculate the line for this row (header + separator + rows)
                            line_num = row_idx + 2  # +2 for header and separator
                            
                            # Calculate the position in this line for the cell
                            line_start = f"{line_num + 1}.0"
                            
                            # Get the full line text
                            line_text = modified_text.get(f"{line_num + 1}.0", f"{line_num + 1}.end")
                            
                            # Skip the row number and find the position of the column
                            # Simple approach: Calculate approximate position based on column widths
                            pos = 5  # Start after row number column
                            for j, c in enumerate(modified_df_sample.columns):
                                if j == col_idx:
                                    break
                                # Add width of this column to position
                                col_width = max(len(str(c)), modified_df_sample[c].astype(str).str.len().max()) + 3
                                pos += col_width
                                
                            # Apply tag from this position to end of the cell
                            cell_value = str(modified_df_sample.iloc[row_idx, col_idx])
                            start_pos = f"{line_num + 1}.{pos}"
                            end_pos = f"{start_pos}+{len(cell_value)}c"
                            modified_text.tag_add("invalid", start_pos, end_pos)
        else:
            # Just add the table without styling
            modified_text.insert(tk.END, format_dataframe_as_table(modified_df_sample))
            
        modified_text.config(state="disabled")
        
        # Create and position scrollbars
        modified_vsb = ttk.Scrollbar(modified_frame, orient="vertical", command=modified_text.yview)
        modified_hsb = ttk.Scrollbar(modified_frame, orient="horizontal", command=modified_text.xview)
        modified_text.configure(yscrollcommand=modified_vsb.set, xscrollcommand=modified_hsb.set)
        
        # Grid layout for text and scrollbars
        modified_frame.grid_rowconfigure(0, weight=1)
        modified_frame.grid_columnconfigure(0, weight=1)
        modified_text.grid(row=0, column=0, sticky="nsew")
        modified_vsb.grid(row=0, column=1, sticky="ns")
        modified_hsb.grid(row=1, column=0, sticky="ew")
        
        # Add note about validation styling if needed
        if has_styling:
            validation_note = ttk.Label(
                modified_tab, 
                text=self.texts.get('validation_highlight_note', "Cells highlighted in red failed validation."),
                font=("", 9, "italic")
            )
            validation_note.pack(pady=(5, 0))
        
        # Add OK button at the bottom
        ttk.Button(main_frame, text="OK", command=preview_dialog.destroy).pack(pady=10)

    def preview_operation(self):
        """Show side-by-side preview: original loaded sample vs current output sample."""
        if self.dataframe is None or self.dataframe.empty:
            messagebox.showwarning(
                self.texts['warning'],
                self.texts['preview_no_data'],
                parent=self.root
            )
            self.update_status(
                self.texts['preview_status_message'].format(message=self.texts['preview_no_data'])
            )
            return

        orig = getattr(self, 'original_df', self.dataframe)
        original_sample = orig.head(PREVIEW_ROWS).copy(deep=True)

        # Use generate_preview to apply the operation preview
        op_text = self.selected_operation.get()
        op_key = self.get_operation_key(op_text)
        modified_sample = self.dataframe.head(PREVIEW_ROWS).copy(deep=True)
        if op_key:
            preview_df, success, msg = generate_preview(self, op_key, self.selected_column.get(), modified_sample, PREVIEW_ROWS)
            if success:
                modified_sample = preview_df
            else:
                self.update_status(self.texts.get('preview_failed', "Preview failed: {error}").format(error=msg))
                # Optionally show a warning dialog
                # messagebox.showwarning(self.texts['warning'], msg, parent=self.root)

        # Preserve styling information (if present) using a safer approach
        if hasattr(self.dataframe, '_styled_columns'):
            object.__setattr__(modified_sample, '_styled_columns', {})
            for col, mask in self.dataframe._styled_columns.items():
                if col in modified_sample.columns:
                    modified_sample._styled_columns[col] = mask.head(PREVIEW_ROWS).copy()

        self.show_preview_dialog(
            original_sample,
            modified_sample,
            self.texts.get('preview_output_title', "Output Preview")
        )

    def preview_output_file(self):
        """Show preview of the current state of the dataframe (all operations applied)."""
        if self.dataframe is None or self.dataframe.empty:
            messagebox.showwarning(
                self.texts['warning'],
                self.texts.get('preview_no_data', "No data to preview."),
                parent=self.root
            )
            self.update_status("Output file preview failed: No data available.")
            return

        # Get a sample of the original file for comparison
        orig = getattr(self, 'original_df', self.dataframe)
        original_sample = orig.head(PREVIEW_ROWS).copy(deep=True)
            
        # Get a sample of the current state with all operations applied
        current_sample = self.dataframe.head(PREVIEW_ROWS).copy(deep=True)
            
        # Preserve styling information if present
        if hasattr(self.dataframe, '_styled_columns'):
            object.__setattr__(current_sample, '_styled_columns', {})
            for col, mask in self.dataframe._styled_columns.items():
                if col in current_sample.columns:
                    current_sample._styled_columns[col] = mask.head(PREVIEW_ROWS).copy()
                
        # Show the preview dialog comparing original to current state
        self.show_preview_dialog(
            original_sample,
            current_sample,
            self.texts.get('output_file_preview_title', "Current Output File")
        )
        
        # Display summary of changes in status log
        original_cols = set(original_sample.columns)
        current_cols = set(current_sample.columns)
        added_cols = current_cols - original_cols
        removed_cols = original_cols - current_cols
        
        summary_parts = []
        
        # Report on column changes
        if added_cols:
            summary_parts.append(f"Added columns: {', '.join(sorted(added_cols))}")
        if removed_cols:
            summary_parts.append(f"Removed columns: {', '.join(sorted(removed_cols))}")
            
        # Report on row changes
        orig_rows = len(self.original_df) if hasattr(self, 'original_df') else "unknown"
        current_rows = len(self.dataframe)
        if orig_rows != current_rows:
            summary_parts.append(f"Rows changed: {orig_rows} → {current_rows}")
        
        # Report on output format
        output_format = self.output_extension.get()
        summary_parts.append(f"Output format: {output_format.upper()}")
        
        # Log the summary
        if summary_parts:
            summary = " | ".join(summary_parts)
            self.update_status(f"Output file preview: {summary}")
            
            
    # --- Undo/Redo Methods ---
    def _commit_undoable_action(self, old_df):
        """Saves the current 'old' DataFrame in undo stack and clears redo stack."""
        self.undo_stack.append(old_df)
        self.redo_stack.clear()
        self.update_undo_redo_buttons()

    def update_undo_redo_buttons(self):
        """Enable or disable Undo/Redo buttons based on stack states."""
        if self.undo_stack:
            self.undo_button.config(state="normal")
        else:
            self.undo_button.config(state="disabled")

        if self.redo_stack:
            self.redo_button.config(state="normal")
        else:
            self.redo_button.config(state="disabled")

    def undo_action(self):
        if not self.undo_stack:
            messagebox.showwarning(self.texts['warning'], self.texts['nothing_to_undo'], parent=self.root)
            return
        # Move current state to redo stack
        self.redo_stack.append(self.dataframe)
        # Restore from undo stack
        last_df = self.undo_stack.pop()
        self.dataframe = last_df
        self.update_column_combobox()
        messagebox.showinfo(self.texts['success'], self.texts['undo_success'], parent=self.root)
        self.update_undo_redo_buttons()
        self.update_status("Undo action performed.")

    def redo_action(self):
        if not self.redo_stack:
            messagebox.showwarning(self.texts['warning'], self.texts['nothing_to_redo'], parent=self.root)
            return
        # Move current state to undo stack
        self.undo_stack.append(self.dataframe)
        # Restore from redo stack
        next_df = self.redo_stack.pop()
        self.dataframe = next_df
        self.update_column_combobox()
        messagebox.showinfo(self.texts['success'], self.texts['redo_success'], parent=self.root)
        self.update_undo_redo_buttons()
        self.update_status("Redo action performed.")
    def apply_operation(self):
        if self.dataframe is None:
            messagebox.showwarning(self.texts['warning'], self.texts['no_file'])
            self.update_status("Operation failed: No file loaded.")
            return

        col = self.selected_column.get()
        op_text = self.selected_operation.get()
        op_key = self.get_operation_key(op_text)

        if op_key == "op_concatenate":
            self.apply_concatenate_ui()
            return
        if op_key == "op_remove_duplicates":
            if not col:
                messagebox.showwarning(self.texts['warning'], self.texts['no_column'])
                self.update_status("Operation failed: No column selected.")
                return
            self.apply_remove_duplicates_ui(col)
            return
        if op_key == "op_merge_columns":
            self.apply_merge_columns_ui()
            return
        if op_key == "op_rename_column":
            # ask for new name, ensure unique
            new_name = self.get_new_column_name(col)
            if not new_name:
                return
            old_df = self.dataframe
            new_df = self.dataframe.copy(deep=True)
            new_df, (status_type, status_message) = apply_rename_column(new_df, col, new_name, self.texts)
            if status_type == 'success':
                self._commit_undoable_action(old_df.copy(deep=True))
                self.dataframe = new_df
                self.update_column_combobox(new_name)
                self.update_status(f"Renamed '{col}' to '{new_name}'.")
                messagebox.showinfo(self.texts['success'], status_message, parent=self.root)
            else:
                messagebox.showerror(self.texts['error'], status_message, parent=self.root)
            return

        if not col:
            messagebox.showwarning(self.texts['warning'], self.texts['no_column'])
            self.update_status("Operation failed: No column selected.")
            return
        if not op_key:
            messagebox.showwarning(self.texts['warning'], self.texts['no_operation'])
            self.update_status("Operation failed: No operation selected.")
            return

        # Work on a copy so we can commit changes as a single undoable step on success
        old_df = self.dataframe
        new_df = self.dataframe.copy(deep=True)

        rows_before = len(new_df)
        cols_before = len(new_df.columns)

        try:
            new_dataframe = None
            status_type = 'info'
            status_message = ""
            refresh_columns = False

            if op_key == "op_mask":
                new_df[col] = new_df[col].astype(str).apply(mask_data, column_name=col)
                status_type = 'success'
                status_message = self.texts['masked_success'].format(col=col)
                self.update_status(f"Masking applied to column '{col}'.")
            elif op_key == "op_mask_email":
                # Create a new column for tracking invalid emails
                invalid_mask = pd.Series(False, index=new_df.index)
                
                # Apply masking with tracking
                result_series = new_df[col].astype(str).apply(
                    lambda x: mask_data(x, mode='email', column_name=col, track_invalid=True)
                )
                
                # Separate the masked values and validity flags
                new_df[col] = result_series.apply(lambda x: x[0] if isinstance(x, tuple) else x)
                
                # Track which rows had invalid emails
                invalid_mask = result_series.apply(lambda x: isinstance(x, tuple) and not x[1])
                
                # Set up the styling for invalid cells
                if not hasattr(new_df, '_styled_columns'):
                    object.__setattr__(new_df, '_styled_columns', {})
                new_df._styled_columns[col] = invalid_mask
                
                status_type = 'success'
                status_message = self.texts['email_masked_success'].format(col=col)
                if invalid_mask.any():
                    status_message += f" ({invalid_mask.sum()} invalid emails highlighted)"
                self.update_status(f"Email masking applied to column '{col}'. {invalid_mask.sum()} invalid emails highlighted.")
            elif op_key == "op_mask_words":
                new_df[col] = new_df[col].astype(str).apply(mask_words, column_name=col)
                status_type = 'success'
                status_message = self.texts['masked_words_success'].format(col=col)
                self.update_status(f"Masked words in column '{col}'.")
            elif op_key == "op_trim":
                new_df[col] = new_df[col].astype(str).apply(trim_spaces, column_name=col)
                status_type = 'success'
                status_message = self.texts['trimmed_success'].format(col=col)
                self.update_status(f"Trimmed spaces in column '{col}'.")
            elif op_key == "op_split_delimiter":
                delimiter = self.get_input('input_needed', 'enter_delimiter')
                if delimiter is None:
                    return
                if delimiter == "":
                    messagebox.showwarning(self.texts['warning'], "Empty delimiter provided. Please enter a valid delimiter.", parent=self.root)
                    return
                new_dataframe, (status_type, status_message) = apply_split_by_delimiter(new_df, col, delimiter, self.texts)
                refresh_columns = True
                if status_type == 'success':
                    self.update_status(f"Split column '{col}' by delimiter '{delimiter}'.")
            elif op_key == "op_split_surname":
                new_dataframe, (status_type, status_message) = apply_split_surname(new_df, col, self.texts)
                refresh_columns = True
                if status_type == 'success':
                    self.update_status(f"Split surname from column '{col}'.")
            elif op_key == "op_upper":
                new_df[col] = new_df[col].astype(str).apply(change_case, case_type='upper', column_name=col)
                status_type = 'success'
                status_message = self.texts['case_change_success'].format(col=col, case_type='UPPERCASE')
                self.update_status(f"Changed case in column '{col}' to UPPERCASE.")
            elif op_key == "op_lower":
                new_df[col] = new_df[col].astype(str).apply(change_case, case_type='lower', column_name=col)
                status_type = 'success'
                status_message = self.texts['case_change_success'].format(col=col, case_type='lowercase')
                self.update_status(f"Changed case in column '{col}' to lowercase.")
            elif op_key == "op_title":
                new_df[col] = new_df[col].astype(str).apply(change_case, case_type='title', column_name=col)
                status_type = 'success'
                status_message = self.texts['case_change_success'].format(col=col, case_type='Title Case')
                self.update_status(f"Changed case in column '{col}' to Title Case.")
            elif op_key == "op_find_replace":
                find_text = self.get_input('input_needed', 'enter_find_text')
                if find_text is not None:
                    replace_text = self.get_input('input_needed', 'enter_replace_text')
                    if replace_text is not None:
                        new_df[col] = new_df[col].astype(str).apply(find_replace, find_text=find_text, replace_text=replace_text, column_name=col)
                        status_type = 'success'
                        status_message = self.texts['find_replace_success'].format(col=col)
                        self.update_status(f"Performed find/replace in column '{col}'.")
            elif op_key == "op_remove_specific":
                chars = self.get_input('input_needed', 'enter_chars_to_remove')
                if chars:
                    new_df[col] = new_df[col].astype(str).apply(remove_chars, mode='specific', chars_to_remove=chars, column_name=col)
                    status_type = 'success'
                    status_message = self.texts['remove_chars_success'].format(col=col)
                    self.update_status(f"Removed specific characters in column '{col}'.")
            elif op_key == "op_remove_non_numeric":
                new_df[col] = new_df[col].astype(str).apply(remove_chars, mode='non_numeric', column_name=col)
                status_type = 'success'
                status_message = self.texts['remove_chars_success'].format(col=col)
                self.update_status(f"Removed non-numeric characters in column '{col}'.")
            elif op_key == "op_remove_non_alpha":
                new_df[col] = new_df[col].astype(str).apply(remove_chars, mode='non_alphabetic', column_name=col)
                status_type = 'success'
                status_message = self.texts['remove_chars_success'].format(col=col)
                self.update_status(f"Removed non-alphabetic characters in column '{col}'.")
            elif op_key == "op_extract_pattern":
                pattern = self.get_input('input_needed', 'enter_regex_pattern')
                if pattern:
                    try:
                        re.compile(pattern)
                        new_col_base = f"{col}_extracted"
                        new_col_name = self.get_new_column_name(new_col_base)
                        if new_col_name:
                            new_dataframe, (status_type, status_message) = apply_extract_pattern(new_df, col, new_col_name, pattern, self.texts)
                            refresh_columns = True
                            if status_type == 'success':
                                self.update_status(f"Extracted pattern from column '{col}' into '{new_col_name}'.")
                    except re.error as e:
                        status_type = 'error'
                        status_message = self.texts['regex_error'].format(error=e)
                        self.update_status(f"Regex error: {e}")
            elif op_key == "op_fill_missing":
                fill_val = self.get_input('input_needed', 'enter_fill_value')
                if fill_val is not None:
                    new_df[col] = new_df[col].apply(fill_missing, fill_value=fill_val, column_name=col)
                    status_type = 'success'
                    status_message = self.texts['fill_missing_success'].format(col=col)
                    self.update_status(f"Filled missing values in column '{col}'.")
            elif op_key == "op_mark_duplicates":
                new_col_base = f"{col}_is_duplicate"
                new_col_name = self.get_new_column_name(new_col_base)
                if new_col_name:
                    new_dataframe, (status_type, status_message) = apply_mark_duplicates(new_df, col, new_col_name, self.texts)
                    refresh_columns = True
                    if status_type == 'success':
                        self.update_status(f"Marked duplicates in column '{col}' into '{new_col_name}'.")
            elif op_key == "op_round_numbers":
                decimals = simpledialog.askinteger(self.texts['input_needed'], self.texts['enter_decimals'], parent=self.root, minvalue=0)
                if decimals is not None:
                    new_df, (status_type, status_message) = apply_round_numbers(new_df, col, decimals, self.texts)
                    if status_type == 'success':
                        self.update_status(f"Rounded column '{col}' to {decimals} decimal places.")
            elif op_key == "op_calculate_column_constant":
                operation = simpledialog.askstring(self.texts['input_needed'], self.texts['select_calculation_operation'], parent=self.root)
                if operation not in ['+', '-', '*', '/']:
                    status_type = 'error'
                    status_message = "Invalid operation. Choose +, -, *, or /."
                else:
                    value = simpledialog.askfloat(self.texts['input_needed'], self.texts['enter_constant_value'], parent=self.root)
                    if value is not None:
                        new_df, (status_type, status_message) = numeric_operations.apply_calculate_column_constant(new_df, col, operation, value, self.texts)
                        if status_type == 'success':
                            self.update_status(f"Calculated column '{col}' by constant {value} using operation {operation}.")
            elif op_key == "op_create_calculated_column":
                # Use column selection dialog instead of text input
                available_columns = [c for c in new_df.columns if c != col]
                if not available_columns:
                    messagebox.showwarning(self.texts['warning'], "There are no other columns available for calculation.", parent=self.root)
                    return
                    
                # Create a simpler single column selection dialog
                col2_dialog = tk.Toplevel(self.root)
                col2_dialog.title(self.texts['select_second_column_calc'])
                col2_dialog.transient(self.root)
                col2_dialog.grab_set()
                col2_dialog.geometry("300x300")
                
                ttk.Label(col2_dialog, text=self.texts['select_second_column_calc']).pack(pady=5)
                
                listbox_frame = ttk.Frame(col2_dialog)
                listbox_frame.pack(expand=True, fill="both", padx=10, pady=5)
                
                listbox = tk.Listbox(listbox_frame, selectmode="single", exportselection=False)
                listbox.pack(side="left", expand=True, fill="both")
                
                scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=listbox.yview)
                scrollbar.pack(side="right", fill="y")
                listbox.config(yscrollcommand=scrollbar.set)
                
                for column_name in available_columns:
                    listbox.insert(tk.END, column_name)
                
                selected_col2 = [None]  # Use a list to store the selection
                
                def on_col2_ok():
                    selected_indices = listbox.curselection()
                    if not selected_indices:
                        messagebox.showwarning(self.texts['warning'], self.texts['no_column_selected'], parent=col2_dialog)
                        return
                    selected_col2[0] = listbox.get(selected_indices[0])
                    col2_dialog.destroy()
                
                def on_col2_cancel():
                    col2_dialog.destroy()
                
                button_frame = ttk.Frame(col2_dialog)
                button_frame.pack(pady=10)
                ttk.Button(button_frame, text="OK", command=on_col2_ok).pack(side="left", padx=5)
                ttk.Button(button_frame, text="Cancel", command=on_col2_cancel).pack(side="left", padx=5)
                
                self.root.wait_window(col2_dialog)
                
                col2 = selected_col2[0]
                if col2 is None:
                    return
                
                operation = simpledialog.askstring(self.texts['input_needed'], self.texts['select_arithmetic_operation'], parent=self.root)
                new_col_name = self.get_new_column_name(f"{col}_{col2}_calculated")
                if col2 and operation and new_col_name:
                     new_df, (status_type, status_message) = numeric_operations.apply_create_calculated_column(new_df, col, col2, operation, new_col_name, self.texts)
                     if status_type == 'success':
                         self.update_status(f"Created new column '{new_col_name}' by calculating '{col}' and '{col2}' using operation '{operation}'.")
                         refresh_columns = True
            elif op_key in ["op_validate_email", "op_validate_phone", "op_validate_date", 
                           "op_validate_numeric", "op_validate_alphanumeric", "op_validate_url"]:
                # Extract validation type from operation key
                validation_type = op_key.replace("op_validate_", "")
                
                # Apply validation directly using the type from the operation key
                try:
                    new_df, result = apply_validation(new_df, col, validation_type, self.texts)
                    
                    if result[0] == 'success':
                        status_type, status_message = result[0], result[1]
                        self._commit_undoable_action(old_df.copy(deep=True))
                        self.dataframe = new_df
                        self.update_column_combobox()
                        self.update_status(f"Validated column '{col}' with type '{validation_type}'.")
                        messagebox.showinfo(self.texts['success'], status_message, parent=self.root)
                    else:
                        messagebox.showerror(self.texts['error'], result[1], parent=self.root)
                        
                    refresh_columns = True
                    return  # Already handled the dataframe update
                except Exception as e:
                    messagebox.showerror(
                        self.texts['error'], 
                        self.texts['operation_error'].format(error=e), 
                        parent=self.root
                    )
                    self.update_status(f"Validation operation failed: {e}")
                    return  # Exit early after error
                    
            # Apply all the changes to the main dataframe after successful operation
            if status_type == 'success':
                if new_dataframe is not None:
                    # If a new dataframe was created (e.g., for split operations)
                    self._commit_undoable_action(old_df.copy(deep=True))
                    self.dataframe = new_dataframe
                    if refresh_columns:
                        self.update_column_combobox()
                    messagebox.showinfo(self.texts['success'], status_message, parent=self.root)
                else:
                    # If we modified the existing dataframe
                    self._commit_undoable_action(old_df.copy(deep=True))
                    self.dataframe = new_df
                    if refresh_columns:
                        self.update_column_combobox()
                    messagebox.showinfo(self.texts['success'], status_message, parent=self.root)
            elif status_type == 'error':
                messagebox.showerror(self.texts['error'], status_message, parent=self.root)

        except Exception as e:
            status_type = 'error'
            status_message = self.texts['operation_error'].format(error=e)
            self.update_status(f"Operation failed: {e}")
            messagebox.showerror(self.texts['error'], status_message, parent=self.root)

    def apply_concatenate_ui(self):
        cols_to_concat = self.get_multiple_columns('input_needed', 'select_columns_concat')
        if not cols_to_concat or len(cols_to_concat) < 2:
            if cols_to_concat is not None:
                messagebox.showwarning(self.texts['warning'], "Please select at least two columns.", parent=self.root)
            return

        separator = self.get_input('input_needed', 'enter_separator')
        if separator is None:
            return

        new_col_base = "_".join(cols_to_concat) + "_concat"
        new_col_name = self.get_new_column_name(new_col_base)
        if not new_col_name:
            return

        # Prepare for undo
        old_df = self.dataframe
        new_df = self.dataframe.copy(deep=True)

        try:
            new_df, (status_type, status_message) = apply_concatenate(new_df, cols_to_concat, new_col_name, separator, self.texts)
            if status_type == 'success':
                self._commit_undoable_action(old_df.copy(deep=True))
                self.dataframe = new_df
                self.update_column_combobox(new_col_name)
                self.update_status(f"Concatenated columns into '{new_col_name}'. Columns: {len(self.dataframe.columns)}")
                messagebox.showinfo(self.texts['success'], status_message, parent=self.root)
            else:
                messagebox.showerror(self.texts['error'], status_message, parent=self.root)
        except Exception as e:
            messagebox.showerror(self.texts['error'], self.texts['operation_error'].format(error=e), parent=self.root)
            self.update_status(f"Concatenate operation failed: {e}")

    def apply_remove_duplicates_ui(self, col):
        if messagebox.showyesno(self.texts['warning'],
                               f"This will permanently remove rows based on duplicates in column '{col}'.\nAre you sure?",
                               parent=self.root):
            old_df = self.dataframe
            new_df = self.dataframe.copy(deep=True)

            rows_before = len(new_df)
            try:
                new_df, (status_type, status_message) = apply_remove_duplicates(new_df, col, self.texts)
                if status_type == 'success':
                    self._commit_undoable_action(old_df.copy(deep=True))
                    self.dataframe = new_df
                    self.update_column_combobox(col)
                    rows_after = len(self.dataframe)
                    self.update_status(f"Removed {rows_before - rows_after} duplicate rows based on '{col}'.")
                    messagebox.showinfo(self.texts['success'], status_message, parent=self.root)
                else:
                    messagebox.showerror(self.texts['error'], status_message, parent=self.root)
            except Exception as e:
                messagebox.showerror(self.texts['error'], self.texts['operation_error'].format(error=e), parent=self.root)
                self.update_status(f"Remove duplicates operation failed: {e}")

    def apply_merge_columns_ui(self):
        cols_to_merge = self.get_multiple_columns('input_needed', 'select_columns_merge')
        if not cols_to_merge or len(cols_to_merge) < 2:
            if cols_to_merge is not None:
                messagebox.showwarning(self.texts['warning'], "Please select at least two columns to merge.", parent=self.root)
            return

        separator = self.get_input('input_needed', 'enter_separator')
        if separator is None:
            return

        fill_missing = messagebox.askyesno(
            self.texts['input_needed'],
            self.texts['fill_missing_merge'],
            parent=self.root
        )

        new_col_base = "_".join(cols_to_merge) + "_merged"
        new_col_name = self.get_new_column_name(new_col_base)
        if not new_col_name:
            return

        # Prepare for undo
        old_df = self.dataframe
        new_df = self.dataframe.copy(deep=True)

        try:
            new_df, (status_type, status_message) = apply_merge_columns(
                new_df, cols_to_merge, new_col_name, separator, fill_missing, self.texts
            )
            if status_type == 'success':
                self._commit_undoable_action(old_df.copy(deep=True))
                self.dataframe = new_df
                self.update_column_combobox(new_col_name)
                self.update_status(f"Merged columns into '{new_col_name}'. Columns: {len(self.dataframe.columns)}")
                messagebox.showinfo(self.texts['success'], status_message, parent=self.root)
            else:
                messagebox.showerror(self.texts['error'], status_message, parent=self.root)
        except Exception as e:
            messagebox.showerror(self.texts['error'], self.texts['operation_error'].format(error=e), parent=self.root)
            self.update_status(f"Merge columns operation failed: {e}")

    def update_column_combobox(self, preferred_selection=None):
        if self.dataframe is not None:
            cols = list(self.dataframe.columns)
            self.column_combobox['values'] = cols
            if preferred_selection and preferred_selection in cols:
                self.selected_column.set(preferred_selection)
            elif cols:
                self.selected_column.set(cols[0])
            else:
                self.selected_column.set("")
        else:
            self.column_combobox['values'] = []
            self.selected_column.set("")

    def save_file(self):
        if self.dataframe is None:
            messagebox.showwarning(self.texts['warning'], self.texts['no_data_to_save'])
            self.update_status("Save operation failed: No data to save.")
            return

        chosen_ext = self.output_extension.get()  # 'xls', 'xlsx', or 'csv'

        original_name = os.path.splitext(os.path.basename(self.file_path.get()))[0]
        suggested_name = f"{original_name}_modified.{chosen_ext}"


        save_path = filedialog.asksaveasfilename(
            initialdir=self.last_dir,
            title=self.texts['save_modified_file'],
            initialfile=suggested_name,
            defaultextension="." + chosen_ext,
            filetypes=[(self.texts['excel_files'], "*.xlsx *.xls *.csv")]
        )

        if save_path:
            self.last_dir = os.path.dirname(save_path)
            self.save_last_directory()  # Save the directory for future sessions
            try:
                if chosen_ext == "csv":
                    # CSV doesn't support styling, so just save normally
                    self.dataframe.to_csv(save_path, index=False)
                elif chosen_ext == "json":
                    # JSON doesn't support styling, so just save normally
                    self.dataframe.to_json(save_path, orient="records", indent=2)
                elif chosen_ext == "html":
                    # For HTML, we can apply the styling
                    if hasattr(self.dataframe, '_styled_columns'):
                        # Create a styled dataframe
                        styled_df = self.dataframe.style
                        for col, invalid_mask in self.dataframe._styled_columns.items():
                            styled_df = styled_df.apply(
                                lambda s: ['background-color: #FFCCCC' if invalid_mask.iloc[i] else '' 
                                          for i in range(len(s))], 
                                axis=0, 
                                subset=[col]
                            )
                        styled_df.to_html(save_path, index=False)
                    else:
                        self.dataframe.to_html(save_path, index=False)
                elif chosen_ext in ("md", "markdown"):
                    # Markdown doesn't support styling, so just save normally
                    with open(save_path, "w") as f:
                        f.write(self.dataframe.to_markdown(index=False))
                else:
                    # xls or xlsx - We can apply styling here
                    if hasattr(self.dataframe, '_styled_columns'):
                        # Create a styled dataframe
                        styled_df = self.dataframe.style
                        for col, invalid_mask in self.dataframe._styled_columns.items():
                            styled_df = styled_df.apply(
                                lambda s: ['background-color: #FFCCCC' if invalid_mask.iloc[i] else '' 
                                          for i in range(len(s))], 
                                axis=0, 
                                subset=[col]
                            )
                        styled_df.to_excel(save_path, index=False, engine='openpyxl')
                    else:
                        # Normal save without styling
                        self.dataframe.to_excel(save_path, index=False, engine='openpyxl')
                        
                messagebox.showinfo(self.texts['success'], self.texts['file_saved_success'].format(path=save_path))
                self.update_status(f"File saved successfully to {os.path.basename(save_path)}.")

            except Exception as e:
                self.update_status(f"Error saving file: {e}")
                messagebox.showerror(self.texts['error'], self.texts['save_error'].format(error=e))
        else:
            self.update_status("Save operation cancelled.")

# Add the main block to start the application
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelEditorApp(root)
    root.mainloop()