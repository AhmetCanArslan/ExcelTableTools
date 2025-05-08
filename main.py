import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, scrolledtext
import pandas as pd
import os
import re

# Import operations
from operations.masking import mask_data, mask_email, mask_words  # Import new function
from operations.trimming import trim_spaces
from operations.splitting import apply_split_surname, apply_split_by_delimiter
from operations.case_change import change_case
from operations.find_replace import find_replace
from operations.remove_chars import remove_chars
from operations.concatenate import apply_concatenate
from operations.extract_pattern import apply_extract_pattern
from operations.fill_missing import fill_missing
from operations.duplicates import apply_mark_duplicates, apply_remove_duplicates

# Import translations
from translations import LANGUAGES

PREVIEW_ROWS = 5  # Number of rows to show in preview

# --- Helper for unique column name ---
def get_unique_col_name(base_name, existing_columns):
    """Generates a unique column name based on existing ones."""
    new_name = base_name
    counter = 1
    while new_name in existing_columns:
        new_name = f"{base_name}_{counter}"
        counter += 1
    return new_name

# --- GUI Application ---
class ExcelEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.geometry("750x550")  # Increased width slightly for preview button

        self.file_path = tk.StringVar()
        self.selected_column = tk.StringVar()
        self.selected_operation = tk.StringVar()
        self.dataframe = None
        self.current_lang = 'en'
        self.texts = LANGUAGES[self.current_lang]

        # --- Undo/Redo History ---
        self.undo_stack = []
        self.redo_stack = []

        # --- Main Content Frame ---
        main_content_frame = ttk.Frame(root)
        main_content_frame.pack(fill="both", expand=True, side=tk.TOP)

        # --- Top Frame for Language Button ---
        top_frame = ttk.Frame(main_content_frame)
        top_frame.pack(fill="x", padx=10, pady=(5, 0))
        self.lang_button = ttk.Button(top_frame, text=self.texts['change_language'], command=self.toggle_language)
        self.lang_button.pack(side="right")
        # Add Refresh button at upper left
        self.refresh_button = ttk.Button(top_frame, text=self.texts['refresh'], command=self.refresh_app)
        self.refresh_button.pack(side="left")

        # --- File Selection ---
        self.file_frame = ttk.LabelFrame(main_content_frame, text=self.texts['file_selection'])
        self.file_frame.pack(padx=10, pady=10, fill="x")

        self.file_label = ttk.Label(self.file_frame, text=self.texts['excel_file'])
        self.file_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.file_entry = ttk.Entry(self.file_frame, textvariable=self.file_path, width=50, state="readonly")
        self.file_entry.grid(row=0, column=1, padx=5, pady=5)
        self.browse_button = ttk.Button(self.file_frame, text=self.texts['browse'], command=self.browse_file)
        self.browse_button.grid(row=0, column=2, padx=5, pady=5)

        # --- Operations ---
        self.ops_frame = ttk.LabelFrame(main_content_frame, text=self.texts['operations'])
        self.ops_frame.pack(padx=10, pady=10, fill="x")

        self.column_label = ttk.Label(self.ops_frame, text=self.texts['column'])
        self.column_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.column_combobox = ttk.Combobox(self.ops_frame, textvariable=self.selected_column, state="disabled", width=45)
        self.column_combobox.grid(row=0, column=1, padx=5, pady=5)

        self.operation_label = ttk.Label(self.ops_frame, text=self.texts['operation'])
        self.operation_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.operation_combobox = ttk.Combobox(self.ops_frame, textvariable=self.selected_operation, state="disabled", width=45)
        self.operation_combobox.grid(row=1, column=1, padx=5, pady=5)

        self.apply_button = ttk.Button(self.ops_frame, text=self.texts['apply_operation'], command=self.apply_operation)
        self.apply_button.grid(row=2, column=0, padx=5, pady=10, sticky="ew")

        self.preview_button = ttk.Button(self.ops_frame, text=self.texts.get('preview_button', "Preview"), command=self.preview_operation)  # Use .get for safety during init
        self.preview_button.grid(row=2, column=1, padx=5, pady=10, sticky="ew")

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
        self.output_formats = ["xls", "xlsx", "csv"]
        self.extension_dropdown = ttk.Combobox(save_frame, textvariable=self.output_extension,
                                               values=self.output_formats, state="readonly", width=5)
        self.extension_dropdown.pack(side="right", padx=5)

        self.save_button = ttk.Button(save_frame, text=self.texts['save_changes'], command=self.save_file)
        self.save_button.pack(side="right", padx=5)

        # Set default to 'xlsx' initially
        self.output_extension.set("xlsx")

        # --- Status Area (CLI-like) ---
        status_frame = ttk.LabelFrame(root, text=self.texts['status_log'])
        status_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(5, 10))

        self.status_text = scrolledtext.ScrolledText(status_frame, height=5, wrap=tk.WORD, state='disabled')
        self.status_text.pack(fill="both", expand=True, padx=5, pady=5)

        self.update_status("Ready.")

        self.update_ui_language()

    def update_status(self, message):
        """Appends a message to the status text area."""
        self.status_text.config(state='normal')
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state='disabled')

    def update_ui_language(self):
        self.texts = LANGUAGES[self.current_lang]
        self.root.title(self.texts['title'])
        self.lang_button.config(text=self.texts['change_language'])

        self.file_frame.config(text=self.texts['file_selection'])
        self.ops_frame.config(text=self.texts['operations'])

        self.file_label.config(text=self.texts['excel_file'])
        self.browse_button.config(text=self.texts['browse'])
        self.column_label.config(text=self.texts['column'])
        self.operation_label.config(text=self.texts['operation'])
        self.apply_button.config(text=self.texts['apply_operation'])
        self.preview_button.config(text=self.texts['preview_button'])  # Added for preview button
        self.save_button.config(text=self.texts['save_changes'])
        self.refresh_button.config(text=self.texts['refresh'])

        # Update Undo/Redo button texts
        self.undo_button.config(text=self.texts.get('undo', "Undo"))
        self.redo_button.config(text=self.texts.get('redo', "Redo"))

        self.operation_keys = [
            "op_mask", "op_trim", "op_split_space", "op_split_colon", "op_split_surname",
            "op_upper", "op_lower", "op_title",
            "op_find_replace", "op_remove_specific", "op_remove_non_numeric", "op_remove_non_alpha",
            "op_concatenate", "op_extract_pattern", "op_fill_missing",
            "op_mark_duplicates", "op_remove_duplicates",
            "op_mask_email",  # Added
            "op_mask_words"   # Added
        ]
        translated_ops = [self.texts[key] for key in self.operation_keys]
        current_selection_text = self.selected_operation.get()
        self.operation_combobox['values'] = translated_ops

        if current_selection_text:
            try:
                current_key = None
                old_lang = 'tr' if self.current_lang == 'en' else 'en'
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

        status_frame_children = self.root.winfo_children()
        for child in status_frame_children:
            if isinstance(child, ttk.LabelFrame) and hasattr(child, 'status_text'):
                child.config(text=self.texts.get('status_log', "Status Log"))
                break

    def toggle_language(self):
        self.current_lang = 'tr' if self.current_lang == 'en' else 'en'
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

        # Inform user
        self.update_status("Application refreshed.")

    def get_operation_key(self, translated_op_text):
        for key in self.operation_keys:
            if self.texts[key] == translated_op_text:
                return key
        return None

    def browse_file(self):
        path = filedialog.askopenfilename(
            title=self.texts['select_excel_file'],
            filetypes=[(self.texts['excel_files'], "*.xlsx *.xls *.csv")]
        )
        if path:
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
                self.dataframe = pd.read_csv(path)
            elif path.lower().endswith('.xlsx'):
                self.dataframe = pd.read_excel(path, engine='openpyxl')
            else:
                self.dataframe = pd.read_excel(path)

            # Reset history after loading
            self.undo_stack = []
            self.redo_stack = []

            # Commit the loaded DataFrame as a baseline state
            self._commit_undoable_action(self.dataframe.copy(deep=True))

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
        preview_dialog.geometry("700x450")  # Adjusted size

        main_frame = ttk.Frame(preview_dialog, padding="10")
        main_frame.pack(expand=True, fill="both")

        ttk.Label(main_frame, text=f"{op_text} - {self.texts['preview_display_title']}").pack(pady=5)

        notebook = ttk.Notebook(main_frame)
        notebook.pack(expand=True, fill="both", pady=5)

        # Original Data Tab
        original_tab = ttk.Frame(notebook)
        notebook.add(original_tab, text=self.texts['preview_original_data'].format(n=PREVIEW_ROWS))
        
        original_text_area = scrolledtext.ScrolledText(original_tab, wrap=tk.NONE, height=10)
        original_text_area.insert(tk.END, original_df_sample.to_string())
        original_text_area.config(state='disabled')
        original_text_area.pack(expand=True, fill="both", padx=5, pady=5)
        # Add horizontal scrollbar for original_text_area
        original_h_scroll = ttk.Scrollbar(original_tab, orient=tk.HORIZONTAL, command=original_text_area.xview)
        original_h_scroll.pack(fill=tk.X, side=tk.BOTTOM)
        original_text_area.config(xscrollcommand=original_h_scroll.set)

        # Modified Data Tab
        modified_tab = ttk.Frame(notebook)
        notebook.add(modified_tab, text=self.texts['preview_modified_data'].format(n=PREVIEW_ROWS))

        modified_text_area = scrolledtext.ScrolledText(modified_tab, wrap=tk.NONE, height=10)
        modified_text_area.insert(tk.END, modified_df_sample.to_string())
        modified_text_area.config(state='disabled')
        modified_text_area.pack(expand=True, fill="both", padx=5, pady=5)
        # Add horizontal scrollbar for modified_text_area
        modified_h_scroll = ttk.Scrollbar(modified_tab, orient=tk.HORIZONTAL, command=modified_text_area.xview)
        modified_h_scroll.pack(fill=tk.X, side=tk.BOTTOM)
        modified_text_area.config(xscrollcommand=modified_h_scroll.set)
        
        ttk.Button(main_frame, text="OK", command=preview_dialog.destroy).pack(pady=10)

    def preview_operation(self):
        if self.dataframe is None or self.dataframe.empty:
            messagebox.showwarning(self.texts['warning'], self.texts['preview_no_data'], parent=self.root)
            self.update_status(self.texts['preview_status_message'].format(message=self.texts['preview_no_data']))
            return

        col = self.selected_column.get()
        op_text = self.selected_operation.get()
        op_key = self.get_operation_key(op_text)

        if not op_key:
            messagebox.showwarning(self.texts['warning'], self.texts['no_operation'], parent=self.root)
            self.update_status(self.texts['preview_status_message'].format(message=self.texts['no_operation']))
            return

        is_concatenate_op = op_key == "op_concatenate"
        if not is_concatenate_op and not col:
            messagebox.showwarning(self.texts['warning'], self.texts['no_column'], parent=self.root)
            self.update_status(self.texts['preview_status_message'].format(message=self.texts['no_column']))
            return

        original_sample = self.dataframe.head(PREVIEW_ROWS).copy(deep=True)
        preview_df = self.dataframe.head(PREVIEW_ROWS).copy(deep=True)
        
        status_message = ""
        preview_successful = True
        requires_input_ops = ["op_find_replace", "op_remove_specific", "op_fill_missing", "op_extract_pattern", "op_concatenate"]

        if op_key in requires_input_ops:
            self.update_status(self.texts['preview_requires_input'])

        try:
            if op_key == "op_mask":
                preview_df[col] = preview_df[col].astype(str).apply(mask_data)
            elif op_key == "op_mask_email":
                preview_df[col] = preview_df[col].astype(str).apply(mask_data, mode='email')
            elif op_key == "op_mask_words":
                preview_df[col] = preview_df[col].astype(str).apply(mask_words)
            elif op_key == "op_trim":
                preview_df[col] = preview_df[col].astype(str).apply(trim_spaces)
            elif op_key in ["op_upper", "op_lower", "op_title"]:
                case_type_map = {"op_upper": "upper", "op_lower": "lower", "op_title": "title"}
                preview_df[col] = preview_df[col].astype(str).apply(change_case, case_type=case_type_map[op_key])
            elif op_key == "op_find_replace":
                find_text = simpledialog.askstring(self.texts['input_needed'], self.texts['enter_find_text'] + " (for preview)", parent=self.root)
                if find_text is not None:
                    replace_text = simpledialog.askstring(self.texts['input_needed'], self.texts['enter_replace_text'] + " (for preview)", parent=self.root)
                    if replace_text is not None:
                        preview_df[col] = preview_df[col].astype(str).apply(find_replace, find_text=find_text, replace_text=replace_text)
                    else: preview_successful = False; status_message = "Replace text cancelled for preview."
                else: preview_successful = False; status_message = "Find text cancelled for preview."
            elif op_key == "op_remove_specific":
                chars = simpledialog.askstring(self.texts['input_needed'], self.texts['enter_chars_to_remove'] + " (for preview)", parent=self.root)
                if chars is not None:
                    preview_df[col] = preview_df[col].astype(str).apply(remove_chars, mode='specific', chars_to_remove=chars)
                else: preview_successful = False; status_message = "Chars to remove input cancelled for preview."
            elif op_key == "op_remove_non_numeric":
                preview_df[col] = preview_df[col].astype(str).apply(remove_chars, mode='non_numeric')
            elif op_key == "op_remove_non_alpha":
                preview_df[col] = preview_df[col].astype(str).apply(remove_chars, mode='non_alphabetic')
            elif op_key == "op_fill_missing":
                fill_val = simpledialog.askstring(self.texts['input_needed'], self.texts['enter_fill_value'] + " (for preview)", parent=self.root)
                if fill_val is not None:
                    preview_df[col] = preview_df[col].apply(fill_missing, fill_value=fill_val)
                else: preview_successful = False; status_message = "Fill value input cancelled for preview."
            elif op_key == "op_split_space":
                preview_df, (stype, smessage) = apply_split_by_delimiter(preview_df.copy(), col, ' ', self.texts)
                if stype == 'error': preview_successful = False; status_message = smessage
                elif stype == 'warning': self.update_status(self.texts['preview_status_message'].format(message=smessage))  # Show warning but continue
            elif op_key == "op_split_colon":
                preview_df, (stype, smessage) = apply_split_by_delimiter(preview_df.copy(), col, ':', self.texts)
                if stype == 'error': preview_successful = False; status_message = smessage
                elif stype == 'warning': self.update_status(self.texts['preview_status_message'].format(message=smessage))
            elif op_key == "op_split_surname":
                preview_df, (stype, smessage) = apply_split_surname(preview_df.copy(), col, self.texts)
                if stype == 'error': preview_successful = False; status_message = smessage
            elif op_key == "op_extract_pattern":
                pattern = simpledialog.askstring(self.texts['input_needed'], self.texts['enter_regex_pattern'] + " (for preview)", parent=self.root)
                if pattern is not None:
                    try:
                        re.compile(pattern)  # Validate regex
                        new_col_name = get_unique_col_name(f"{col}_extracted_preview", preview_df.columns)
                        preview_df, (stype, smessage) = apply_extract_pattern(preview_df.copy(), col, new_col_name, pattern, self.texts)
                        if stype == 'error': preview_successful = False; status_message = smessage
                    except re.error as e:
                        preview_successful = False; status_message = self.texts['regex_error'].format(error=e)
                else: preview_successful = False; status_message = "Regex pattern input cancelled for preview."
            elif op_key == "op_mark_duplicates":
                new_col_name = get_unique_col_name(f"{col}_is_duplicate_preview", preview_df.columns)
                preview_df, (stype, smessage) = apply_mark_duplicates(preview_df.copy(), col, new_col_name, self.texts)
                if stype == 'error': preview_successful = False; status_message = smessage
            elif op_key == "op_remove_duplicates":
                preview_df, (stype, smessage) = apply_remove_duplicates(preview_df.copy(), col, self.texts)
                if stype == 'error': preview_successful = False; status_message = smessage
            elif op_key == "op_concatenate":
                messagebox.showinfo(self.texts['info'], self.texts['preview_not_available_complex'], parent=self.root)
                preview_successful = False  # Simplified: Mark as not successful to avoid showing dialog
                status_message = self.texts['preview_not_available_complex']
            else:
                preview_successful = False
                status_message = self.texts['not_implemented'].format(op=op_text)
                messagebox.showinfo(self.texts['info'], status_message, parent=self.root)
                self.update_status(self.texts['preview_status_message'].format(message=f"Preview for '{op_text}' not available."))
                return

            if preview_successful:
                self.show_preview_dialog(original_sample, preview_df, op_text)
                self.update_status(self.texts['preview_status_message'].format(message=f"Displayed for '{op_text}'."))
            elif status_message:  # If preview failed but there's a specific message (e.g. input cancelled)
                messagebox.showwarning(self.texts['warning'], self.texts['preview_failed'].format(error=status_message), parent=self.root)
                self.update_status(self.texts['preview_status_message'].format(message=f"Preview for '{op_text}' failed or cancelled: {status_message}"))

        except ValueError as ve:  # Catch issues like column not found if logic missed it
            messagebox.showwarning(self.texts['warning'], str(ve), parent=self.root)
            self.update_status(self.texts['preview_status_message'].format(message=f"Preview failed: {str(ve)}"))
        except Exception as e:
            messagebox.showerror(self.texts['error'], self.texts['preview_failed'].format(error=e), parent=self.root)
            self.update_status(self.texts['preview_status_message'].format(message=f"Preview for '{op_text}' failed with error: {e}"))

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
                new_df[col] = new_df[col].astype(str).apply(mask_data)
                status_type = 'success'
                status_message = self.texts['masked_success'].format(col=col)
                self.update_status(f"Masking applied to column '{col}'.")
            elif op_key == "op_mask_email":
                new_df[col] = new_df[col].astype(str).apply(mask_data, mode='email')
                status_type = 'success'
                status_message = self.texts['email_masked_success'].format(col=col)
                self.update_status(f"Email masking applied to column '{col}'.")
            elif op_key == "op_mask_words":
                new_df[col] = new_df[col].astype(str).apply(mask_words)
                status_type = 'success'
                status_message = self.texts['masked_words_success'].format(col=col)
                self.update_status(f"Masked words in column '{col}'.")
            elif op_key == "op_trim":
                new_df[col] = new_df[col].astype(str).apply(trim_spaces)
                status_type = 'success'
                status_message = self.texts['trimmed_success'].format(col=col)
                self.update_status(f"Trimmed spaces in column '{col}'.")
            elif op_key == "op_split_space":
                new_dataframe, (status_type, status_message) = apply_split_by_delimiter(new_df, col, ' ', self.texts)
                refresh_columns = True
                if status_type == 'success':
                    self.update_status(f"Split column '{col}' by space.")
            elif op_key == "op_split_colon":
                new_dataframe, (status_type, status_message) = apply_split_by_delimiter(new_df, col, ':', self.texts)
                refresh_columns = True
                if status_type == 'success':
                    self.update_status(f"Split column '{col}' by colon.")
            elif op_key == "op_split_surname":
                new_dataframe, (status_type, status_message) = apply_split_surname(new_df, col, self.texts)
                refresh_columns = True
                if status_type == 'success':
                    self.update_status(f"Split surname from column '{col}'.")
            elif op_key == "op_upper":
                new_df[col] = new_df[col].astype(str).apply(change_case, case_type='upper')
                status_type = 'success'
                status_message = self.texts['case_change_success'].format(col=col, case_type='UPPERCASE')
                self.update_status(f"Changed case in column '{col}' to UPPERCASE.")
            elif op_key == "op_lower":
                new_df[col] = new_df[col].astype(str).apply(change_case, case_type='lower')
                status_type = 'success'
                status_message = self.texts['case_change_success'].format(col=col, case_type='lowercase')
                self.update_status(f"Changed case in column '{col}' to lowercase.")
            elif op_key == "op_title":
                new_df[col] = new_df[col].astype(str).apply(change_case, case_type='title')
                status_type = 'success'
                status_message = self.texts['case_change_success'].format(col=col, case_type='Title Case')
                self.update_status(f"Changed case in column '{col}' to Title Case.")
            elif op_key == "op_find_replace":
                find_text = self.get_input('input_needed', 'enter_find_text')
                if find_text is not None:
                    replace_text = self.get_input('input_needed', 'enter_replace_text')
                    if replace_text is not None:
                        new_df[col] = new_df[col].astype(str).apply(find_replace, find_text=find_text, replace_text=replace_text)
                        status_type = 'success'
                        status_message = self.texts['find_replace_success'].format(col=col)
                        self.update_status(f"Performed find/replace in column '{col}'.")
            elif op_key == "op_remove_specific":
                chars = self.get_input('input_needed', 'enter_chars_to_remove')
                if chars:
                    new_df[col] = new_df[col].astype(str).apply(remove_chars, mode='specific', chars_to_remove=chars)
                    status_type = 'success'
                    status_message = self.texts['remove_chars_success'].format(col=col)
                    self.update_status(f"Removed specific characters in column '{col}'.")
            elif op_key == "op_remove_non_numeric":
                new_df[col] = new_df[col].astype(str).apply(remove_chars, mode='non_numeric')
                status_type = 'success'
                status_message = self.texts['remove_chars_success'].format(col=col)
                self.update_status(f"Removed non-numeric characters in column '{col}'.")
            elif op_key == "op_remove_non_alpha":
                new_df[col] = new_df[col].astype(str).apply(remove_chars, mode='non_alphabetic')
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
                    new_df[col] = new_df[col].apply(fill_missing, fill_value=fill_val)
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
            else:
                status_type = 'warning'
                status_message = self.texts['not_implemented'].format(op=op_text)
                self.update_status(f"Operation '{op_text}' is not implemented.")

            if new_dataframe is not None:
                new_df = new_dataframe

            # Commit changes to history only if operation succeeded or partially succeeded
            if status_type != 'error':
                self._commit_undoable_action(old_df.copy(deep=True))
                self.dataframe = new_df

            rows_after = len(self.dataframe)
            cols_after = len(self.dataframe.columns)
            row_diff = rows_after - rows_before
            col_diff = cols_after - cols_before

            final_status_msg = f"Operation '{op_text}' finished."
            if status_type == 'success':
                final_status_msg += " (Success)"
                if row_diff != 0:
                    final_status_msg += f" Rows changed by {row_diff}."
                if col_diff != 0:
                    final_status_msg += f" Columns changed by {col_diff}."
            elif status_type == 'warning':
                final_status_msg += f" (Warning: {status_message})"
            elif status_type == 'error':
                final_status_msg += f" (Error: {status_message})"

            self.update_status(final_status_msg)

            if status_message and status_type != 'info':
                if status_type == 'success':
                    messagebox.showinfo(self.texts['success'], status_message, parent=self.root)
                elif status_type == 'warning':
                    messagebox.showwarning(self.texts['warning'], status_message, parent=self.root)
                elif status_type == 'error':
                    messagebox.showerror(self.texts['error'], status_message, parent=self.root)

            if refresh_columns:
                self.update_column_combobox(col)

        except Exception as e:
            messagebox.showerror(self.texts['error'], self.texts['operation_error'].format(error=e), parent=self.root)
            self.update_status(f"Operation '{op_text}' failed with error: {e}")

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
        if messagebox.askyesno(self.texts['warning'],
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
        suggested_name = "modified_excel." + chosen_ext

        save_path = filedialog.asksaveasfilename(
            title=self.texts['save_modified_file'],
            initialfile=suggested_name,
            defaultextension="." + chosen_ext,
            filetypes=[(self.texts['excel_files'], "*.xlsx *.xls *.csv")]
        )

        if save_path:
            try:
                if chosen_ext == "csv":
                    self.dataframe.to_csv(save_path, index=False)
                else:
                    # Default to Excel if chosen_ext is xls or xlsx
                    self.dataframe.to_excel(save_path, index=False, engine='openpyxl')
                messagebox.showinfo(self.texts['success'], self.texts['file_saved_success'].format(path=save_path))
                self.update_status(f"File saved successfully to {os.path.basename(save_path)}.")
            except Exception as e:
                messagebox.showerror(self.texts['error'], self.texts['save_error'].format(error=e))
                self.update_status(f"Error saving file: {e}")
        else:
            self.update_status("Save operation cancelled.")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelEditorApp(root)
    root.mainloop()

