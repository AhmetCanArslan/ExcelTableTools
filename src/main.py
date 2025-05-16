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
PREVIEW_ROWS = 5  # Number of rows to show in preview
RESOURCES_DIR = os.path.join(project_root, 'resources')

# --- GUI Application ---
class ExcelEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.geometry("750x550")  # Increased width slightly for preview button

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

        # Remember last browsing directory
        self.last_dir = os.getcwd()

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

        self.preview_button = ttk.Button(self.ops_frame,
            text=self.texts.get('preview_button', "Preview"),
            command=self.preview_operation,
            state="disabled")  # start disabled
        self.preview_button.grid(row=2, column=1, padx=5, pady=10, sticky="ew")

        # toggle preview button when operation selection changes
        self.selected_operation.trace_add("write", self._on_operation_change)

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
        self.preview_button.config(text=self.texts['preview_button'])  # Added for preview button
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

    def browse_file(self):
        path = filedialog.askopenfilename(
            initialdir=self.last_dir,
            title=self.texts['select_excel_file'],
            filetypes=[(self.texts['excel_files'], "*.xlsx *.xls *.csv")]
        )
        if path:
            self.last_dir = os.path.dirname(path)
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
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        preview_dialog.geometry(f"{int(screen_width * 0.8)}x{int(screen_height * 0.8)}")


        main_frame = ttk.Frame(preview_dialog, padding="10")
        main_frame.pack(expand=True, fill="both")

        ttk.Label(main_frame, text=f"{op_text} - {self.texts['preview_display_title']}").pack(pady=5)

        notebook = ttk.Notebook(main_frame)
        notebook.pack(expand=True, fill="both", pady=5)

        def create_html_table(df, has_styling=False, styled_columns=None):
            import html
            
            # Convert DataFrame to HTML table with consistent styling
            html_content = "<html><body><style>table {border-collapse: collapse; width: 100%;} "
            html_content += "th, td {border: 1px solid #ddd; padding: 8px; text-align: left;} "
            html_content += "th {background-color: #f2f2f2;}</style><table>\n"
            
            # Add header row
            html_content += "<tr>"
            for col in df.columns:
                html_content += f"<th>{html.escape(str(col))}</th>"
            html_content += "</tr>\n"
            
            # Add data rows with styling if applicable
            for idx, row in df.iterrows():
                html_content += "<tr>"
                for col in df.columns:
                    cell_value = str(row[col]) if not pd.isna(row[col]) else ""
                    
                    # Apply styling if this column has validation styling
                    if has_styling and styled_columns and col in styled_columns:
                        is_invalid = styled_columns[col].iloc[idx]
                        if is_invalid:
                            html_content += f'<td style="background-color: #FFCCCC">{html.escape(cell_value)}</td>'
                        else:
                            html_content += f'<td>{html.escape(cell_value)}</td>'
                    else:
                        html_content += f'<td>{html.escape(cell_value)}</td>'
                html_content += "</tr>\n"
            
            html_content += "</table></body></html>"
            return html_content

        try:
            from tkinterweb import HtmlFrame

            # Original Data Tab with HTML
            original_tab = ttk.Frame(notebook)
            notebook.add(original_tab, text=self.texts['preview_original_data'].format(n=PREVIEW_ROWS))
            original_html = create_html_table(original_df_sample)
            original_html_frame = HtmlFrame(original_tab, messages_enabled=False)
            original_html_frame.load_html(original_html)
            original_html_frame.pack(expand=True, fill="both", padx=5, pady=5)
            # Only pack scrollbars if they exist and are not False
            if hasattr(original_html_frame, "horizontal_scrollbar") and hasattr(original_html_frame.horizontal_scrollbar, "pack"):
                original_html_frame.horizontal_scrollbar.pack(side="bottom", fill="x")
            if hasattr(original_html_frame, "vertical_scrollbar") and hasattr(original_html_frame.vertical_scrollbar, "pack"):
             original_html_frame.vertical_scrollbar.pack(side="right", fill="y")


            # Modified Data Tab with HTML (with styling if applicable)
            modified_tab = ttk.Frame(notebook)
            notebook.add(modified_tab, text=self.texts['preview_modified_data'].format(n=PREVIEW_ROWS))
            has_styling = hasattr(modified_df_sample, '_styled_columns')
            styled_columns = getattr(modified_df_sample, '_styled_columns', None) if has_styling else None
            modified_html = create_html_table(modified_df_sample, has_styling, styled_columns)
            modified_html_frame = HtmlFrame(modified_tab, messages_enabled=False)
            modified_html_frame.load_html(modified_html)
            modified_html_frame.pack(expand=True, fill="both", padx=5, pady=5)

            if hasattr(modified_html_frame, "horizontal_scrollbar") and hasattr(modified_html_frame.horizontal_scrollbar, "pack"):
                modified_html_frame.horizontal_scrollbar.pack(side="bottom", fill="x")
            if hasattr(modified_html_frame, "vertical_scrollbar") and hasattr(modified_html_frame.vertical_scrollbar, "pack"):
                modified_html_frame.vertical_scrollbar.pack(side="right", fill="y")

        except ImportError:
            # Fallback to text representation if tkinterweb is not available
            original_tab = ttk.Frame(notebook)
            notebook.add(original_tab, text=self.texts['preview_original_data'].format(n=PREVIEW_ROWS))
            
            # Add both scrollbars to Text widget
            original_text_area = tk.Text(original_tab, wrap=tk.NONE, height=10)
            original_text_area.insert(tk.END, original_df_sample.to_string())
            original_text_area.config(state='disabled')
            # Place in a frame for scrollbars
            text_frame = ttk.Frame(original_tab)
            text_frame.pack(expand=True, fill="both", padx=5, pady=5)
            original_text_area.pack(in_=text_frame, side="left", expand=True, fill="both")
            original_v_scroll = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=original_text_area.yview)
            original_v_scroll.pack(side="right", fill="y")
            original_h_scroll = ttk.Scrollbar(original_tab, orient=tk.HORIZONTAL, command=original_text_area.xview)
            original_h_scroll.pack(fill=tk.X, side=tk.BOTTOM)
            original_text_area.config(xscrollcommand=original_h_scroll.set, yscrollcommand=original_v_scroll.set)

            modified_tab = ttk.Frame(notebook)
            notebook.add(modified_tab, text=self.texts['preview_modified_data'].format(n=PREVIEW_ROWS))
            # Add both scrollbars to ScrolledText widget
            mod_text_frame = ttk.Frame(modified_tab)
            mod_text_frame.pack(expand=True, fill="both", padx=5, pady=5)
            modified_text_area = scrolledtext.ScrolledText(mod_text_frame, wrap=tk.NONE, height=10)
            modified_text_area.insert(tk.END, modified_df_sample.to_string())
            modified_text_area.insert(tk.END, "\n\n" + self.texts.get('validation_preview_note', 
                "Note: Install 'tkinterweb' package to see colored validation in preview."))
            modified_text_area.config(state='disabled')
            modified_text_area.pack(side="left", expand=True, fill="both")
            modified_v_scroll = ttk.Scrollbar(mod_text_frame, orient=tk.VERTICAL, command=modified_text_area.yview)
            modified_v_scroll.pack(side="right", fill="y")
            modified_h_scroll = ttk.Scrollbar(modified_tab, orient=tk.HORIZONTAL, command=modified_text_area.xview)
            modified_h_scroll.pack(fill=tk.X, side=tk.BOTTOM)
            modified_text_area.config(xscrollcommand=modified_h_scroll.set, yscrollcommand=modified_v_scroll.set)
        
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
        self.update_status(
            self.texts.get('preview_output_status', "Output preview displayed.")
        )

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
                new_df[col] = new_df[col].astype(str).apply(mask_data, mode='email', column_name=col)
                status_type = 'success'
                status_message = self.texts['email_masked_success'].format(col=col)
                self.update_status(f"Email masking applied to column '{col}'.")
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
                col2 = simpledialog.askstring(self.texts['input_needed'], self.texts['select_second_column_calc'], parent=self.root)
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
                except Exception as e:
                    messagebox.showerror(
                        self.texts['error'], 
                        self.texts['operation_error'].format(error=e), 
                        parent=self.root
                    )
                    self.update_status(f"Validation operation failed: {e}")
                    
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
        suggested_name = "modified_excel." + chosen_ext

        save_path = filedialog.asksaveasfilename(
            initialdir=self.last_dir,
            title=self.texts['save_modified_file'],
            initialfile=suggested_name,
            defaultextension="." + chosen_ext,
            filetypes=[(self.texts['excel_files'], "*.xlsx *.xls *.csv")]
        )

        if save_path:
            self.last_dir = os.path.dirname(save_path)
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