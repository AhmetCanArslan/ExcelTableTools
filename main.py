import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

# Import operations
from operations.masking import mask_data
from operations.trimming import trim_spaces
from operations.splitting import apply_split_surname, apply_split_by_delimiter

# --- Language Translations ---
LANGUAGES = {
    'en': {
        'title': "Excel Table Tools",
        'file_selection': "File Selection",
        'excel_file': "Excel File:",
        'browse': "Browse...",
        'operations': "Operations",
        'column': "Column:",
        'operation': "Operation:",
        'apply_operation': "Apply Operation",
        'save_changes': "Save Changes",
        'select_excel_file': "Select Excel File",
        'excel_files': "Excel files",
        'success': "Success",
        'error': "Error",
        'warning': "Warning",
        'info': "Info",
        'loaded_successfully': "Loaded '{filename}' successfully.",
        'error_loading': "Failed to load Excel file.\nError: {error}",
        'no_file': "Please load an Excel file first.",
        'no_column': "Please select a column.",
        'no_operation': "Please select an operation.",
        'masked_success': "Masked data in column '{col}'.",
        'trimmed_success': "Trimmed spaces in column '{col}'.",
        'split_success': "Split column '{col}' by '{delimiter}' into {count} new columns.",
        'surname_split_success': "Split surname from column '{col}' into new column '{new_col}'.",
        'split_warning_delimiter_not_found': "The delimiter '{delimiter}' was not found in column '{col}'. No changes made.",
        'column_not_found': "Column '{col}' not found.",
        'operation_error': "An error occurred during the operation:\n{error}",
        'not_implemented': "Operation '{op}' is not yet implemented.",
        'no_data_to_save': "No data to save. Load and modify a file first.",
        'save_modified_file': "Save Modified Excel File",
        'file_saved_success': "File saved successfully to:\n{path}",
        'save_error': "Failed to save the file.\nError: {error}",
        'change_language': "Türkçe", # Button text shows the *other* language
        'op_mask': "Mask Column (Keep 2+2)",
        'op_trim': "Trim Spaces",
        'op_split_space': "Split Column (Space)",
        'op_split_colon': "Split Column (:)",
        'op_split_surname': "Split Surname (Last Word)"
    },
    'tr': {
        'title': "Excel Tablo Araçları",
        'file_selection': "Dosya Seçimi",
        'excel_file': "Excel Dosyası:",
        'browse': "Gözat...",
        'operations': "İşlemler",
        'column': "Sütun:",
        'operation': "İşlem:",
        'apply_operation': "İşlemi Uygula",
        'save_changes': "Değişiklikleri Kaydet",
        'select_excel_file': "Excel Dosyası Seç",
        'excel_files': "Excel dosyaları",
        'success': "Başarılı",
        'error': "Hata",
        'warning': "Uyarı",
        'info': "Bilgi",
        'loaded_successfully': "'{filename}' başarıyla yüklendi.",
        'error_loading': "Excel dosyası yüklenemedi.\nHata: {error}",
        'no_file': "Lütfen önce bir Excel dosyası yükleyin.",
        'no_column': "Lütfen bir sütun seçin.",
        'no_operation': "Lütfen bir işlem seçin.",
        'masked_success': "'{col}' sütunundaki veriler maskelendi.",
        'trimmed_success': "'{col}' sütunundaki boşluklar temizlendi.",
        'split_success': "'{col}' sütunu '{delimiter}' ile {count} yeni sütuna bölündü.",
        'surname_split_success': "Soyadı '{col}' sütunundan ayırıp '{new_col}' sütununa yazıldı.",
        'split_warning_delimiter_not_found': "'{delimiter}' ayıracı '{col}' sütununda bulunamadı. Değişiklik yapılmadı.",
        'column_not_found': "'{col}' sütunu bulunamadı.",
        'operation_error': "İşlem sırasında bir hata oluştu:\n{error}",
        'not_implemented': "'{op}' işlemi henüz uygulanmadı.",
        'no_data_to_save': "Kaydedilecek veri yok. Önce bir dosya yükleyin ve değiştirin.",
        'save_modified_file': "Değiştirilmiş Excel Dosyasını Kaydet",
        'file_saved_success': "Dosya başarıyla şuraya kaydedildi:\n{path}",
        'save_error': "Dosya kaydedilemedi.\nHata: {error}",
        'change_language': "English", # Button text shows the *other* language
        'op_mask': "Sütunu Maskele (2+2 Sakla)",
        'op_trim': "Boşlukları Temizle",
        'op_split_space': "Sütunu Böl (Boşluk)",
        'op_split_colon': "Sütunu Böl (:)",
        'op_split_surname': "Soyadını Ayır (Son Kelime)"
    }
}

# --- GUI Application ---
class ExcelEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.geometry("600x400")

        self.file_path = tk.StringVar()
        self.selected_column = tk.StringVar()
        self.selected_operation = tk.StringVar()
        self.dataframe = None
        self.current_lang = 'en' # Default language
        self.texts = LANGUAGES[self.current_lang]

        # --- Top Frame for Language Button ---
        top_frame = ttk.Frame(root)
        top_frame.pack(fill="x", padx=10, pady=(5, 0))
        self.lang_button = ttk.Button(top_frame, text=self.texts['change_language'], command=self.toggle_language)
        self.lang_button.pack(side="right")

        # --- File Selection ---
        self.file_frame = ttk.LabelFrame(root, text=self.texts['file_selection'])
        self.file_frame.pack(padx=10, pady=10, fill="x")

        self.file_label = ttk.Label(self.file_frame, text=self.texts['excel_file'])
        self.file_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.file_entry = ttk.Entry(self.file_frame, textvariable=self.file_path, width=50, state="readonly")
        self.file_entry.grid(row=0, column=1, padx=5, pady=5)
        self.browse_button = ttk.Button(self.file_frame, text=self.texts['browse'], command=self.browse_file)
        self.browse_button.grid(row=0, column=2, padx=5, pady=5)

        # --- Operations ---
        self.ops_frame = ttk.LabelFrame(root, text=self.texts['operations'])
        self.ops_frame.pack(padx=10, pady=10, fill="x")

        self.column_label = ttk.Label(self.ops_frame, text=self.texts['column'])
        self.column_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.column_combobox = ttk.Combobox(self.ops_frame, textvariable=self.selected_column, state="disabled", width=30)
        self.column_combobox.grid(row=0, column=1, padx=5, pady=5)

        self.operation_label = ttk.Label(self.ops_frame, text=self.texts['operation'])
        self.operation_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.operation_combobox = ttk.Combobox(self.ops_frame, textvariable=self.selected_operation, state="disabled", width=30)
        self.operation_combobox.grid(row=1, column=1, padx=5, pady=5)

        self.apply_button = ttk.Button(self.ops_frame, text=self.texts['apply_operation'], command=self.apply_operation)
        self.apply_button.grid(row=2, column=0, columnspan=2, padx=5, pady=10)

        # --- Save ---
        save_frame = ttk.Frame(root)
        save_frame.pack(padx=10, pady=10, fill="x")
        self.save_button = ttk.Button(save_frame, text=self.texts['save_changes'], command=self.save_file)
        self.save_button.pack(side="right", padx=5)

        self.update_ui_language() # Set initial language

    def update_ui_language(self):
        """Updates all UI elements with text from the current language dictionary."""
        self.texts = LANGUAGES[self.current_lang]
        self.root.title(self.texts['title'])
        self.lang_button.config(text=self.texts['change_language'])

        # Update frame labels
        self.file_frame.config(text=self.texts['file_selection'])
        self.ops_frame.config(text=self.texts['operations'])

        # Update labels and buttons
        self.file_label.config(text=self.texts['excel_file'])
        self.browse_button.config(text=self.texts['browse'])
        self.column_label.config(text=self.texts['column'])
        self.operation_label.config(text=self.texts['operation'])
        self.apply_button.config(text=self.texts['apply_operation'])
        self.save_button.config(text=self.texts['save_changes'])

        # Update operation combobox values (store original keys)
        self.operation_keys = ["op_mask", "op_trim", "op_split_space", "op_split_colon", "op_split_surname"]
        translated_ops = [self.texts[key] for key in self.operation_keys]
        current_selection_text = self.selected_operation.get()
        self.operation_combobox['values'] = translated_ops

        # Try to re-select the operation based on the new language text
        if current_selection_text:
            try:
                # Find the key corresponding to the old text
                current_key = None
                old_lang = 'tr' if self.current_lang == 'en' else 'en'
                for key, text in LANGUAGES[old_lang].items():
                    if text == current_selection_text and key in self.operation_keys:
                        current_key = key
                        break
                # Set the new text based on the found key
                if current_key:
                    self.selected_operation.set(self.texts[current_key])
                else:
                    self.selected_operation.set("") # Clear if mapping fails
            except Exception:
                 self.selected_operation.set("") # Clear on any error
        else:
            self.selected_operation.set("")

    def toggle_language(self):
        """Switches the application language between English and Turkish."""
        self.current_lang = 'tr' if self.current_lang == 'en' else 'en'
        self.update_ui_language()

    def get_operation_key(self, translated_op_text):
        """Gets the internal operation key from the translated text."""
        for key in self.operation_keys:
            if self.texts[key] == translated_op_text:
                return key
        return None # Should not happen if list is correct

    def browse_file(self):
        """Opens a file dialog to select an Excel file."""
        path = filedialog.askopenfilename(
            title=self.texts['select_excel_file'],
            filetypes=[(self.texts['excel_files'], "*.xlsx *.xls")]
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
            if path.endswith('.xlsx'):
                self.dataframe = pd.read_excel(path, engine='openpyxl')
            else: # For .xls
                 self.dataframe = pd.read_excel(path)

            # Update column combobox
            self.column_combobox['values'] = list(self.dataframe.columns)
            self.column_combobox.config(state="readonly")
            if self.dataframe.columns.any():
                 self.selected_column.set(self.dataframe.columns[0]) # Default to first column
            self.operation_combobox.config(state="readonly") # Enable operations dropdown
            messagebox.showinfo(self.texts['success'], self.texts['loaded_successfully'].format(filename=os.path.basename(path)))
        except Exception as e:
            messagebox.showerror(self.texts['error'], self.texts['error_loading'].format(error=e))
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
            messagebox.showwarning(self.texts['warning'], self.texts['no_file'])
            return

        col = self.selected_column.get()
        op_text = self.selected_operation.get()
        op_key = self.get_operation_key(op_text)

        if not col:
            messagebox.showwarning(self.texts['warning'], self.texts['no_column'])
            return
        if not op_key:
            messagebox.showwarning(self.texts['warning'], self.texts['no_operation'])
            return

        try:
            new_dataframe = None
            status_type = None
            status_message = ""

            if op_key == "op_mask":
                # Apply directly
                self.dataframe[col] = self.dataframe[col].astype(str).apply(mask_data)
                status_type = 'success'
                status_message = self.texts['masked_success'].format(col=col)
            elif op_key == "op_trim":
                 # Apply directly
                 self.dataframe[col] = self.dataframe[col].astype(str).apply(trim_spaces)
                 status_type = 'success'
                 status_message = self.texts['trimmed_success'].format(col=col)
            elif op_key == "op_split_space":
                new_dataframe, (status_type, status_message) = apply_split_by_delimiter(self.dataframe, col, ' ', self.texts)
            elif op_key == "op_split_colon":
                 new_dataframe, (status_type, status_message) = apply_split_by_delimiter(self.dataframe, col, ':', self.texts)
            elif op_key == "op_split_surname":
                new_dataframe, (status_type, status_message) = apply_split_surname(self.dataframe, col, self.texts)
            else:
                messagebox.showwarning(self.texts['warning'], self.texts['not_implemented'].format(op=op_text))
                return

            # Update dataframe if a new one was returned by the operation function
            if new_dataframe is not None:
                self.dataframe = new_dataframe

            # Show status message
            if status_type == 'success':
                messagebox.showinfo(self.texts['success'], status_message)
            elif status_type == 'warning':
                messagebox.showwarning(self.texts['warning'], status_message)
            elif status_type == 'error': # Should ideally not happen here if checks are done in functions
                messagebox.showerror(self.texts['error'], status_message)

            # Refresh column list if dataframe structure might have changed
            if op_key in ["op_split_space", "op_split_colon", "op_split_surname"]:
                self.column_combobox['values'] = list(self.dataframe.columns)
                # Try to keep the original column selected if it still exists
                if col in self.dataframe.columns:
                    self.selected_column.set(col)
                elif self.dataframe.columns.any():
                    self.selected_column.set(self.dataframe.columns[0])
                else:
                    self.selected_column.set("")

        except Exception as e:
            # Catch any unexpected errors during operation application
            messagebox.showerror(self.texts['error'], self.texts['operation_error'].format(error=e))
            # Optionally: Reload the original dataframe state if an error occurs
            # self.load_excel() # Or store original df state before applying

    def save_file(self):
        """Saves the modified DataFrame to a new Excel file."""
        if self.dataframe is None:
            messagebox.showwarning(self.texts['warning'], self.texts['no_data_to_save'])
            return

        original_path = self.file_path.get()
        if original_path:
            dir_name = os.path.dirname(original_path)
            base_name = os.path.basename(original_path)
            name, ext = os.path.splitext(base_name)
            suggested_name = os.path.join(dir_name, f"{name}_modified{ext}")
        else:
            suggested_name = "modified_excel.xlsx"

        save_path = filedialog.asksaveasfilename(
            title=self.texts['save_modified_file'],
            initialfile=suggested_name,
            defaultextension=".xlsx",
            filetypes=[(self.texts['excel_files'], "*.xlsx *.xls")]
        )

        if save_path:
            try:
                if save_path.endswith('.xlsx'):
                    self.dataframe.to_excel(save_path, index=False, engine='openpyxl')
                else: # For .xls
                    self.dataframe.to_excel(save_path, index=False)
                messagebox.showinfo(self.texts['success'], self.texts['file_saved_success'].format(path=save_path))
            except Exception as e:
                messagebox.showerror(self.texts['error'], self.texts['save_error'].format(error=e))


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelEditorApp(root)
    root.mainloop()
