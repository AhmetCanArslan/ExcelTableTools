import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import os
import re

# Import operations
from operations.masking import mask_data
from operations.trimming import trim_spaces
from operations.splitting import apply_split_surname, apply_split_by_delimiter
from operations.case_change import change_case
from operations.find_replace import find_replace
from operations.remove_chars import remove_chars
from operations.concatenate import apply_concatenate
from operations.extract_pattern import apply_extract_pattern
from operations.fill_missing import fill_missing
from operations.duplicates import apply_mark_duplicates, apply_remove_duplicates

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
        'change_language': "Türkçe",
        'op_mask': "Mask Column (Keep 2+2)",
        'op_trim': "Trim Spaces",
        'op_split_space': "Split Column (Space)",
        'op_split_colon': "Split Column (:)",
        'op_split_surname': "Split Surname (Last Word)",
        'op_upper': "Change Case: UPPERCASE",
        'op_lower': "Change Case: lowercase",
        'op_title': "Change Case: Title Case",
        'op_find_replace': "Find and Replace...",
        'op_remove_specific': "Remove Specific Characters...",
        'op_remove_non_numeric': "Remove Non-numeric Chars",
        'op_remove_non_alpha': "Remove Non-alphabetic Chars",
        'op_concatenate': "Concatenate Columns...",
        'op_extract_pattern': "Extract with Regex...",
        'op_fill_missing': "Fill Missing Values...",
        'op_mark_duplicates': "Mark Duplicate Rows (by Column)",
        'op_remove_duplicates': "Remove Duplicate Rows (by Column)",
        'case_change_success': "Changed case in column '{col}' to {case_type}.",
        'find_replace_success': "Performed find/replace in column '{col}'.",
        'remove_chars_success': "Removed characters in column '{col}'.",
        'concatenate_success': "Concatenated {count} columns into new column '{new_col}'.",
        'extract_success': "Extracted pattern from '{col}' into new column '{new_col}'.",
        'fill_missing_success': "Filled missing values in column '{col}'.",
        'duplicates_marked_success': "Marked duplicate rows based on '{col}' in new column '{new_col}'.",
        'duplicates_removed_success': "Removed {count} duplicate rows based on column '{col}'.",
        'regex_error': "Invalid Regular Expression: {error}",
        'input_needed': "Input Needed",
        'enter_find_text': "Enter text to find:",
        'enter_replace_text': "Enter text to replace with:",
        'enter_chars_to_remove': "Enter characters to remove:",
        'enter_fill_value': "Enter value to fill missing cells with:",
        'enter_regex_pattern': "Enter Regex pattern (e.g., \\d+):",
        'enter_new_col_name': "Enter name for the new column:",
        'select_columns_concat': "Select columns to concatenate (use Ctrl+Click):",
        'enter_separator': "Enter separator for concatenation:",
        'no_columns_selected': "No columns selected for concatenation.",
        'invalid_column_name': "Invalid or empty new column name.",
        'column_already_exists': "Column '{name}' already exists. Please choose a different name.",
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
        'change_language': "English",
        'op_mask': "Sütunu Maskele (2+2 Sakla)",
        'op_trim': "Boşlukları Temizle",
        'op_split_space': "Sütunu Böl (Boşluk)",
        'op_split_colon': "Sütunu Böl (:)",
        'op_split_surname': "Soyadını Ayır (Son Kelime)",
        'op_upper': "Büyük/Küçük Harf: TÜMÜ BÜYÜK",
        'op_lower': "Büyük/Küçük Harf: tümü küçük",
        'op_title': "Büyük/Küçük Harf: Baş Harfler Büyük",
        'op_find_replace': "Bul ve Değiştir...",
        'op_remove_specific': "Belirli Karakterleri Kaldır...",
        'op_remove_non_numeric': "Sayısal Olmayanları Kaldır",
        'op_remove_non_alpha': "Alfabetik Olmayanları Kaldır",
        'op_concatenate': "Sütunları Birleştir...",
        'op_extract_pattern': "Regex ile Çıkart...",
        'op_fill_missing': "Boş Değerleri Doldur...",
        'op_mark_duplicates': "Yinelenen Satırları İşaretle (Sütuna Göre)",
        'op_remove_duplicates': "Yinelenen Satırları Kaldır (Sütuna Göre)",
        'case_change_success': "'{col}' sütunundaki harf durumu {case_type} olarak değiştirildi.",
        'find_replace_success': "'{col}' sütununda bul/değiştir yapıldı.",
        'remove_chars_success': "'{col}' sütunundaki karakterler kaldırıldı.",
        'concatenate_success': "{count} sütun birleştirilerek '{new_col}' sütunu oluşturuldu.",
        'extract_success': "'{col}' sütunundan desen '{new_col}' sütununa çıkartıldı.",
        'fill_missing_success': "'{col}' sütunundaki boş değerler dolduruldu.",
        'duplicates_marked_success': "'{col}' sütununa göre yinelenen satırlar '{new_col}' sütununda işaretlendi.",
        'duplicates_removed_success': "'{col}' sütununa göre {count} yinelenen satır kaldırıldı.",
        'regex_error': "Geçersiz Düzenli İfade: {error}",
        'input_needed': "Girdi Gerekiyor",
        'enter_find_text': "Bulunacak metni girin:",
        'enter_replace_text': "Yerine konulacak metni girin:",
        'enter_chars_to_remove': "Kaldırılacak karakterleri girin:",
        'enter_fill_value': "Boş hücrelerin doldurulacağı değeri girin:",
        'enter_regex_pattern': "Regex desenini girin (örn: \\d+):",
        'enter_new_col_name': "Yeni sütun için ad girin:",
        'select_columns_concat': "Birleştirilecek sütunları seçin (Ctrl+Tık kullanın):",
        'enter_separator': "Birleştirme için ayırıcı girin:",
        'no_columns_selected': "Birleştirme için sütun seçilmedi.",
        'invalid_column_name': "Geçersiz veya boş yeni sütun adı.",
        'column_already_exists': "'{name}' sütunu zaten var. Lütfen farklı bir ad seçin.",
    }
}

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
        self.root.geometry("700x500")

        self.file_path = tk.StringVar()
        self.selected_column = tk.StringVar()
        self.selected_operation = tk.StringVar()
        self.dataframe = None
        self.current_lang = 'en'
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

        self.update_ui_language()

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
        self.save_button.config(text=self.texts['save_changes'])

        self.operation_keys = [
            "op_mask", "op_trim", "op_split_space", "op_split_colon", "op_split_surname",
            "op_upper", "op_lower", "op_title",
            "op_find_replace", "op_remove_specific", "op_remove_non_numeric", "op_remove_non_alpha",
            "op_concatenate", "op_extract_pattern", "op_fill_missing",
            "op_mark_duplicates", "op_remove_duplicates"
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

    def toggle_language(self):
        self.current_lang = 'tr' if self.current_lang == 'en' else 'en'
        self.update_ui_language()

    def get_operation_key(self, translated_op_text):
        for key in self.operation_keys:
            if self.texts[key] == translated_op_text:
                return key
        return None

    def browse_file(self):
        path = filedialog.askopenfilename(
            title=self.texts['select_excel_file'],
            filetypes=[(self.texts['excel_files'], "*.xlsx *.xls")]
        )
        if path:
            self.file_path.set(path)
            self.load_excel()

    def load_excel(self):
        path = self.file_path.get()
        if not path:
            return
        try:
            if path.endswith('.xlsx'):
                self.dataframe = pd.read_excel(path, engine='openpyxl')
            else:
                self.dataframe = pd.read_excel(path)

            self.column_combobox['values'] = list(self.dataframe.columns)
            self.column_combobox.config(state="readonly")
            if self.dataframe.columns.any():
                self.selected_column.set(self.dataframe.columns[0])
            self.operation_combobox.config(state="readonly")
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

    def apply_operation(self):
        if self.dataframe is None:
            messagebox.showwarning(self.texts['warning'], self.texts['no_file'])
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
                return
            self.apply_remove_duplicates_ui(col)
            return

        if not col:
            messagebox.showwarning(self.texts['warning'], self.texts['no_column'])
            return
        if not op_key:
            messagebox.showwarning(self.texts['warning'], self.texts['no_operation'])
            return

        try:
            new_dataframe = None
            status_type = 'info'
            status_message = ""
            refresh_columns = False

            if op_key == "op_mask":
                self.dataframe[col] = self.dataframe[col].astype(str).apply(mask_data)
                status_type = 'success'
                status_message = self.texts['masked_success'].format(col=col)
            elif op_key == "op_trim":
                self.dataframe[col] = self.dataframe[col].astype(str).apply(trim_spaces)
                status_type = 'success'
                status_message = self.texts['trimmed_success'].format(col=col)
            elif op_key == "op_split_space":
                new_dataframe, (status_type, status_message) = apply_split_by_delimiter(self.dataframe, col, ' ', self.texts)
                refresh_columns = True
            elif op_key == "op_split_colon":
                new_dataframe, (status_type, status_message) = apply_split_by_delimiter(self.dataframe, col, ':', self.texts)
                refresh_columns = True
            elif op_key == "op_split_surname":
                new_dataframe, (status_type, status_message) = apply_split_surname(self.dataframe, col, self.texts)
                refresh_columns = True
            elif op_key == "op_upper":
                self.dataframe[col] = self.dataframe[col].astype(str).apply(change_case, case_type='upper')
                status_type = 'success'
                status_message = self.texts['case_change_success'].format(col=col, case_type='UPPERCASE')
            elif op_key == "op_lower":
                self.dataframe[col] = self.dataframe[col].astype(str).apply(change_case, case_type='lower')
                status_type = 'success'
                status_message = self.texts['case_change_success'].format(col=col, case_type='lowercase')
            elif op_key == "op_title":
                self.dataframe[col] = self.dataframe[col].astype(str).apply(change_case, case_type='title')
                status_type = 'success'
                status_message = self.texts['case_change_success'].format(col=col, case_type='Title Case')
            elif op_key == "op_find_replace":
                find_text = self.get_input('input_needed', 'enter_find_text')
                if find_text is not None:
                    replace_text = self.get_input('input_needed', 'enter_replace_text')
                    if replace_text is not None:
                        self.dataframe[col] = self.dataframe[col].astype(str).apply(find_replace, find_text=find_text, replace_text=replace_text)
                        status_type = 'success'
                        status_message = self.texts['find_replace_success'].format(col=col)
            elif op_key == "op_remove_specific":
                chars = self.get_input('input_needed', 'enter_chars_to_remove')
                if chars:
                    self.dataframe[col] = self.dataframe[col].astype(str).apply(remove_chars, mode='specific', chars_to_remove=chars)
                    status_type = 'success'
                    status_message = self.texts['remove_chars_success'].format(col=col)
            elif op_key == "op_remove_non_numeric":
                self.dataframe[col] = self.dataframe[col].astype(str).apply(remove_chars, mode='non_numeric')
                status_type = 'success'
                status_message = self.texts['remove_chars_success'].format(col=col)
            elif op_key == "op_remove_non_alpha":
                self.dataframe[col] = self.dataframe[col].astype(str).apply(remove_chars, mode='non_alphabetic')
                status_type = 'success'
                status_message = self.texts['remove_chars_success'].format(col=col)
            elif op_key == "op_extract_pattern":
                pattern = self.get_input('input_needed', 'enter_regex_pattern')
                if pattern:
                    try:
                        re.compile(pattern)
                        new_col_base = f"{col}_extracted"
                        new_col_name = self.get_new_column_name(new_col_base)
                        if new_col_name:
                            new_dataframe, (status_type, status_message) = apply_extract_pattern(self.dataframe, col, new_col_name, pattern, self.texts)
                            refresh_columns = True
                    except re.error as e:
                        status_type = 'error'
                        status_message = self.texts['regex_error'].format(error=e)
            elif op_key == "op_fill_missing":
                fill_val = self.get_input('input_needed', 'enter_fill_value')
                if fill_val is not None:
                    self.dataframe[col] = self.dataframe[col].apply(fill_missing, fill_value=fill_val)
                    status_type = 'success'
                    status_message = self.texts['fill_missing_success'].format(col=col)
            elif op_key == "op_mark_duplicates":
                new_col_base = f"{col}_is_duplicate"
                new_col_name = self.get_new_column_name(new_col_base)
                if new_col_name:
                    new_dataframe, (status_type, status_message) = apply_mark_duplicates(self.dataframe, col, new_col_name, self.texts)
                    refresh_columns = True
            else:
                status_type = 'warning'
                status_message = self.texts['not_implemented'].format(op=op_text)

            if new_dataframe is not None:
                self.dataframe = new_dataframe

            if status_message:
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

        try:
            new_dataframe, (status_type, status_message) = apply_concatenate(self.dataframe, cols_to_concat, new_col_name, separator, self.texts)
            self.dataframe = new_dataframe
            if status_type == 'success':
                messagebox.showinfo(self.texts['success'], status_message, parent=self.root)
                self.update_column_combobox(new_col_name)
            else:
                messagebox.showerror(self.texts['error'], status_message, parent=self.root)
        except Exception as e:
            messagebox.showerror(self.texts['error'], self.texts['operation_error'].format(error=e), parent=self.root)

    def apply_remove_duplicates_ui(self, col):
        if messagebox.askyesno(self.texts['warning'],
                               f"This will permanently remove rows based on duplicates in column '{col}'.\nAre you sure?",
                               parent=self.root):
            try:
                new_dataframe, (status_type, status_message) = apply_remove_duplicates(self.dataframe, col, self.texts)
                self.dataframe = new_dataframe
                if status_type == 'success':
                    messagebox.showinfo(self.texts['success'], status_message, parent=self.root)
                    self.update_column_combobox(col)
                else:
                    messagebox.showerror(self.texts['error'], status_message, parent=self.root)
            except Exception as e:
                messagebox.showerror(self.texts['error'], self.texts['operation_error'].format(error=e), parent=self.root)

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
                else:
                    self.dataframe.to_excel(save_path, index=False)
                messagebox.showinfo(self.texts['success'], self.texts['file_saved_success'].format(path=save_path))
            except Exception as e:
                messagebox.showerror(self.texts['error'], self.texts['save_error'].format(error=e))


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelEditorApp(root)
    root.mainloop()
