# Excel Table Tools

A comprehensive GUI application built with Python and Tkinter for data cleaning, manipulation, and validation operations on Excel and CSV files.

## Screenshots

<img src="media/v2/main%20screen.png" alt="Main Interface" width="400"/> <img src="media/v2/russian%20page.png" alt="Multi-language Support" width="400"/>

<img src="media/v2/preview.png" alt="Operation Preview" width="400"/> <img src="media/v2/turkish page.png" alt="Turkish Interface" width="400"/>

## ✨ Key Features

- **Multi-Format Support**: Excel (`.xlsx`, `.xls`), CSV (`.csv`) input/output
- **Interactive Preview**: See changes before applying operations
- **Undo/Redo System**: Complete operation history
- **Multi-language**: English, Turkish, and Russian support
- **Data Validation**: Email, phone, date, URL validation with visual feedback
- **29 Operations**: Text processing, masking, validation, calculations

## 🚀 Quick Start

### Executables 
Download from **[Releases](https://github.com/AhmetCanArslan/ExcelTableTools/releases)** page.
I've generated these executables with pyinstaller. You can find "GenerateExecutables" in source.
(*If you're using windows, you'll get antivirus warning because app is not signed and it'll take a couple of seconds to open.*)

### Installation

1. **Clone and Setup**:
   ```bash
   git clone github.com/AhmetCanArslan/ExcelTableTools
   cd ExcelTableTools
   pip install -r requirements.txt
   ```

2. **Run Application**:
   ```bash
   python excel_table_tools.py
   ```

### Basic Workflow
1. Load your Excel or CSV file
2. Select target column and operation
3. Preview changes before applying
4. Apply operation and save results

## 📋 Available Operations

- **op_mask** - Keep first 2 and last 2 characters (e.g., `12345678` → `12****78`)
- **op_mask_email** - Protect email addresses (e.g., `user@domain.com` → `us***@domain.com`)
- **op_mask_words** - Mask individual words (e.g., `John Doe` → `Jo** D**`)
- **op_trim** - Remove leading/trailing whitespace
- **op_split_delimiter** - Split column by custom delimiter
- **op_split_surname** - Extract surname (last word) to new column
- **op_upper** - Convert text to UPPERCASE
- **op_lower** - Convert text to lowercase
- **op_title** - Convert text to Title Case
- **op_find_replace** - Find and replace text patterns
- **op_remove_specific** - Remove specific characters
- **op_remove_non_numeric** - Keep only numbers
- **op_remove_non_alpha** - Keep only letters
- **op_concatenate** - Join multiple columns with separator
- **op_extract_pattern** - Extract data using regular expressions
- **op_fill_missing** - Fill empty values with specified text
- **op_mark_duplicates** - Mark duplicate rows across columns
- **op_remove_duplicates** - Remove duplicate entries
- **op_merge_columns** - Merge columns with missing value handling
- **op_rename_column** - Rename column with validation
- **op_round_numbers** - Round numeric values to decimal places
- **op_calculate_column_constant** - Perform math operations (+, -, *, /)
- **op_create_calculated_column** - Create calculated columns from two sources
- **op_validate_email** - Validate email format with domain verification
- **op_validate_phone** - Validate phone number formats
- **op_validate_date** - Validate date formats
- **op_validate_numeric** - Validate numeric values
- **op_validate_alphanumeric** - Validate alphanumeric text
- **op_validate_url** - Validate URL addresses
- **op_distinct_group** - Create distinct group numbers for categorization

## 📄 License

This project is licensed under the GNU General Public License v3.0. You can use, modify, and distribute this software freely, but derivative work must also be licensed under GPL v3.

## 🤝 Contributing

Contributions are welcome! Please submit pull requests for any improvements.
