# Excel Table Tools

A simple GUI application built with Python and Tkinter to perform common data cleaning and manipulation operations on Excel files.

## Features

*   Load Excel files (`.xlsx`, `.xls`).
*   Perform various operations on selected columns:
    *   Mask data (keep first 2 and last 2 characters).
    *   Mask email addresses (e.g., `us***@domain.com`).
    *   Trim leading/trailing whitespace.
    *   Split columns by delimiter (space, colon).
    *   Split surname (last word) into a new column.
    *   Change text case (UPPERCASE, lowercase, Title Case).
    *   Find and replace text.
    *   Remove specific characters.
    *   Remove non-numeric or non-alphabetic characters.
    *   Concatenate multiple columns into a new column.
    *   Extract data using regular expressions into a new column.
    *   Fill missing values (NaN, empty strings) with a specified value.
    *   Mark duplicate rows based on a column.
    *   Remove duplicate rows based on a column.
*   Save the modified data to a new Excel file.
*   Basic status logging.
*   Switchable UI language (English/Turkish).

## Requirements

*   Python 3.x
*   pandas
*   openpyxl

You can install the required libraries using pip:
```bash
pip install -r requirements.txt
```

## Usage

1.  Run the `main.py` script:
    ```bash
    python main.py
    ```
2.  Click "Browse..." to load an Excel file.
3.  Select the target column from the dropdown list.
4.  Select the desired operation from the dropdown list.
5.  Click "Apply Operation". Some operations might prompt for additional input (e.g., find/replace text, new column names).
6.  Repeat steps 3-5 for other operations as needed.
7.  Click "Save Changes" to save the modified data to a new Excel file.

## License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.
