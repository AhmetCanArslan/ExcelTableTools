# Excel Table Tools

A simple GUI application built with Python and Tkinter to perform common data cleaning and manipulation operations on Excel files, CSV files, and other tabular data formats.

## Screenshots

*English main application window*

<img src="media/1.png" alt="Screenshot 1" width="400"/>

*Example: Column selection and operations*

![Screenshot 2](media/2.png)

*Example: Operation selection*

![Screenshot 3](media/3.png)

*Status Log*

![Screenshot 4](media/4.png)

*Turkish main application window*

<img src="media/5.png" alt="Screenshot 5" width="400"/>


## Features

*   **Load and save multiple file formats**:
    *   Load Excel files (`.xlsx`, `.xls`) and CSV files (`.csv`).
    *   Save to various formats: Excel (`.xlsx`, `.xls`), CSV (`.csv`), JSON (`.json`), HTML (`.html`), and Markdown (`.md`).
*   **Perform various operations on selected columns**:
    *   **Data masking**:
        *   Mask data (keep first 2 and last 2 characters)
        *   Mask email addresses (e.g., `us***@domain.com`)
        *   Mask words (Keep 2 letters per word)
    *   **Text cleaning**:
        *   Trim leading/trailing whitespace
        *   Change text case (UPPERCASE, lowercase, Title Case)
        *   Find and replace text
        *   Remove specific characters
        *   Remove non-numeric or non-alphabetic characters
    *   **Column operations**:
        *   Split columns by delimiter (space, colon)
        *   Split surname (last word) into a new column
        *   Concatenate multiple columns into a new column
        *   Merge columns with customizable handling of missing values
        *   Rename columns
    *   **Data extraction and transformation**:
        *   Extract data using regular expressions into a new column
        *   Fill missing values (NaN, empty strings) with a specified value
    *   **Row operations**:
        *   Mark duplicate rows based on a column
        *   Remove duplicate rows based on a column
    *   **Numeric operations**:
        *   Round numbers to specified decimal places
        *   Perform calculations on columns with constants (+, -, *, /)
        *   Create calculated columns from two existing columns
*   **Interactive workflow**:
    *   **Preview** operations before applying them to see the effect on your data
    *   **Undo/Redo** functionality for all operations
    *   **Refresh** option to reset the application to its initial state
*   **User interface features**:
    *   Status log to track operations and changes
    *   Bilingual interface (English/Turkish)
    *   Intuitive user interface with descriptive operation names

## Requirements

*   Python 3.x
*   pandas
*   openpyxl
*   tkinter
*   tabulate (for Markdown export)

You can install the required libraries using pip:
```bash
pip install -r requirements.txt
```

## Usage
1. Install requirements:
   ```bash
   pip install -r requirements.txt
   ```
2. Run the `main.py` script:
   ```bash
   python main.py
   ```
3. Click "Browse..." to load an Excel (xlsx, xls) or CSV file.
4. Select the target column from the dropdown list.
5. Select the desired operation from the dropdown list.
6. (Optional) Click "Preview" to see the effect of the operation before applying it.
7. Click "Apply Operation". Some operations might prompt for additional input (e.g., find/replace text, new column names).
8. Repeat steps 4-7 for other operations as needed.
9. Use the "Undo" or "Redo" buttons if needed to revert or reapply operations.
10. Select the desired output format from the dropdown menu next to the Save button.
11. Click "Save Changes" to save the modified data to a new file in the selected format.

## Creating an Executable

A script is included to build a standalone executable using PyInstaller:

```bash
chmod +x MKEXEC.sh
./MKEXEC.sh
```

The executable will be created in the `dist` directory.