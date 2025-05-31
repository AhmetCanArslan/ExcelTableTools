# Excel Table Tools

A comprehensive GUI application built with Python and Tkinter to perform advanced data cleaning, manipulation, and validation operations on Excel files, CSV files, and other tabular data formats.

## Screenshots

<img src="media/v2/main%20screen.png" alt="Main Interface" width="400"/> <img src="media/v2/russian%20page.png" alt="Multi-language Support" width="400"/>

<img src="media/v2/preview.png" alt="Operation Preview" width="400"/> <img src="media/v2/turkish page.png" alt="Turkish Interface" width="400"/>
<img src="media/v2/success%20alert.png" alt="Success Feedback" width="400"/>

## ✨ Key Features

### 📁 **Multi-Format File Support**
- **Input**: Excel files (`.xlsx`, `.xls`), CSV files (`.csv`)
- **Output**: Excel (`.xlsx`, `.xls`), CSV (`.csv`), JSON (`.json`), HTML (`.html`), Markdown (`.md`)
- **Smart Preview**: Load and preview different sections of large files (head, middle, tail)

### 🔧 **Comprehensive Data Operations**

#### **Data Security & Privacy**
- **Mask Data**: Keep first 2 and last 2 characters (e.g., `12345678` → `12****78`)
- **Mask Email**: Protect email addresses (e.g., `user@domain.com` → `us***@domain.com`)
- **Mask Words**: Protect individual words (e.g., `Ahmet Can Arslan` → `Ah*** C** Ar****`)

#### **Text Processing**
- **Trim Spaces**: Remove leading/trailing whitespace
- **Case Changes**: UPPERCASE, lowercase, Title Case
- **Find & Replace**: Search and replace text with pattern support
- **Character Removal**: 
  - Remove specific characters
  - Remove non-numeric characters
  - Remove non-alphabetic characters

#### **Advanced Column Operations**
- **Split Columns**: 
  - Split by custom delimiters (space, comma, colon, etc.)
  - Extract surnames (last word) into new columns
- **Column Management**:
  - Concatenate multiple columns with custom separators
  - Merge columns with advanced missing value handling
  - Rename columns with validation
- **Numeric Operations**:
  - Round numbers to specified decimal places
  - Calculate with constants (+, -, *, /)
  - Create calculated columns from two source columns
- **Advanced Features**:
  - Create distinct group numbers for categorization
  - Extract data using regular expressions with capture groups

#### **Data Validation & Quality**
- **Comprehensive Validation**:
  - Email address format validation with **domain verification using Public Suffix List (PSL)**
  - Phone number format validation
  - Date format validation
  - Numeric value validation
  - Alphanumeric text validation
  - URL address validation
- **Visual Feedback**: Invalid data highlighted in red for easy identification
- **Duplicate Management**:
  - Mark duplicate rows across multiple columns
  - Remove duplicate entries with flexible column selection

#### **Missing Data Handling**
- Fill missing values (NaN, empty strings) with specified values
- Smart handling during merge operations
- Configurable null value replacement

### 🎯 **Advanced Workflow Features**

#### **Interactive Preview System**
- **Operation Preview**: See changes before applying operations
- **Output Preview**: View complete file state with all operations applied
- **Side-by-side Comparison**: Original vs. modified data visualization
- **Color-coded Changes**: 
  - Red highlighting for invalid/failed validations
  - Clear visual feedback for all changes

#### **Smart Operation Management**
- **Undo/Redo System**: Complete operation history with unlimited undo/redo
- **Delayed Processing**: Memory-efficient handling of large files
- **Progress Tracking**: Visual progress bars for long-running operations
- **Batch Operations**: Queue multiple operations before final processing to prevent overloading in memory

#### **User Experience**
- **Multi-language Interface**: English, Turkish, and Russian support
- **Persistent Settings**: Remember language preferences and directories (doesn't apply for released executable)
- **Status Logging**: Comprehensive activity tracking and feedback

### 🛠 **Technical Capabilities**

#### **Performance & Memory**
- **Optimized Processing**: Efficient handling of large datasets
- **Preview Mode**: Work with data samples for fast operations
- **Memory Management**: Smart memory usage for resource-constrained environments
- **Cancellable Operations**: Stop long-running processes when needed

#### **Data Integrity**
- **Validation Engine**: Comprehensive input validation
- **Error Recovery**: Robust error handling with detailed messages
- **Data Type Preservation**: Maintain original data types where appropriate
- **Backup System**: Automatic preservation of original data

## 🌍 Multi-Language Support

Excel Table Tools supports three languages:
- **English** (Default)
- **Turkish** (Türkçe)
- **Russian** (Русский)

Language preference is automatically saved and restored between sessions.


## 🚀 Quick Start

### Executables 
You can find executables in **[Releases](https://github.com/AhmetCanArslan/ExcelTableTools/releases)** page. 

### Installation

1. **Clone Repository**:
   ```bash
   git clone github.com/AhmetCanArslan/ExcelTableTools
   ```
2. Change Directory:
   ```bash
   cd ExcelTableTools
   ```
3. **Install Python Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
4. **Run the Application**:
   ```bash
   python excel_table_tools.py
   ```
   
### Generate Standalone Executable

#### Linux
```bash
chmod +x GenerateExecutable/build_linux.sh
./GenerateExecutable/build_linux.sh
```

#### macOS
```bash
chmod +x GenerateExecutable/build_macos.sh
./GenerateExecutable/build_macos.sh
```

#### Windows
```batch
GenerateExecutable\build_windows.bat
```

Executables will be created in the respective `GenerateExecutable/[platform]/` directories.


### Basic Workflow
1. **Load Data**: Click "Browse..." to select your Excel or CSV file
2. **Choose Preview Position**: Select head, middle, or tail for large files
3. **Select Column**: Choose the target column from the dropdown
4. **Choose Operation**: Select from 25+ available operations
5. **Preview Changes**: Click "Operation Preview" to see effects before applying
6. **Apply Operation**: Click "Apply Operation" to execute
7. **Review Results**: Use "Output File Preview" to see complete file state
8. **Save Results**: Choose output format and save your processed data

## 📋 Complete Operation Reference

### Data Privacy & Masking
| Operation | Description | Example |
|-----------|-------------|---------|
| Mask Column | Keep first/last 2 chars | `12345678` → `12****78` |
| Mask Email | Protect email addresses | `user@domain.com` → `us***@domain.com` |
| Mask Words | Mask individual words | `John Smith` → `Jo** Sm***` |

### Text Processing
| Operation | Description | Use Case |
|-----------|-------------|----------|
| Trim Spaces | Remove whitespace | Clean imported data |
| UPPERCASE | Convert to uppercase | Standardize codes |
| lowercase | Convert to lowercase | Normalize names |
| Title Case | Capitalize words | Format names |
| Find & Replace | Text substitution | Fix common errors |
| Remove Specific Chars | Custom character removal | Clean special chars |
| Remove Non-numeric | Keep only numbers | Extract numeric data |
| Remove Non-alphabetic | Keep only letters | Extract text data |

### Column Operations
| Operation | Description | Parameters |
|-----------|-------------|------------|
| Split by Delimiter | Split using separator | Custom delimiter |
| Split Surname | Extract last word | Automatic detection |
| Concatenate Columns | Join multiple columns | Custom separator |
| Merge Columns | Advanced column joining | Missing value handling |
| Rename Column | Change column name | New name validation |
| Extract with Regex | Pattern-based extraction | Regex pattern, new column |

### Numeric Operations
| Operation | Description | Configuration |
|-----------|-------------|---------------|
| Round Numbers | Round to decimals | Decimal places (0-10) |
| Calculate by Constant | Math with constant | Operation (+,-,*,/), value |
| Create Calculated Column | Math between columns | Two columns, operation |

### Data Validation
| Operation | Validates | Output |
|-----------|-----------|--------|
| Validate Email | Email format | Visual highlighting |
| Validate Phone | Phone numbers | Error identification |
| Validate Date | Date formats | Format checking |
| Validate Numeric | Number values | Type validation |
| Validate Alphanumeric | Text format | Pattern matching |
| Validate URL | Web addresses | URL format check |

### Data Quality
| Operation | Purpose | Options |
|-----------|---------|---------|
| Fill Missing Values | Replace NaN/empty | Custom fill value |
| Mark Duplicates | Identify duplicates | Multi-column selection |
| Remove Duplicates | Delete duplicate rows | Column-based removal |
| Distinct Group Numbers | Categorize unique values | Automatic numbering |

## 🏗 Project Architecture

```
ExcelTableTools/
├── excel_table_tools.py        # Main launcher
├── README.md                   # This documentation
├── CHANGELOG.md               # Version history
├── requirements.txt           # Dependencies
├── src/                       # Core application
│   ├── main.py               # Main GUI application
│   ├── translations.py       # Multi-language support
│   └── operations/           # Operation modules
│       ├── delayed_operations.py  # Batch processing
│       ├── preview_utils.py       # Preview functionality
│       ├── masking.py            # Data masking
│       ├── validation.py        # Data validation
│       ├── numeric_operations.py # Math operations
│       └── [other modules]      # Specific operations
├── resources/                 # Configuration
│   ├── operations_config.json # Operation definitions
│   └── [settings files]      # User preferences
└── GenerateExecutable/       # Build system
    ├── build_*.sh|bat        # Platform builders
    └── [platform]/          # Output directories
```



## 📄 License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.

### What this means:
- ✅ **Freedom to use**: You can use this software for any purpose
- ✅ **Freedom to study**: You can examine how the program works and modify it
- ✅ **Freedom to share**: You can redistribute copies of the software
- ✅ **Freedom to improve**: You can distribute modified versions to help the community

### Key requirements:
- If you distribute this software or any derivative work, you must make the source code available
- Any modifications must also be licensed under GPL v3
- You must include copyright notices and license information

For more information about the GNU GPL v3, visit: https://www.gnu.org/licenses/gpl-3.0.html

## 🤝 Contributing

Contributions are welcome! Please read our contributing guidelines and submit pull requests for any improvements.
