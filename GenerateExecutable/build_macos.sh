#!/bin/bash

# macOS script to build Excel Table Tools executable

# Get the directory where the script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"

# Change to the project root directory
cd "$PROJECT_ROOT"

# --- Ensure Virtual Environment Exists ---
VENV_PATH=""
if [ -d "$PROJECT_ROOT/venv" ]; then
    VENV_PATH="$PROJECT_ROOT/venv/bin/activate"
elif [ -d "$PROJECT_ROOT/.venv" ]; then # Common alternative
    VENV_PATH="$PROJECT_ROOT/.venv/bin/activate"
fi

if [ -f "$VENV_PATH" ]; then
    echo "Activating virtual environment at $VENV_PATH..."
    source "$VENV_PATH"
else
    echo "Virtual environment not found. Creating one in '$PROJECT_ROOT/venv'..."
    python3 -m venv "$PROJECT_ROOT/venv" # Use python3 explicitly on macOS
    if [ -f "$PROJECT_ROOT/venv/bin/activate" ]; then
        VENV_PATH="$PROJECT_ROOT/venv/bin/activate"
        echo "Activating newly created virtual environment..."
        source "$VENV_PATH"
    else
        echo "Error: Failed to create or find virtual environment."
        echo "Please ensure Python 3 and venv are installed, or create 'venv' manually."
        exit 1
    fi
fi

# --- Ensure all Python dependencies are installed ---
echo "Installing/upgrading required Python packages..."
pip install -r "$PROJECT_ROOT/requirements.txt"

# --- Ensure PyInstaller is installed ---
echo "Installing/upgrading pyinstaller..."
pip install -U pyinstaller

# Create target directory for the executable
TARGET_DIR="$SCRIPT_DIR/macos"
mkdir -p "$TARGET_DIR"

# Run PyInstaller
echo "Building ExcelTableTools for macOS..."
pyinstaller --clean \
    --add-data "$PROJECT_ROOT/resources:resources" \
    --hidden-import pandas \
    --hidden-import openpyxl \
    --hidden-import tabulate \
    --hidden-import src \
    --hidden-import src.operations \
    --hidden-import src.operations.masking \
    --hidden-import src.operations.trimming \
    --hidden-import src.operations.splitting \
    --hidden-import src.operations.case_change \
    --hidden-import src.operations.find_replace \
    --hidden-import src.operations.remove_chars \
    --hidden-import src.operations.concatenate \
    --hidden-import src.operations.extract_pattern \
    --hidden-import src.operations.fill_missing \
    --hidden-import src.operations.duplicates \
    --hidden-import src.operations.merge_columns \
    --hidden-import src.operations.rename_column \
    --hidden-import src.operations.preview_utils \
    --hidden-import src.operations.numeric_operations \
    --hidden-import src.operations.validate_inputs \
    --hidden-import src.translations \
    --hidden-import pyobjc-core \
    --hidden-import pyobjc-framework-Cocoa \
    --name "ExcelTableTools" \
    --noconsole \
    --distpath "$TARGET_DIR" \
    --workpath "$TARGET_DIR/build" \
    --specpath "$TARGET_DIR" \
    --onefile \
    --noupx \
    "$PROJECT_ROOT/excel_table_tools.py"

# Add some feedback
if [ $? -eq 0 ]; then
    echo "Build successful! Check the '$TARGET_DIR' folder."
    echo "The executable is at: $TARGET_DIR/ExcelTableTools"
    
    # Remove build artifacts not needed by end user
    rm -rf "$TARGET_DIR/build"
    rm -f "$TARGET_DIR/ExcelTableTools.spec"
    
    # Make the executable file executable
    chmod +x "$TARGET_DIR/ExcelTableTools"
    
    # Create a simple launcher script in the root directory
    echo '#!/bin/bash
# Launcher for ExcelTableTools on macOS
# Get the directory of the script itself
SCRIPT_LAUNCHER_DIR="$( cd "$( dirname "\${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
# Execute the application
"\$SCRIPT_LAUNCHER_DIR/GenerateExecutable/macos/ExcelTableTools" "$@"' > "$PROJECT_ROOT/run_excel_tools_macos.sh"
    chmod +x "$PROJECT_ROOT/run_excel_tools_macos.sh"
    
    echo "A launcher script has been created at: $PROJECT_ROOT/run_excel_tools_macos.sh"
else
    echo "Build failed."
fi

# Keep the terminal window open until Enter is pressed (optional for macOS)
echo "" 
read -p "Press Enter to exit..."

# Optional: Deactivate virtual environment
if type deactivate &>/dev/null; then
    deactivate
fi
