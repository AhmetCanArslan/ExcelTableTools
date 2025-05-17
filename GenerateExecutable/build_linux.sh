#!/bin/bash

# Get the directory where the script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"

# Change to the project root directory
cd "$PROJECT_ROOT"

# --- Activate Virtual Environment ---
# Try to find a virtual environment in common locations
if [ -f "$PROJECT_ROOT/.venv/bin/activate" ]; then
    VENV_PATH="$PROJECT_ROOT/.venv/bin/activate"
elif [ -f "$PROJECT_ROOT/venv/bin/activate" ]; then
    VENV_PATH="$PROJECT_ROOT/venv/bin/activate"
else
    VENV_PATH=""
fi

if [ -n "$VENV_PATH" ]; then
    echo "Activating virtual environment at $VENV_PATH..."
    source "$VENV_PATH"
else
    echo "Warning: Virtual environment not found. Using system Python."
    # Continue anyway, assuming system Python has the needed packages
fi

# Create target directory for the executable
TARGET_DIR="$SCRIPT_DIR/linux"
mkdir -p "$TARGET_DIR"

# Ensure PyInstaller is installed
pip install -U pyinstaller

# Create a more direct and reliable build command
echo "Building ExcelTableTools..."
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
    --name "ExcelTableTools" \
    --noconsole \
    --distpath "$TARGET_DIR" \
    --workpath "$TARGET_DIR/build" \
    --specpath "$TARGET_DIR" \
    --onefile \
    --noupx  \
    "$PROJECT_ROOT/excel_table_tools.py"

# Optional: Add some feedback
if [ $? -eq 0 ]; then
    echo "Build successful! Check the '$TARGET_DIR' folder."
    echo "You can run the application with: $TARGET_DIR/ExcelTableTools"
    
    # Remove build artifacts not needed by end user
    rm -rf "$TARGET_DIR/build"
    rm -f "$TARGET_DIR/ExcelTableTools.spec"
    
    # Make the executable file executable
    chmod +x "$TARGET_DIR/ExcelTableTools"
    
    # Create a simple launcher script in the root directory
    echo '#!/bin/bash
cd "$(dirname "$0")"
./GenerateExecutable/linux/ExcelTableTools "$@"' > "$PROJECT_ROOT/run_excel_tools.sh"
    chmod +x "$PROJECT_ROOT/run_excel_tools.sh"
    
    echo "A launcher script has been created at: $PROJECT_ROOT/run_excel_tools.sh"
else
    echo "Build failed."
fi

# Keep the terminal window open until Enter is pressed
echo "" # Add a blank line for readability
read -p "Press Enter to exit..."
