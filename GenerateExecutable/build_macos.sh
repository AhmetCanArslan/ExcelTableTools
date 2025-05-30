#!/bin/bash

# macOS script to build Excel Table Tools executable

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
TARGET_DIR="$SCRIPT_DIR/macos"
mkdir -p "$TARGET_DIR"

# Create temporary build directory
TEMP_BUILD_DIR="$SCRIPT_DIR/.build_temp"
mkdir -p "$TEMP_BUILD_DIR"

# Ensure PyInstaller is installed and install macOS-specific requirements
pip install -U pyinstaller
pip install -r requirements-macos.txt

# Create a more direct and reliable build command
echo "Building ExcelTableTools..."
pyinstaller --clean \
    --workpath "$TEMP_BUILD_DIR" \
    --distpath "$TARGET_DIR" \
    "$PROJECT_ROOT/excel_table_tools.spec"

# Optional: Add some feedback
if [ $? -eq 0 ]; then
    echo "Build successful! Check the '$TARGET_DIR' folder."
    echo "You can run the application with: $TARGET_DIR/ExcelTableTools"
    
    # Remove build artifacts not needed by end user
    rm -rf "$TEMP_BUILD_DIR"
    
    # Make the executable file executable
    chmod +x "$TARGET_DIR/ExcelTableTools"
    
    # Create a simple launcher script in the root directory
    echo '#!/bin/bash
cd "$(dirname "$0")"
./GenerateExecutable/macos/ExcelTableTools "$@"' > "$PROJECT_ROOT/run_excel_tools.command"
    chmod +x "$PROJECT_ROOT/run_excel_tools.command"
    
    echo "A launcher script has been created at: $PROJECT_ROOT/run_excel_tools.command"
else
    echo "Build failed."
fi

# Keep the terminal window open until Enter is pressed
echo "" # Add a blank line for readability
read -p "Press Enter to exit..."
