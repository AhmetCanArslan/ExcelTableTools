#!/bin/bash

# Get the directory where the script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"

# Change to the project root directory
cd "$PROJECT_ROOT"

# --- Activate Virtual Environment ---
# Adjust the path to your virtual environment's activate script if different
VENV_PATH="$PROJECT_ROOT/.venv/bin/activate" # Common path

if [ -f "$VENV_PATH" ]; then
    echo "Activating virtual environment..."
    source "$VENV_PATH"
else
    echo "Error: Virtual environment activation script not found at $VENV_PATH"
    echo "Please ensure your virtual environment is named '.venv' or adjust VENV_PATH in this script."
    exit 1
fi

# Run PyInstaller
# Ensure pyinstaller is in your PATH or provide the full path to it.
echo "Building ExcelTableToolsApp..."
pyinstaller --onefile --windowed --name ExcelTableToolsApp "$PROJECT_ROOT/src/main.py"

# Optional: Add some feedback
if [ $? -eq 0 ]; then
    echo "Build successful! Check the 'dist' folder."
else
    echo "Build failed."
fi

# Keep the terminal window open until Enter is pressed
echo "" # Add a blank line for readability
read -p "Press Enter to exit..."
