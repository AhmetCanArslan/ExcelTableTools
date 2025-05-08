#!/bin/bash

# Get the directory where the script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"

# Change to the script's directory to ensure main.py is found
cd "$SCRIPT_DIR"

# --- Activate Virtual Environment ---
# Adjust the path to your virtual environment's activate script if different
VENV_PATH="$SCRIPT_DIR/.venv/bin/activate" # Common path

if [ -f "$VENV_PATH" ]; then
    echo "Activating virtual environment..."
    source "$VENV_PATH"
else
    echo "Error: Virtual environment activation script not found at $VENV_PATH"
    echo "Please ensure your virtual environment is named 'venv' or adjust VENV_PATH in this script."
    exit 1
fi

# Run PyInstaller
# Ensure pyinstaller is in your PATH or provide the full path to it.
# If you use a virtual environment, make sure it's activated before running this script,
# or activate it within the script.
echo "Building ExcelTableToolsApp..."
pyinstaller --onefile --windowed --name ExcelTableToolsApp main.py

# Optional: Add some feedback
if [ $? -eq 0 ]; then
    echo "Build successful! Check the 'dist' folder."
else
    echo "Build failed."
fi

# Keep the terminal window open until Enter is pressed
echo "" # Add a blank line for readability
read -p "Press Enter to exit..."

# Optional: Deactivate virtual environment (though script exit will handle this)
# if command -v deactivate &> /dev/null; then
#    echo "Deactivating virtual environment..."
#    deactivate
# fi
