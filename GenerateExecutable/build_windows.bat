@echo off
rem Windows batch script to build Excel Table Tools executable

rem Get the directory where the script is located
set "SCRIPT_DIR=%~dp0"
set "PROJECT_ROOT=%SCRIPT_DIR%.."
cd "%PROJECT_ROOT%"

rem --- Activate Virtual Environment ---
set "VENV_PATH=%PROJECT_ROOT%\venv\Scripts\activate.bat"

if exist "%VENV_PATH%" (
    echo Activating virtual environment...
    call "%VENV_PATH%"
) else (
    echo Error: Virtual environment activation script not found at %VENV_PATH%
    echo Please ensure your virtual environment is named 'venv' or adjust VENV_PATH in this script.
    exit /b 1
)

rem Run PyInstaller
echo Building ExcelTableTools...
pyinstaller --clean ^
    --add-data "resources;resources" ^
    --hidden-import pandas ^
    --hidden-import openpyxl ^
    --hidden-import tabulate ^
    --hidden-import src ^
    --hidden-import src.operations ^
    --hidden-import src.operations.masking ^
    --hidden-import src.operations.trimming ^
    --hidden-import src.operations.splitting ^
    --hidden-import src.operations.case_change ^
    --hidden-import src.operations.find_replace ^
    --hidden-import src.operations.remove_chars ^
    --hidden-import src.operations.concatenate ^
    --hidden-import src.operations.extract_pattern ^
    --hidden-import src.operations.fill_missing ^
    --hidden-import src.operations.duplicates ^
    --hidden-import src.operations.merge_columns ^
    --hidden-import src.operations.rename_column ^
    --hidden-import src.operations.preview_utils ^
    --hidden-import src.operations.numeric_operations ^
    --hidden-import src.translations ^
    --name "ExcelTableTools" ^
    --console ^
    excel_table_tools.py

rem Add some feedback
if %ERRORLEVEL% EQU 0 (
    echo Build successful! Check the 'dist' folder.
    echo You can run the application with: dist\ExcelTableTools\ExcelTableTools.exe
) else (
    echo Build failed.
)

rem Keep the terminal window open until Enter is pressed
echo.
pause

rem Optional: Deactivate virtual environment (though script exit will handle this)
rem deactivate