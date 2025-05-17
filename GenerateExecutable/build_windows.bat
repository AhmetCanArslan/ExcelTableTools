@echo off
rem Windows batch script to build Excel Table Tools executable

rem Get the directory where the script is located
set "SCRIPT_DIR=%~dp0"
set "PROJECT_ROOT=%SCRIPT_DIR%.."
cd "%PROJECT_ROOT%"

rem Create target directory for the executable
set "TARGET_DIR=%SCRIPT_DIR%windows"
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

rem --- Ensure Virtual Environment Exists ---
set "VENV_PATH=%PROJECT_ROOT%\venv\Scripts\activate.bat"
if not exist "%PROJECT_ROOT%\venv\" (
    echo Virtual environment not found. Creating one in '%PROJECT_ROOT%\venv'...
    python -m venv "%PROJECT_ROOT%\venv"
)

if exist "%VENV_PATH%" (
    echo Activating virtual environment...
    call "%VENV_PATH%"
) else (
    echo Error: Virtual environment activation script not found at %VENV_PATH%
    echo Please ensure your virtual environment is named 'venv' or adjust VENV_PATH in this script.
    exit /b 1
)

rem --- Ensure all Python dependencies are installed ---
echo Installing/upgrading required Python packages...
pip install -r "%PROJECT_ROOT%\requirements.txt"

rem --- Ensure PyInstaller is installed ---
echo Installing/upgrading pyinstaller...
pip install -U pyinstaller

rem Run PyInstaller
echo Building ExcelTableTools...
pyinstaller --clean ^
    --add-data "%PROJECT_ROOT%\resources;resources" ^
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
    --hidden-import src.operations.validate_inputs ^
    --hidden-import src.translations ^
    --name "ExcelTableTools" ^
    --noconsole ^
    --distpath "%TARGET_DIR%" ^
    --workpath "%TARGET_DIR%\build" ^
    --specpath "%TARGET_DIR%" ^
    --onefile ^
    --noupx ^
    "%PROJECT_ROOT%\excel_table_tools.py"

rem Add some feedback
if %ERRORLEVEL% EQU 0 (
    echo Build successful! Check the '%TARGET_DIR%' folder.
    echo You can run the application with: %TARGET_DIR%\ExcelTableTools.exe

    rem Remove build artifacts not needed by end user
    if exist "%TARGET_DIR%\build" rd /s /q "%TARGET_DIR%\build"
    if exist "%TARGET_DIR%\ExcelTableTools.spec" del "%TARGET_DIR%\ExcelTableTools.spec"
    
    rem Create a simple launcher script in the root directory
    echo @echo off > "%PROJECT_ROOT%\run_excel_tools.bat"
    echo cd /d "%%~dp0" >> "%PROJECT_ROOT%\run_excel_tools.bat"
    echo .\GenerateExecutable\ExcelTableTools\ExcelTableTools.exe %%* >> "%PROJECT_ROOT%\run_excel_tools.bat"
    
    echo A launcher script has been created at: %PROJECT_ROOT%\run_excel_tools.bat
) else (
    echo Build failed.
)

rem Keep the terminal window open until Enter is pressed
echo.
pause

rem Optional: Deactivate virtual environment (though script exit will handle this)
rem deactivate