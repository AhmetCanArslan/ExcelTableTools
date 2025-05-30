@echo off
setlocal enabledelayedexpansion

:: Get the directory where the script is located
set "SCRIPT_DIR=%~dp0"
:: Remove trailing backslash
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
:: Get parent directory (project root)
for %%I in ("%SCRIPT_DIR%") do set "PROJECT_ROOT=%%~dpI"
:: Remove trailing backslash
set "PROJECT_ROOT=%PROJECT_ROOT:~0,-1%"

:: Change to the project root directory
cd /d "%PROJECT_ROOT%"

:: --- Activate Virtual Environment ---
:: Try to find a virtual environment in common locations
set "VENV_PATHS=.venv\Scripts\activate.bat venv\Scripts\activate.bat"
set "VENV_FOUND="

for %%p in (%VENV_PATHS%) do (
    if exist "%%p" (
        set "VENV_PATH=%%p"
        set "VENV_FOUND=1"
        goto :found_venv
    )
)
:found_venv

if defined VENV_FOUND (
    echo Activating virtual environment at %VENV_PATH%...
    call "%VENV_PATH%"
) else (
    echo Warning: Virtual environment not found. Using system Python.
    :: Continue anyway, assuming system Python has the needed packages
)

:: Create target directory for the executable
set "TARGET_DIR=%SCRIPT_DIR%\windows"
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

:: Create temporary build directory
set "TEMP_BUILD_DIR=%SCRIPT_DIR%\.build_temp"
if not exist "%TEMP_BUILD_DIR%" mkdir "%TEMP_BUILD_DIR%"

:: Ensure PyInstaller is installed and install requirements
python -m pip install -U pyinstaller
python -m pip install -r requirements.txt

:: Check if spec file exists, if not create a basic one
if not exist "%PROJECT_ROOT%\excel_table_tools.spec" (
    echo Creating PyInstaller spec file...
    pyinstaller --onefile --windowed --name ExcelTableTools excel_table_tools.py --specpath "%PROJECT_ROOT%"
    echo Spec file created. You may want to customize it for better results.
)

:: Create a more direct and reliable build command
echo Building ExcelTableTools...
pyinstaller --clean ^
    --workpath "%TEMP_BUILD_DIR%" ^
    --distpath "%TARGET_DIR%" ^
    "%PROJECT_ROOT%\excel_table_tools.spec"

:: Optional: Add some feedback
if %ERRORLEVEL% equ 0 (
    echo Build successful! Check the '%TARGET_DIR%' folder.
    echo You can run the application with: %TARGET_DIR%\ExcelTableTools.exe
    
    :: Remove build artifacts not needed by end user
    if exist "%TEMP_BUILD_DIR%" rmdir /s /q "%TEMP_BUILD_DIR%"
    
    :: Create a simple launcher script in the root directory (only in local builds)
    if not defined GITHUB_ACTIONS (
        echo @echo off > "%PROJECT_ROOT%\run_excel_tools.bat"
        echo cd "%%~dp0" >> "%PROJECT_ROOT%\run_excel_tools.bat"
        echo GenerateExecutable\windows\ExcelTableTools.exe %%* >> "%PROJECT_ROOT%\run_excel_tools.bat"
        echo A launcher script has been created at: %PROJECT_ROOT%\run_excel_tools.bat
    )
) else (
    echo Build failed.
    exit /b 1
)

:: Only pause if not running in GitHub Actions
if not defined GITHUB_ACTIONS (
    echo.
    pause
)