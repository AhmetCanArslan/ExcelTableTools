@echo off
setlocal enabledelayedexpansion

:: Get the directory where the script is located
set "SCRIPT_DIR=%~dp0"
:: Remove trailing backslash
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

:: Get parent directory (project root) - using pushd/popd method to handle paths with parentheses
pushd "%SCRIPT_DIR%\.."
set "PROJECT_ROOT=%CD%"
popd

echo Script directory: %SCRIPT_DIR%
echo Project root: %PROJECT_ROOT%

:: Change to the project root directory
cd /d "%PROJECT_ROOT%"

:: --- Activate Virtual Environment ---
:: Try to find a virtual environment in common locations
set "VENV_FOUND="

if exist ".venv\Scripts\activate.bat" (
    set "VENV_PATH=.venv\Scripts\activate.bat"
    set "VENV_FOUND=1"
) else if exist "venv\Scripts\activate.bat" (
    set "VENV_PATH=venv\Scripts\activate.bat"
    set "VENV_FOUND=1"
)

if defined VENV_FOUND (
    echo Activating virtual environment at %VENV_PATH%...
    call "%VENV_PATH%"
) else (
    echo Warning: Virtual environment not found. Using system Python.
)

:: Create target directory for the executable
set "TARGET_DIR=%SCRIPT_DIR%\windows"
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

:: Create temporary build directory
set "TEMP_BUILD_DIR=%SCRIPT_DIR%\.build_temp"
if not exist "%TEMP_BUILD_DIR%" mkdir "%TEMP_BUILD_DIR%"

:: Ensure required directories exist
echo Ensuring required directories exist...
if not exist "%PROJECT_ROOT%\src\config" mkdir "%PROJECT_ROOT%\src\config"
if not exist "%PROJECT_ROOT%\resources" mkdir "%PROJECT_ROOT%\resources"

:: Ensure PyInstaller is installed and install requirements
echo Installing PyInstaller and requirements...
python -m pip install -U pyinstaller
python -m pip install -r requirements.txt

:: Check if spec file exists, if not create a basic one
if not exist "%PROJECT_ROOT%\excel_table_tools.spec" (
    echo Creating PyInstaller spec file...
    python -m PyInstaller --onefile --windowed --name ExcelTableTools excel_table_tools.py --specpath "%PROJECT_ROOT%"
    echo Spec file created. You may want to customize it for better results.
)

:: Create a more direct and reliable build command
echo Building ExcelTableTools...
python -m PyInstaller --clean --workpath "%TEMP_BUILD_DIR%" --distpath "%TARGET_DIR%" "%PROJECT_ROOT%\excel_table_tools.spec"

:: Check build result
if %ERRORLEVEL% equ 0 (
    echo Build successful! Check the '%TARGET_DIR%' folder.
    if exist "%TARGET_DIR%\ExcelTableTools.exe" (
        echo You can run the application with: %TARGET_DIR%\ExcelTableTools.exe
    ) else (
        echo Warning: ExcelTableTools.exe not found in output directory
        dir "%TARGET_DIR%"
    )
    
    :: Remove build artifacts not needed by end user
    if exist "%TEMP_BUILD_DIR%" (
        echo Cleaning up temporary build directory...
        rmdir /s /q "%TEMP_BUILD_DIR%"
    )
    
    :: Create a simple launcher script in the root directory (only in local builds)
    if not defined GITHUB_ACTIONS (
        echo Creating launcher script...
        echo @echo off > "%PROJECT_ROOT%\run_excel_tools.bat"
        echo cd /d "%%~dp0" >> "%PROJECT_ROOT%\run_excel_tools.bat"
        echo GenerateExecutable\windows\ExcelTableTools.exe %%* >> "%PROJECT_ROOT%\run_excel_tools.bat"
        echo A launcher script has been created at: %PROJECT_ROOT%\run_excel_tools.bat
    )
) else (
    echo Build failed with error code %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)

:: Only pause if not running in GitHub Actions
if not defined GITHUB_ACTIONS (
    echo.
    pause
)