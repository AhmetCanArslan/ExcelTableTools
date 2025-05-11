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
rem Ensure pyinstaller is in your PATH or provide the full path to it.
rem If you use a virtual environment, make sure it's activated before running this script,
rem or activate it within the script.
echo Building ExcelTableToolsApp...
pyinstaller --onefile --windowed --name ExcelTableToolsApp "%PROJECT_ROOT%\src\main.py"

rem Add some feedback
if %ERRORLEVEL% EQU 0 (
    echo Build successful! Check the 'dist' folder.
) else (
    echo Build failed.
)

rem Keep the terminal window open until Enter is pressed
echo.
pause

rem Optional: Deactivate virtual environment (though script exit will handle this)
rem deactivate