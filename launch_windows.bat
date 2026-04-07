@echo off
:: XLSForm Quality Reviewer — Windows Launcher
:: Double-click this file to start the app.

title XLSForm Quality Reviewer

echo ================================================
echo   XLSForm Quality Reviewer
echo ================================================
echo.

:: Change to the directory where this script lives
cd /d "%~dp0"

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not on PATH.
    echo Please follow the instructions in SETUP.md ^(Step 1^) and try again.
    pause
    exit /b 1
)

:: Create virtual environment if it does not exist
if not exist "venv\" (
    echo Setting up the virtual environment for the first time...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo ERROR: Failed to create virtual environment.
        pause
        exit /b 1
    )
)

:: Activate the virtual environment
call venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo ERROR: Failed to activate virtual environment.
    echo If you see a script execution error, please follow the PowerShell note
    echo in SETUP.md ^(Windows Step 4^) and try again.
    pause
    exit /b 1
)

:: Install / update dependencies
echo Checking dependencies...
pip install -q -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies.
    echo Please check your internet connection and try again.
    pause
    exit /b 1
)

echo.
echo Starting the application...
echo The app will open in your browser at http://localhost:8501
echo To stop the app, close this window or press Ctrl + C.
echo.

streamlit run app.py

pause
