@echo off
title Job Monitoring Auto Fill Web App

echo ==========================================
echo   Job Monitoring Auto Fill - Setup/Run
echo ==========================================
echo.

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo Python is not installed or not added to PATH.
    echo Please install Python 3.10+ and check "Add Python to PATH".
    pause
    exit /b
)

if not exist venv (
    echo Creating virtual environment...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo Failed to create virtual environment.
        pause
        exit /b
    )
)

echo Activating virtual environment...
call venv\Scripts\activate

echo Installing dependencies...
python -m pip install --upgrade pip
pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo Failed to install dependencies.
    pause
    exit /b
)

echo.
echo Starting Flask app...
echo Open your browser and go to:
echo http://127.0.0.1:5000
echo.
python app.py

pause
