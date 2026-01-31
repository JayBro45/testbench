@echo off
TITLE AVR Test Bench Software
echo ========================================================
echo      STATCON ELECTRONICS - AVR TEST BENCH LAUNCHER
echo ========================================================
echo.

:: 1. Navigate to the script's directory to avoid path errors
cd /d "%~dp0"

:: 2. Check if Python is accessible
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [CRITICAL ERROR] Python is not installed or not in your PATH.
    echo Please install Python 3.10+ and check "Add to PATH" during installation.
    echo.
    pause
    exit /b
)

:: 3. Launch the Application
echo Starting Application...
echo Logging to: %~dp0logs\testbench.log
echo.
python main.py

:: 4. Pause only if the app crashes (error level is not 0)
if %errorlevel% neq 0 (
    echo.
    echo ========================================================
    echo [ERROR] The application has crashed!
    echo Please check the error message above or the log file.
    echo ========================================================
    pause
)