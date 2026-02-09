@echo off
TITLE Build IPS Testing Software
cd /d "%~dp0"

echo ========================================================
echo   Building IPS Testing Software (PyInstaller)
echo ========================================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found. Install Python and add it to PATH.
    pause
    exit /b 1
)

echo Ensuring dependencies are installed...
pip install -r requirements.txt -q
pip install pyinstaller -q

echo.
echo Running PyInstaller...
pyinstaller testbench.spec

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

echo.
echo ========================================================
echo   Build complete.
echo   Output: dist\IPS_Testing_Software.exe
echo ========================================================
echo.
echo When distributing: copy config.json next to the exe (same folder).
echo Logs will be created in a "logs" folder next to the exe.
echo.
pause
