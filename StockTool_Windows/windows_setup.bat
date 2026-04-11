@echo off
echo ============================================
echo   Stock Tool - Windows Setup
echo ============================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed!
    echo Please download Python from: https://www.python.org/downloads/
    echo IMPORTANT: Check "Add Python to PATH" during installation!
    pause
    exit /b 1
)

echo [OK] Python found:
python --version
echo.

REM Create virtual environment
echo Creating virtual environment...
python -m venv venv
echo [OK] Virtual environment created
echo.

REM Activate and install packages
echo Installing packages...
call venv\Scripts\activate.bat
pip install yfinance openpyxl
echo.
echo [OK] Packages installed
echo.

REM Test
echo Testing stock_fetcher.py with AAPL...
python stock_fetcher.py AAPL
echo.

echo ============================================
echo   Setup complete!
echo ============================================
echo.
echo Next steps:
echo   1. Open output.xlsm in Excel
echo   2. Press Alt+F11 to open VBA editor
echo   3. Read vba_windows.bas for VBA code
echo   4. Change PYTHON_PATH and SCRIPT_PATH in VBA:
echo      PYTHON_PATH = "%cd%\venv\Scripts\python.exe"
echo      SCRIPT_PATH = "%cd%\stock_fetcher.py"
echo   5. Type a stock symbol in A2 (e.g. AAPL)
echo.
pause
