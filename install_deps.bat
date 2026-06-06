@echo off
REM ============================================
REM Invoice Processor - Windows Dependency Installer
REM ============================================
title Invoice Processor Setup
echo ========================================
echo  Invoice Processor - Windows Setup
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed!
    echo.
    echo Download Python 3.9+ from:
    echo   https://www.python.org/downloads/
    echo.
    echo IMPORTANT: During installation, check "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

echo [OK] Python found:
python --version

REM Check if pip is available
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo [WARNING] pip not found. Attempting to install...
    python -m ensurepip --upgrade
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to install pip.
        echo Try: python -m ensurepip
        pause
        exit /b 1
    )
)

echo [OK] pip found:
pip --version
echo.

REM Install dependencies
echo Installing dependencies...
pip install PyMuPDF openpyxl
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

echo.
echo ========================================
echo  Setup complete!
echo ========================================
echo.
echo To run the app:
echo   python invoice_app.py
echo.
echo To build a standalone .exe:
echo   pip install pyinstaller
echo   pyinstaller --onefile --windowed --name "Invoice Processor" invoice_app.py
echo   (dist\Invoice Processor.exe)
echo.
pause
