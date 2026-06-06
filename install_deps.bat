@echo off
REM ============================================
REM Invoice Processor - Windows Dependencies
REM ============================================
echo Installing Invoice Processor dependencies...
echo.

pip install PyMuPDF openpyxl pyinstaller

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] pip install failed. Make sure Python is installed.
    echo Download from: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo.
echo Dependencies installed successfully!
echo.
echo To build .exe:
echo   pyinstaller --onefile --windowed --icon=appicon.ico --name "Invoice Processor" invoice_app.py
echo.
pause
