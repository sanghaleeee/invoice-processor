@echo off
REM ============================================
REM Invoice Processor - Windows Build Script
REM ============================================
echo Building Invoice Processor...

pyinstaller --onefile --windowed ^
    --name "Invoice Processor" ^
    --icon=appicon.ico ^
    --add-data "process_invoice.py;." ^
    invoice_app.py

if %errorlevel% neq 0 (
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

echo.
echo Build complete!
echo EXE: dist\Invoice Processor.exe
echo.
echo You can drag PDF/.xlsx files onto the .exe, or double-click to launch.
pause
