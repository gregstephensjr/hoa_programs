@echo off
REM ============================================================================
REM PDF Batch Processor - Easy Run Script
REM Double-click this file to process PDFs in the current folder
REM ============================================================================

echo.
echo ================================================================================
echo                      PDF BATCH PROCESSOR
echo ================================================================================
echo.
echo This script will:
echo   1. Count three-letter codes from all PDFs
echo   2. Create "add to service charges.xlsx" spreadsheet
echo   3. Create "Print these" folder with combined and multi-page PDFs
echo.
echo ================================================================================
echo.

REM Get the directory where this batch file is located
set "SCRIPT_DIR=%~dp0"

REM Remove trailing backslash if present
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

echo Processing PDFs in: %SCRIPT_DIR%
echo.

REM Run the Python script
python "%SCRIPT_DIR%\batch_process.py" "%SCRIPT_DIR%"

REM Check if the script ran successfully
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ================================================================================
    echo SUCCESS! Processing complete.
    echo ================================================================================
    echo.
    echo Check the following:
    echo   - "add to service charges.xlsx" in this folder
    echo   - "Print these" folder with files ready to print
    echo.
) else (
    echo.
    echo ================================================================================
    echo ERROR! Something went wrong.
    echo ================================================================================
    echo.
    echo Make sure:
    echo   1. Python is installed on this computer
    echo   2. Required libraries are installed (pdfplumber, pypdf, openpyxl)
    echo   3. batch_process.py is in the same folder as this batch file
    echo.
)

pause
