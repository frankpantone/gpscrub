@echo off
REM Automated Email PDF Processor Launcher
REM Connects to Outlook, downloads PDFs, creates email drafts

echo ================================================================================
echo HERTZ EMAIL PDF PROCESSOR - AUTOMATED
echo ================================================================================
echo.
echo This script will:
echo   1. Connect to your Outlook (hertzlogistics@hertz.com)
echo   2. Search CarMax GP Folder for emails with KUNES and EASTON
echo   3. Download PDF attachments from TODAY
echo   4. Create email drafts for each PDF
echo.
echo ================================================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo [OK] Python found
echo.

REM Run the email processor
echo Starting email processor...
echo.
python email_pdf_processor.py

echo.
echo ================================================================================
echo.
pause




