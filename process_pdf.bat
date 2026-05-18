@echo off
REM Simple launcher for Hertz PDF Email Generator
REM Double-click this file to process PDFs

echo ================================================================================
echo HERTZ PDF PROCESSOR - SIMPLE LAUNCHER
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

REM Check for required packages
python -c "import pypdf" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Installing required package: pypdf
    pip install pypdf
    echo.
)

python -c "import win32com.client" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Installing required package: pywin32
    pip install pywin32
    echo.
)

echo [OK] All requirements satisfied
echo.
echo ================================================================================
echo.

REM Count PDF files in current directory
set count=0
for %%f in (*.pdf) do set /a count+=1

if %count%==0 (
    echo [INFO] No PDF files found in current directory
    echo.
    echo Please:
    echo 1. Place your PDF files in this folder
    echo 2. Run this script again
    echo.
    pause
    exit /b 0
)

if %count%==1 (
    echo Found 1 PDF file - Processing...
    echo.
    python hertz_email_generator.py --auto
) else (
    echo Found %count% PDF files
    echo.
    echo Choose an option:
    echo   1. Process all PDFs (batch mode)
    echo   2. Process single PDF (select from list)
    echo   3. Exit
    echo.
    set /p choice="Enter your choice (1-3): "
    
    if "%choice%"=="1" (
        echo.
        echo Processing all PDFs...
        echo.
        python batch_process_pdfs.py --auto
    ) else if "%choice%"=="2" (
        echo.
        python hertz_email_generator.py
    ) else (
        echo Exiting...
        exit /b 0
    )
)

echo.
echo ================================================================================
echo DONE!
echo ================================================================================
echo.
pause




