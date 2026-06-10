@echo off
echo ===============================================
echo IGXL CHM Documentation Ingestion
echo ===============================================
echo.

REM Activate virtual environment if it exists
if exist "..\.venv\Scripts\activate.bat" (
    echo Activating virtual environment...
    call ..\.venv\Scripts\activate.bat
) else (
    echo Virtual environment not found, using system Python
)

echo.
echo Starting IGXL CHM batch ingestion...
echo This will process 109 CHM files from:
echo C:\Users\cpatle2\Desktop\DOC from IFX\UltraFlex Learning\IGXL_Help\IGXL_Help\IGXL_Help
echo.
echo Press Ctrl+C to cancel, or any other key to continue...
pause > nul

python batch_ingest_igxl_chm.py

echo.
echo ===============================================
echo Ingestion complete!
echo Check igxl_ingestion.log for details
echo ===============================================
pause