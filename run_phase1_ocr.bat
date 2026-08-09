@echo off
echo ============================================================
echo  PHASE 1: MADRASA OCR DATA EXTRACTION
echo ============================================================

echo.
echo [1/2] Running OCR pipeline (images/ folder)...
python scripts\ocr_pipeline.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo WARNING: OCR pipeline had errors or was manually interrupted.
    set /p choice="Do you want to continue and build the Excel sheet? (Y/N): "
    if /i not "%choice%"=="Y" exit /b %ERRORLEVEL%
)

echo.
echo [2/2] Injecting new rows into forms_v5.xlsm (CLOSE IT IN EXCEL FIRST!)...
python scripts\inject_csv_to_excel.py

echo.
echo ============================================================
echo  PHASE 1 COMPLETE! You can now run Phase 2 (Trains).
echo ============================================================
pause
