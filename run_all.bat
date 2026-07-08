@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo  MADRASA DB AUTOMATION RUNNER
echo ============================================================

:: ── Find Python ──────────────────────────────────────────────
set PYTHON=python
where python >nul 2>&1 || (
    echo ERROR: Python not found on PATH. Please install Python 3.10+.
    pause & exit /b 1
)

:: ── Install dependencies if needed ───────────────────────────
echo.
echo [1/4] Checking dependencies...
%PYTHON% -m pip install --upgrade pandas requests python-dotenv pywin32 openpyxl
echo      Done.

:: ── Step 1: Seed / regenerate base CSV ───────────────────────
echo.
echo [2/4] Updating base data CSV...
%PYTHON% scripts\populate_data.py || (
    echo ERROR in populate_data.py
    pause
    exit /b 1
)

:: ── Step 2: OCR pipeline ─────────────────────────────────────
echo.
echo [3/4] Running OCR pipeline (images/ folder)...
%PYTHON% scripts\ocr_pipeline.py || echo WARNING: OCR pipeline had errors — continuing.

:: ── Step 3: Transport finder ─────────────────────────────────
echo.
echo [4/4] Finding nearest transport for new rows...
%PYTHON% scripts\nearest_finder.py || echo WARNING: Transport finder had errors — continuing.

:: ── Step 4: Rebuild Excel ────────────────────────────────────
echo.
echo [5/5] Building forms_v4.xlsm...
%PYTHON% scripts\build_excel.py || (
    echo ERROR in build_excel.py
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  DONE! Open sheets\forms_v3.xlsm (password: adinul)
echo ============================================================
pause
