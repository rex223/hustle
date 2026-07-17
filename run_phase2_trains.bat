@echo off
echo ============================================================
echo  PHASE 2: TRAIN AND STATION LOCATOR
echo ============================================================

echo.
echo [1/2] Finding Nearest Railway Stations and Train Data...
python scripts\find_trains.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo ERROR: Failed to find train data.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [2/2] Injecting Train Data into Excel...
python scripts\inject_trains.py

echo.
echo ============================================================
echo  PHASE 2 COMPLETE! You can now run Phase 3 (Local).
echo ============================================================
pause
