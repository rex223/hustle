@echo off
echo ============================================================
echo  PHASE 3: LOCAL TRANSPORT LOCATOR
echo ============================================================

echo.
echo [1/2] Finding Nearest Bazaars and Autostands...
python scripts\find_local.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo ERROR: Failed to find local transport data.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [2/2] Injecting Local Transport Data into Excel...
python scripts\inject_local.py

echo.
echo ============================================================
echo  PHASE 3 COMPLETE! All data is fully populated!
echo ============================================================
pause
