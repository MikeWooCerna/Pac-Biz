@echo off
setlocal

set "MASTERLIST_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Masterlist"

cd /d "%MASTERLIST_DIR%"

echo.
echo ========================================
echo Refreshing Masterlist movement cache
echo ========================================
py -3 "%MASTERLIST_DIR%\masterlist_fetch.py"
if errorlevel 1 (
    echo [ERROR] Masterlist movement cache refresh failed.
    exit /b 1
)

echo.
echo ========================================
echo Sending processed movement notifications
echo ========================================
py -3 "%MASTERLIST_DIR%\check_movement_notifications.py"
if errorlevel 1 (
    echo [ERROR] Movement notifications failed.
    exit /b 1
)

echo.
echo Done.
exit /b 0
