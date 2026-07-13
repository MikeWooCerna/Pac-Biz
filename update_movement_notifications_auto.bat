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
echo ========================================
echo Reconciling movement notification ledger and dashboard snapshot
echo ========================================
py -3 "%MASTERLIST_DIR%\movement_reconcile.py"
if errorlevel 1 (
    echo [ERROR] Movement reconciliation failed.
    exit /b 1
)

echo.
echo ========================================
echo Publishing reconciled movement snapshot if changed
echo ========================================
git pull --rebase --autostash
if errorlevel 1 exit /b 1

git add masterlist_dashboard.html update_movement_notifications_auto.bat movement_reconcile.py
git diff --cached --quiet
if errorlevel 1 (
    git commit -m "Reconcile movement dashboard snapshot"
    if errorlevel 1 exit /b 1
    git pull --rebase --autostash
    if errorlevel 1 exit /b 1
    git push
    if errorlevel 1 exit /b 1
) else (
    echo No movement dashboard changes to publish.
)

echo.
echo Done.
exit /b 0
