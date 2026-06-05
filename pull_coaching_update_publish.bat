@echo off
setlocal

set "COACHING_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Coaching"
set "MASTERLIST_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Masterlist"

echo.
echo ========================================
echo Pulling Coaching data
echo ========================================
cd /d "%COACHING_DIR%"
if errorlevel 1 goto :fail

py asana_pull.py
if errorlevel 1 goto :fail

echo.
echo ========================================
echo Updating masterlist_dashboard.html
echo ========================================
cd /d "%MASTERLIST_DIR%"
if errorlevel 1 goto :fail

echo Syncing latest GitHub changes...
git pull --rebase --autostash
if errorlevel 1 goto :fail

py dashboard.py
if errorlevel 1 goto :fail

echo.
echo ========================================
echo Publishing to GitHub
echo ========================================
git add masterlist_dashboard.html
if errorlevel 1 goto :fail

git diff --cached --quiet
if not errorlevel 1 (
    echo No dashboard changes to publish.
    goto :done
)

git commit -m "Update coaching dashboard"
if errorlevel 1 goto :fail

git pull --rebase --autostash
if errorlevel 1 goto :fail

git push
if errorlevel 1 goto :fail

:done
echo.
echo Done. GitHub Pages may take about 1 minute to update.
pause
exit /b 0

:fail
echo.
echo Update failed. Check the message above.
pause
exit /b 1
