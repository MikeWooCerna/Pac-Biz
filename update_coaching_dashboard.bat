@echo off
setlocal

set "COACHING_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Coaching"
set "MASTERLIST_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Masterlist"

echo.
echo ========================================
echo Updating Coaching data from Asana
echo ========================================
cd /d "%COACHING_DIR%"
if errorlevel 1 goto :fail

py -3 asana_pull.py
if errorlevel 1 goto :fail

echo.
echo ========================================
echo Rebuilding dashboard
echo ========================================
cd /d "%MASTERLIST_DIR%"
if errorlevel 1 goto :fail

echo Syncing latest dashboard repo changes...
git pull --rebase --autostash
if errorlevel 1 goto :fail

py -3 dashboard.py
if errorlevel 1 goto :fail

echo.
echo ========================================
echo Publishing to GitHub
echo ========================================
git add dashboard.py masterlist_dashboard.html
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
