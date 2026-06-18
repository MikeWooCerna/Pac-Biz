@echo off
setlocal

set "COACHING_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Coaching"
set "M7_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\M7"
set "PARENTIS_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\Parentis Health"
set "BRITELIFT_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\Britelift"
set "RIDEX_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\RideX"
set "HAMILTON_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\Hamilton"
set "SKYLINE_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\Skyline"
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
echo Updating M7 QA data from Google Sheets
echo ========================================
cd /d "%M7_DIR%"
if errorlevel 1 goto :fail

py -3 m7_pull.py
if errorlevel 1 goto :fail

echo.
echo ========================================
echo Updating Parentis Health QA data from Google Sheets
echo ========================================
cd /d "%PARENTIS_DIR%"
if errorlevel 1 goto :fail

py -3 parentis_pull.py
if errorlevel 1 goto :fail

echo.
echo ========================================
echo Updating Britelift QA data from Google Sheets
echo ========================================
cd /d "%BRITELIFT_DIR%"
if errorlevel 1 goto :fail

py -3 britelift_pull.py
if errorlevel 1 goto :fail

echo.
echo ========================================
echo Updating RideX QA data from Google Sheets
echo ========================================
cd /d "%RIDEX_DIR%"
if errorlevel 1 goto :fail

py -3 Ridex_pull.py
if errorlevel 1 goto :fail

echo.
echo ========================================
echo Updating Hamilton QA data from Google Sheets
echo ========================================
cd /d "%HAMILTON_DIR%"
if errorlevel 1 goto :fail

py -3 Hamilton_pull.py
if errorlevel 1 goto :fail

echo.
echo ========================================
echo Updating Skyline QA data from Google Sheets
echo ========================================
cd /d "%SKYLINE_DIR%"
if errorlevel 1 goto :fail

py -3 Skyline_pull.py
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
