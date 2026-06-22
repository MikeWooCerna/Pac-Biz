@echo off
setlocal

set "COACHING_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Coaching"
set "M7_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\M7"
set "PARENTIS_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\Parentis Health"
set "BRITELIFT_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\Britelift"
set "BLC_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\Britelift Chat"
set "RIDEX_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\RideX"
set "HAMILTON_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\Hamilton"
set "SKYLINE_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\Skyline"
set "VIP_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\VIP"
set "CH_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\C&H"
set "RC_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\Reno Cab"
set "TI_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\Trans Iowa"
set "DC_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\Data Carz"
set "AC_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\Associated Cab"
set "OL_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\Ollies"
set "CT_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\Circle Taxi"
set "YCOV_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\YCOV"
set "KEL_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\Kelowna"
set "VT_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\Vermont"
set "YCDC_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\YCDC"
set "BL_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Quality\Blueline"
set "MASTERLIST_DIR=C:\Users\Mike Woo Cerna\Documents\PB\Masterlist"
set "MONITOR_DONE=0"

cd /d "%MASTERLIST_DIR%"
py -3 "%MASTERLIST_DIR%\log_step.py" init

echo.
echo ========================================
echo Updating Coaching data from Asana
echo ========================================
cd /d "%COACHING_DIR%"
if errorlevel 1 goto :fail

py -3 asana_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Coaching" "asana_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "Coaching" "asana_pull.py" 0

echo.
echo ========================================
echo Updating M7 QA data from Google Sheets
echo ========================================
cd /d "%M7_DIR%"
if errorlevel 1 goto :fail

py -3 m7_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "M7" "m7_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "M7" "m7_pull.py" 0

echo.
echo ========================================
echo Updating Parentis Health QA data from Google Sheets
echo ========================================
cd /d "%PARENTIS_DIR%"
if errorlevel 1 goto :fail

py -3 parentis_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Parentis Health" "parentis_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "Parentis Health" "parentis_pull.py" 0

echo.
echo ========================================
echo Updating Britelift QA data from Google Sheets
echo ========================================
cd /d "%BRITELIFT_DIR%"
if errorlevel 1 goto :fail

py -3 britelift_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Britelift" "britelift_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "Britelift" "britelift_pull.py" 0

echo.
echo ========================================
echo Updating Britelift Chat QA data from Google Sheets
echo ========================================
cd /d "%BLC_DIR%"
if errorlevel 1 goto :fail

py -3 britelift_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Britelift Chat" "britelift_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "Britelift Chat" "britelift_pull.py" 0

echo.
echo ========================================
echo Updating RideX QA data from Google Sheets
echo ========================================
cd /d "%RIDEX_DIR%"
if errorlevel 1 goto :fail

py -3 Ridex_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "RideX" "Ridex_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "RideX" "Ridex_pull.py" 0

echo.
echo ========================================
echo Updating Hamilton QA data from Google Sheets
echo ========================================
cd /d "%HAMILTON_DIR%"
if errorlevel 1 goto :fail

py -3 Hamilton_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Hamilton" "Hamilton_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "Hamilton" "Hamilton_pull.py" 0

echo.
echo ========================================
echo Updating Skyline QA data from Google Sheets
echo ========================================
cd /d "%SKYLINE_DIR%"
if errorlevel 1 goto :fail

py -3 Skyline_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Skyline" "Skyline_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "Skyline" "Skyline_pull.py" 0

echo.
echo ========================================
echo Updating VIP QA data from Google Sheets
echo ========================================
cd /d "%VIP_DIR%"
if errorlevel 1 goto :fail

py -3 vip_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "VIP" "vip_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "VIP" "vip_pull.py" 0

echo.
echo ========================================
echo Updating C^&H QA data from Google Sheets
echo ========================================
cd /d "%CH_DIR%"
if errorlevel 1 goto :fail

py -3 ch_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "C&H" "ch_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "C&H" "ch_pull.py" 0

echo.
echo ========================================
echo Updating Reno Cab QA data from Google Sheets
echo ========================================
cd /d "%RC_DIR%"
if errorlevel 1 goto :fail

py -3 rc_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Reno Cab" "rc_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "Reno Cab" "rc_pull.py" 0

echo.
echo ========================================
echo Updating Trans Iowa QA data from Google Sheets
echo ========================================
cd /d "%TI_DIR%"
if errorlevel 1 goto :fail

py -3 ti_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Trans Iowa" "ti_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "Trans Iowa" "ti_pull.py" 0

echo.
echo ========================================
echo Updating Data Carz QA data from Google Sheets
echo ========================================
cd /d "%DC_DIR%"
if errorlevel 1 goto :fail

py -3 dc_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Data Carz" "dc_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "Data Carz" "dc_pull.py" 0

echo.
echo ========================================
echo Updating Associated Cab QA data from Google Sheets
echo ========================================
cd /d "%AC_DIR%"
if errorlevel 1 goto :fail

py -3 ac_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Associated Cab" "ac_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "Associated Cab" "ac_pull.py" 0

echo.
echo ========================================
echo Updating Ollies QA data from Google Sheets
echo ========================================
cd /d "%OL_DIR%"
if errorlevel 1 goto :fail

py -3 ol_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Ollies" "ol_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "Ollies" "ol_pull.py" 0

echo.
echo ========================================
echo Updating Circle Taxi QA data from Google Sheets
echo ========================================
cd /d "%CT_DIR%"
if errorlevel 1 goto :fail

py -3 ct_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Circle Taxi" "ct_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "Circle Taxi" "ct_pull.py" 0

echo.
echo ========================================
echo Updating YCOV QA data from Google Sheets
echo ========================================
cd /d "%YCOV_DIR%"
if errorlevel 1 goto :fail

py -3 ycov_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "YCOV" "ycov_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "YCOV" "ycov_pull.py" 0

echo.
echo ========================================
echo Updating Kelowna QA data from Google Sheets
echo ========================================
cd /d "%KEL_DIR%"
if errorlevel 1 goto :fail

py -3 kel_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Kelowna" "kel_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "Kelowna" "kel_pull.py" 0

echo.
echo ========================================
echo Updating Vermont QA data from Google Sheets
echo ========================================
cd /d "%VT_DIR%"
if errorlevel 1 goto :fail

py -3 vt_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Vermont" "vt_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "Vermont" "vt_pull.py" 0

echo.
echo ========================================
echo Updating YCDC QA data from Google Sheets
echo ========================================
cd /d "%YCDC_DIR%"
if errorlevel 1 goto :fail

py -3 ycdc_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "YCDC" "ycdc_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "YCDC" "ycdc_pull.py" 0

echo.
echo ========================================
echo Updating Blueline QA data from Google Sheets
echo ========================================
cd /d "%BL_DIR%"
if errorlevel 1 goto :fail

py -3 bl_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Blueline" "bl_pull.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "Blueline" "bl_pull.py" 0

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
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Build" "dashboard.py" 1
    goto :fail
)
py -3 "%MASTERLIST_DIR%\log_step.py" step "Build" "dashboard.py" 0

echo.
echo ========================================
echo Publishing to GitHub
echo ========================================
git add masterlist_dashboard.html update_coaching_dashboard.bat update_coaching_dashboard_auto.bat
if errorlevel 1 goto :fail

git diff --cached --quiet
if not errorlevel 1 (
    echo No dashboard changes to publish.
    goto :publish_monitor
)

git commit -m "Update coaching dashboard"
if errorlevel 1 goto :fail

git pull --rebase --autostash
if errorlevel 1 goto :fail

git push
if errorlevel 1 goto :fail

py -3 "%MASTERLIST_DIR%\log_step.py" step "Git Push" "git push" 0

:publish_monitor
py -3 "%MASTERLIST_DIR%\log_step.py" finish success
py -3 "%MASTERLIST_DIR%\generate_monitor.py"
set MONITOR_DONE=1
git add pipeline_status.json pipeline_monitor.html
git diff --cached --quiet
if errorlevel 1 (
    git commit -m "Update pipeline monitor"
    git pull --rebase --autostash
    git push
)

:done
echo.
echo Done. GitHub Pages may take about 1 minute to update.
exit /b 0

:fail
echo.
echo Update failed. Check the message above.
if "%MONITOR_DONE%"=="0" (
    cd /d "%MASTERLIST_DIR%" 2>nul
    py -3 "%MASTERLIST_DIR%\log_step.py" finish failed 2>nul
    py -3 "%MASTERLIST_DIR%\generate_monitor.py" 2>nul
    git add pipeline_status.json pipeline_monitor.html 2>nul
    git diff --cached --quiet 2>nul
    if errorlevel 1 (
        git commit -m "Update pipeline monitor -- run failed" 2>nul
        git pull --rebase --autostash 2>nul
        git push 2>nul
    )
)
exit /b 1
