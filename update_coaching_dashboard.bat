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
set "ANY_ACCOUNT_FAILED=0"

cd /d "%MASTERLIST_DIR%"
py -3 "%MASTERLIST_DIR%\log_step.py" init

echo.
echo ========================================
echo Fetching Masterlist from Google Sheets
echo ========================================
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step masterlist_fetch.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Masterlist" "masterlist_fetch.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] Masterlist fetch failed -- continuing with cached data
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Masterlist" "masterlist_fetch.py" 0
)

echo.
echo ========================================
echo Sending movement notifications
echo ========================================
cd /d "%MASTERLIST_DIR%"
py -3 "%MASTERLIST_DIR%\check_movement_notifications.py"
if errorlevel 1 (
    echo [WARN] Movement notifications failed -- continuing
)

echo.
echo ========================================
echo Updating Coaching data from Asana
echo ========================================
cd /d "%COACHING_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step asana_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Coaching" "asana_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] Coaching failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Coaching" "asana_pull.py" 0
)

echo.
echo ========================================
echo Updating M7 QA data from Google Sheets
echo ========================================
cd /d "%M7_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step m7_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "M7" "m7_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] M7 failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "M7" "m7_pull.py" 0
)

echo.
echo ========================================
echo Updating Parentis Health QA data from Google Sheets
echo ========================================
cd /d "%PARENTIS_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step parentis_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Parentis Health" "parentis_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] Parentis Health failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Parentis Health" "parentis_pull.py" 0
)

echo.
echo ========================================
echo Updating Britelift QA data from Google Sheets
echo ========================================
cd /d "%BRITELIFT_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step britelift_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Britelift" "britelift_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] Britelift failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Britelift" "britelift_pull.py" 0
)

echo.
echo ========================================
echo Updating Britelift Chat QA data from Google Sheets
echo ========================================
cd /d "%BLC_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step britelift_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Britelift Chat" "britelift_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] Britelift Chat failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Britelift Chat" "britelift_pull.py" 0
)

echo.
echo ========================================
echo Updating RideX QA data from Google Sheets
echo ========================================
cd /d "%RIDEX_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step Ridex_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "RideX" "Ridex_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] RideX failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "RideX" "Ridex_pull.py" 0
)

echo.
echo ========================================
echo Updating Hamilton QA data from Google Sheets
echo ========================================
cd /d "%HAMILTON_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step Hamilton_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Hamilton" "Hamilton_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] Hamilton failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Hamilton" "Hamilton_pull.py" 0
)

echo.
echo ========================================
echo Updating Skyline QA data from Google Sheets
echo ========================================
cd /d "%SKYLINE_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step Skyline_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Skyline" "Skyline_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] Skyline failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Skyline" "Skyline_pull.py" 0
)

echo.
echo ========================================
echo Updating VIP QA data from Google Sheets
echo ========================================
cd /d "%VIP_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step vip_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "VIP" "vip_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] VIP failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "VIP" "vip_pull.py" 0
)

echo.
echo ========================================
echo Updating C^&H QA data from Google Sheets
echo ========================================
cd /d "%CH_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step ch_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "C&H" "ch_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] C^&H failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "C&H" "ch_pull.py" 0
)

echo.
echo ========================================
echo Updating Reno Cab QA data from Google Sheets
echo ========================================
cd /d "%RC_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step rc_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Reno Cab" "rc_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] Reno Cab failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Reno Cab" "rc_pull.py" 0
)

echo.
echo ========================================
echo Updating Trans Iowa QA data from Google Sheets
echo ========================================
cd /d "%TI_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step ti_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Trans Iowa" "ti_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] Trans Iowa failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Trans Iowa" "ti_pull.py" 0
)

echo.
echo ========================================
echo Updating Data Carz QA data from Google Sheets
echo ========================================
cd /d "%DC_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step dc_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Data Carz" "dc_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] Data Carz failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Data Carz" "dc_pull.py" 0
)

echo.
echo ========================================
echo Updating Associated Cab QA data from Google Sheets
echo ========================================
cd /d "%AC_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step ac_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Associated Cab" "ac_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] Associated Cab failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Associated Cab" "ac_pull.py" 0
)

echo.
echo ========================================
echo Updating Ollies QA data from Google Sheets
echo ========================================
cd /d "%OL_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step ol_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Ollies" "ol_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] Ollies failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Ollies" "ol_pull.py" 0
)

echo.
echo ========================================
echo Updating Circle Taxi QA data from Google Sheets
echo ========================================
cd /d "%CT_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step ct_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Circle Taxi" "ct_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] Circle Taxi failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Circle Taxi" "ct_pull.py" 0
)

echo.
echo ========================================
echo Updating YCOV QA data from Google Sheets
echo ========================================
cd /d "%YCOV_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step ycov_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "YCOV" "ycov_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] YCOV failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "YCOV" "ycov_pull.py" 0
)

echo.
echo ========================================
echo Updating Kelowna QA data from Google Sheets
echo ========================================
cd /d "%KEL_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step kel_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Kelowna" "kel_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] Kelowna failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Kelowna" "kel_pull.py" 0
)

echo.
echo ========================================
echo Updating Vermont QA data from Google Sheets
echo ========================================
cd /d "%VT_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step vt_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Vermont" "vt_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] Vermont failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Vermont" "vt_pull.py" 0
)

echo.
echo ========================================
echo Updating YCDC QA data from Google Sheets
echo ========================================
cd /d "%YCDC_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step ycdc_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "YCDC" "ycdc_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] YCDC failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "YCDC" "ycdc_pull.py" 0
)

echo.
echo ========================================
echo Updating Blueline QA data from Google Sheets
echo ========================================
cd /d "%BL_DIR%"
py -3 "%MASTERLIST_DIR%\self_heal.py" run-step bl_pull.py
if errorlevel 1 (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Blueline" "bl_pull.py" 1
    set "ANY_ACCOUNT_FAILED=1"
    echo [WARN] Blueline failed -- continuing
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" step "Blueline" "bl_pull.py" 0
)

echo.
echo ========================================
echo Rebuilding dashboard
echo ========================================
cd /d "%MASTERLIST_DIR%"

echo.
echo ========================================
echo Self-heal: checking for count drops
echo ========================================
py -3 "%MASTERLIST_DIR%\self_heal.py" heal-drops

echo Syncing latest dashboard repo changes...
git pull --rebase --autostash
if errorlevel 1 goto :fail

py -3 "%MASTERLIST_DIR%\self_heal.py" run-step dashboard.py
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
if "%ANY_ACCOUNT_FAILED%"=="1" (
    py -3 "%MASTERLIST_DIR%\log_step.py" finish partial
) else (
    py -3 "%MASTERLIST_DIR%\log_step.py" finish success
)
py -3 "%MASTERLIST_DIR%\generate_monitor.py"
set MONITOR_DONE=1
git add pipeline_status.json pipeline_monitor.html pipeline_log.json pipeline_rowcount_baseline.json pipeline_highvol_notified.json pipeline_drops_notified.json
git diff --cached --quiet
if errorlevel 1 (
    git commit -m "Update pipeline monitor"
    git pull --rebase --autostash
    git push
)

:done
echo.
if "%ANY_ACCOUNT_FAILED%"=="1" (
    echo Done with warnings -- some accounts failed, check pipeline monitor.
) else (
    echo Done. GitHub Pages may take about 1 minute to update.
)
pause
exit /b 0

:fail
echo.
echo Pipeline failed at a critical step. Check the message above.
if "%MONITOR_DONE%"=="0" (
    cd /d "%MASTERLIST_DIR%" 2>nul
    py -3 "%MASTERLIST_DIR%\log_step.py" finish failed 2>nul
    py -3 "%MASTERLIST_DIR%\generate_monitor.py" 2>nul
    git add pipeline_status.json pipeline_monitor.html pipeline_log.json pipeline_rowcount_baseline.json pipeline_highvol_notified.json pipeline_drops_notified.json 2>nul
    git diff --cached --quiet 2>nul
    if errorlevel 1 (
        git commit -m "Update pipeline monitor -- run failed" 2>nul
        git pull --rebase --autostash 2>nul
        git push 2>nul
    )
)
pause
exit /b 1
