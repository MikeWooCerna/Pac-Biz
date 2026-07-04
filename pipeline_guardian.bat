@echo off
cd /d "C:\Users\Mike Woo Cerna\Documents\PB\Masterlist"

echo Running Pac-Biz Pipeline Guardian...
py -3 pipeline_guardian.py --fix --push --live

echo.
echo Done.
pause
