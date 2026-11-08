@echo off
title Build All Dashboards
echo ================================================================
echo    Building dashboards for all input files
echo    Input folder:  input\
echo    Output folder:  dashboard\
echo ================================================================
echo.

REM Create folders if not exist
if not exist "input" mkdir input
if not exist "dashboard" mkdir dashboard

REM Process each *_input.xlsx file in input\ folder
for %%f in (input\*_input.xlsx) do (
    set "input=%%f"
    set "name=%%~nf"
    setlocal enabledelayedexpansion
    set "outname=!name:_input=!"
    set "output=dashboard\!outname!_dashboard.xlsx"
    echo Processing: !input!  -^>  !output!
    python make_dashboard.py "!input!" --output "!output!" --name "!outname!"
    endlocal
    echo.
)

echo ================================================================
echo Done! All dashboards saved in: dashboard\
echo ================================================================
pause
