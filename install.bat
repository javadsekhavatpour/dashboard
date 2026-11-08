@echo off
title TSE Screener - Install
echo ================================================================
echo    TSE Screener - Installing dependencies
echo ================================================================
echo.
echo Installing: streamlit, requests, openpyxl
echo.
pip install streamlit requests openpyxl
echo.
echo ================================================================
if %ERRORLEVEL% EQU 0 (
    echo OK - Installation successful!
    echo.
    echo Now double-click: run.bat
) else (
    echo ERROR - Installation failed!
    echo.
    echo Try manually:
    echo     pip install streamlit requests openpyxl
)
echo ================================================================
pause
