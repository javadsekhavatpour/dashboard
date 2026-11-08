@echo off
title Summary Dashboard
echo ================================================================
echo    Summary Dashboard - Latest Day for All Symbols
echo ================================================================
echo.
echo Browser will open automatically at:
echo    http://localhost:8501
echo.
echo To stop: press Ctrl+C or close this window.
echo ================================================================
echo.
streamlit run summary.py
pause
