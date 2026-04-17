@echo off
title Deflection Check Tool - SAP2000
echo ============================================
echo   Deflection Check Tool - SAP2000
echo ============================================
echo.
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found.
    pause
    exit /b 1
)
echo Checking dependencies...
pip show comtypes >nul 2>&1 || pip install comtypes
pip show openpyxl >nul 2>&1 || pip install openpyxl
echo.
echo Starting... Make sure SAP2000 is running.
python "%~dp0main.py"
if errorlevel 1 pause
