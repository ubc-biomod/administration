@echo off
title UBC BioMod Email Sender
cd /d "%~dp0"

echo.
echo ============================================
echo   UBC BioMod Email Sender
echo ============================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo Could not find Python on this computer.
    echo.
    echo Please install Python first. Open the README.md file in this
    echo folder and follow the "Install Python" steps.
    echo.
    pause
    exit /b 1
)

python email.py
echo.
pause
