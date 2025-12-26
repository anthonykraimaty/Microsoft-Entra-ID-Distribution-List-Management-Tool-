@echo off
title Distribution List Manager - Debug Mode
cd /d "%~dp0"

echo ============================================
echo  Distribution List Manager - DEBUG MODE
echo ============================================
echo.
echo Starting application with console output...
echo Logs will appear below:
echo --------------------------------------------

python -u gui.py 2>&1

echo.
echo --------------------------------------------
echo Application closed.
if errorlevel 1 (
    echo [ERROR] Application exited with errors.
)
echo.

