@echo off
title Distribution List Manager - Setup
cd /d "%~dp0"

echo ============================================
echo  Distribution List Manager - Setup
echo ============================================
echo.

echo Installing Python dependencies...
pip install -r requirements.txt

echo.
echo ============================================
echo  Setup complete!
echo ============================================
echo.
echo Next steps:
echo   1. Copy .env.example to .env
echo   2. Edit .env with your Azure AD credentials
echo   3. Run: run.bat
echo.
pause
