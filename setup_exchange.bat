@echo off
title Exchange Online Setup
cd /d "%~dp0"

echo ============================================
echo  Exchange Online Setup for DL Manager
echo ============================================
echo.
echo This will configure Exchange Online access
echo to manage distribution lists from the app.
echo.
echo Press any key to continue...
pause > nul

powershell -ExecutionPolicy Bypass -File "%~dp0setup_exchange.ps1"
