@echo off
REM ============================================================
REM   RTO Rebrander - One-Click Updater (aEX Institute)
REM   Double-click this file to update to the latest version.
REM ============================================================
REM   Delegates to update.ps1 (PowerShell does the real work).
REM   This .bat exists so staff can double-click without having
REM   to deal with PowerShell execution policies.
REM ============================================================

echo.
echo ============================================================
echo   RTO Rebrander - aEX Institute Internal Updater
echo ============================================================
echo.
echo This will update the RTO Rebrander plugin to the latest version.
echo Make sure Claude Code Desktop is CLOSED before continuing.
echo.
pause

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0update.ps1"

echo.
echo ============================================================
echo Updater finished. Press any key to close this window.
echo ============================================================
pause >nul
