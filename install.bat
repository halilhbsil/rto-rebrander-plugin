@echo off
REM ============================================================
REM   RTO Rebrander - One-Click Installer (aEX Institute)
REM   Double-click this file to install everything.
REM ============================================================
REM   This file delegates the real work to install.ps1 (PowerShell).
REM   It exists so staff can double-click instead of running a
REM   PowerShell script directly (Windows is friendlier to .bat).
REM ============================================================

echo.
echo ============================================================
echo   RTO Rebrander - aEX Institute Internal Installer
echo ============================================================
echo.
echo This will set up the RTO Rebrander plugin on your laptop.
echo It will install Python (if needed) and prepare Claude Code.
echo.
echo You may see a security prompt - click Yes to continue.
echo.
pause

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"

echo.
echo ============================================================
echo Installer finished. Press any key to close this window.
echo ============================================================
pause >nul
