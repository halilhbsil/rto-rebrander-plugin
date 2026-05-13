@echo off
REM ============================================================
REM   RTO Rebrander - One-Click Updater (aEX Institute)
REM
REM   Self-contained: PowerShell logic embedded at the bottom of
REM   this file. See install.bat for how the self-extracting
REM   pattern works (same design).
REM ============================================================

setlocal
echo.
echo ============================================================
echo   RTO Rebrander - aEX Institute Internal Updater
echo ============================================================
echo.
echo This will update the RTO Rebrander plugin to the latest version.
echo Make sure Claude Code Desktop is CLOSED before continuing.
echo.
pause

set "PS_TMP=%TEMP%\rto-rebrander-update-%RANDOM%-%TIME:~6,2%%TIME:~9,2%.ps1"

REM Extract the embedded PowerShell body to PS_TMP
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
  "$lines = Get-Content -LiteralPath '%~f0'; " ^
  "$hit = $lines | Select-String -Pattern '^# === PS_BODY_BEGINS ===$' | Select-Object -First 1; " ^
  "if (-not $hit) { Write-Error 'Embedded script marker not found'; exit 1 }; " ^
  "$body = $lines[$hit.LineNumber..($lines.Length - 1)] -join \"`r`n\"; " ^
  "[System.IO.File]::WriteAllText('%PS_TMP%', $body, (New-Object System.Text.UTF8Encoding $false))"

if errorlevel 1 (
    echo.
    echo Failed to prepare the updater. Please email halil.houssein@aexinstitute.com.au
    pause
    exit /b 1
)

REM Run the extracted PowerShell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_TMP%"
set "PSEXIT=%ERRORLEVEL%"

REM Cleanup
del "%PS_TMP%" 2>nul

echo.
echo ============================================================
echo Updater finished. Press any key to close this window.
echo ============================================================
pause >nul
endlocal & exit /b %PSEXIT%

REM ============================================================
REM Everything below is PowerShell, not batch.
REM ============================================================

# === PS_BODY_BEGINS ===
# =============================================================
#  RTO Rebrander - Updater (aEX Institute internal use)
# =============================================================

# Use Continue (not Stop) so native command stderr doesn't kill
# the script before our explicit $LASTEXITCODE checks run.
$ErrorActionPreference = "Continue"
$Marketplace = "aex-internal"
$PluginName  = "rto-rebrander"
$PluginQualified = "${PluginName}@${Marketplace}"

function Write-Step { param([string]$Message) Write-Host "`n>> $Message" -ForegroundColor Cyan }
function Write-OK   { param([string]$Message) Write-Host "   [OK]   $Message" -ForegroundColor Green }
function Write-Fail { param([string]$Message) Write-Host "   [FAIL] $Message" -ForegroundColor Red }

# -------------------------------------------------------------
# 1. Confirm Claude Code Desktop is installed
# -------------------------------------------------------------
Write-Step "Checking that Claude Code Desktop is installed"

$claudeCmd = Get-Command claude -ErrorAction SilentlyContinue
if (-not $claudeCmd) {
    Write-Fail "The 'claude' command was not found on your laptop."
    Write-Fail "Claude Code Desktop must be installed before you can update the plugin."
    Write-Fail "Please install Claude Code Desktop from https://claude.com/code first."
    exit 1
}
Write-OK "Found Claude Code at $($claudeCmd.Source)"

# -------------------------------------------------------------
# 2. Pull the latest marketplace manifest
# -------------------------------------------------------------
Write-Step "Fetching the latest version info from GitHub"

$mpOutput = & claude plugin marketplace update $Marketplace 2>&1
$mpExit = $LASTEXITCODE
$mpOutput | ForEach-Object { Write-Host "   $_" -ForegroundColor DarkGray }
if ($mpExit -ne 0) {
    Write-Fail ""
    Write-Fail "Could not refresh the marketplace."
    Write-Fail "This usually means the rebrander hasn't been installed yet."
    Write-Fail "Please run install.bat first, then try this updater again."
    exit 1
}
Write-OK "Marketplace refreshed"

# -------------------------------------------------------------
# 3. Update the plugin to the latest version
# -------------------------------------------------------------
Write-Step "Installing the latest plugin version"

$updateOutput = & claude plugin update $PluginQualified 2>&1
$updateExit = $LASTEXITCODE
$updateOutput | ForEach-Object { Write-Host "   $_" -ForegroundColor DarkGray }

if ($updateExit -ne 0) {
    Write-Fail "Plugin update failed. See output above."
    Write-Fail "Email the screen contents to halil.houssein@aexinstitute.com.au"
    exit 1
}
Write-OK "Plugin update check complete"

# -------------------------------------------------------------
# 4. Final instructions
# -------------------------------------------------------------
Write-Host ""
Write-Host "=============================================================" -ForegroundColor Yellow
Write-Host "  Update complete!" -ForegroundColor Yellow
Write-Host "=============================================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "ONE LAST STEP - please do this manually:" -ForegroundColor Yellow
Write-Host ""
Write-Host "  1. Make sure Claude Code Desktop is CLOSED"
Write-Host "  2. Open it again"
Write-Host "  3. The new version will be active automatically"
Write-Host ""
Write-Host "If anything looks wrong, email halil.houssein@aexinstitute.com.au"
Write-Host ""
