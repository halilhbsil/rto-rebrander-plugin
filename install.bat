@echo off
REM ============================================================
REM   RTO Rebrander - One-Click Installer (aEX Institute)
REM
REM   Self-contained: PowerShell logic embedded below a marker.
REM   The batch wrapper extracts it to a temp file, runs it,
REM   then deletes the temp. cmd.exe never sees the PowerShell
REM   content because of the 'exit /b' before the marker.
REM ============================================================

setlocal
echo.
echo ============================================================
echo   RTO Rebrander - aEX Institute Internal Installer
echo ============================================================
echo.
echo This will install the RTO Rebrander plugin for Claude Code.
echo It will:
echo   - Install Python if needed
echo   - Register the aEX marketplace with Claude Code
echo   - Install the rto-rebrander plugin
echo.
echo You may see a security prompt - click Yes to continue.
echo.
pause

set "PS_TMP=%TEMP%\rto-rebrander-install-%RANDOM%-%TIME:~6,2%%TIME:~9,2%.ps1"

REM Extract the embedded PowerShell body to PS_TMP
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
  "$lines = Get-Content -LiteralPath '%~f0'; " ^
  "$hit = $lines | Select-String -Pattern '^# === PS_BODY_BEGINS ===$' | Select-Object -First 1; " ^
  "if (-not $hit) { Write-Error 'Embedded script marker not found'; exit 1 }; " ^
  "$body = $lines[$hit.LineNumber..($lines.Length - 1)] -join \"`r`n\"; " ^
  "[System.IO.File]::WriteAllText('%PS_TMP%', $body, (New-Object System.Text.UTF8Encoding $false))"

if errorlevel 1 (
    echo.
    echo Failed to prepare the installer. Please email halil.houssein@aexinstitute.com.au
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_TMP%"
set "PSEXIT=%ERRORLEVEL%"

del "%PS_TMP%" 2>nul

echo.
echo ============================================================
echo Installer finished. Press any key to close this window.
echo ============================================================
pause >nul
endlocal & exit /b %PSEXIT%

REM ============================================================
REM Everything below is PowerShell, not batch.
REM ============================================================

# === PS_BODY_BEGINS ===
# =============================================================
#  RTO Rebrander - Installer (aEX Institute internal use)
# =============================================================
#  This script is normally embedded in install.bat. The bat
#  wrapper extracts it to a temp file and runs it.
#
#  Steps:
#    1. Make sure Python 3.10+ is installed (winget install if not)
#    2. Make sure Claude Code Desktop is installed (claude on PATH)
#    3. Remove any old skill-style install
#    4. Add the aEX marketplace via `claude plugin marketplace add`
#    5. Install the plugin via `claude plugin install`
#    6. Tell the staff member to restart Claude Code Desktop
# =============================================================

$ErrorActionPreference = "Stop"
$PluginRepo  = "halilhbsil/rto-rebrander-plugin"
$Marketplace = "aex-internal"
$PluginName  = "rto-rebrander"

function Write-Step { param([string]$Message) Write-Host "`n>> $Message" -ForegroundColor Cyan }
function Write-OK   { param([string]$Message) Write-Host "   [OK]   $Message" -ForegroundColor Green }
function Write-Skip { param([string]$Message) Write-Host "   [SKIP] $Message" -ForegroundColor DarkGray }
function Write-Fail { param([string]$Message) Write-Host "   [FAIL] $Message" -ForegroundColor Red }

# -------------------------------------------------------------
# 1. Check / install Python 3.10+
# -------------------------------------------------------------
Write-Step "Checking for Python 3.10 or newer"

function Test-PythonOK {
    try {
        $verRaw = & python --version 2>&1
        if ($LASTEXITCODE -ne 0) { return $false }
        if ($verRaw -match "Python (\d+)\.(\d+)") {
            $major = [int]$Matches[1]; $minor = [int]$Matches[2]
            return (($major -gt 3) -or ($major -eq 3 -and $minor -ge 10))
        }
    } catch { }
    return $false
}

if (Test-PythonOK) {
    Write-OK "Python 3.10+ already installed"
} else {
    Write-Host "   Python 3.10+ not found. Installing via winget..."
    try {
        winget install --id Python.Python.3.12 `
            --silent --accept-package-agreements --accept-source-agreements `
            --scope user | Out-Host
        # Refresh PATH for this session so the new python is visible
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" +
                    [System.Environment]::GetEnvironmentVariable("Path","User")
        if (Test-PythonOK) {
            Write-OK "Python installed successfully"
        } else {
            Write-Fail "Python install ran but version check still fails. Please install manually from https://www.python.org/downloads/ and re-run this installer."
            exit 1
        }
    } catch {
        Write-Fail "winget unavailable. Please install Python 3.10+ manually from https://www.python.org/downloads/"
        Write-Fail "(make sure to tick 'Add Python to PATH' during install)"
        exit 1
    }
}

# -------------------------------------------------------------
# 2. Verify Claude Code Desktop is installed
# -------------------------------------------------------------
Write-Step "Checking that Claude Code Desktop is installed"

$claudeCmd = Get-Command claude -ErrorAction SilentlyContinue
if (-not $claudeCmd) {
    Write-Fail "The 'claude' command was not found on your laptop."
    Write-Fail "Claude Code Desktop must be installed before running this installer."
    Write-Fail "Please install it from https://claude.com/code first, then re-run install.bat."
    exit 1
}
Write-OK "Found Claude Code at $($claudeCmd.Source)"

# -------------------------------------------------------------
# 3. Remove old skill-style install if it exists
# -------------------------------------------------------------
Write-Step "Cleaning up any previous skill-style install"

$OldSkill = Join-Path $env:USERPROFILE ".claude\skills\rto-document-rebrander"
if (Test-Path $OldSkill) {
    Remove-Item -Recurse -Force $OldSkill
    Write-OK "Removed old skill at $OldSkill"
} else {
    Write-Skip "No old skill folder found - nothing to clean up"
}

# -------------------------------------------------------------
# 4. Add the aEX marketplace to Claude Code
# -------------------------------------------------------------
Write-Step "Registering the aEX marketplace"

$mpOutput = & claude plugin marketplace add $PluginRepo 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Fail "Could not register the marketplace. Output:"
    $mpOutput | ForEach-Object { Write-Host "   $_" -ForegroundColor DarkGray }
    Write-Fail "Please email halil.houssein@aexinstitute.com.au with this screen."
    exit 1
}
$mpOutput | ForEach-Object { Write-Host "   $_" -ForegroundColor DarkGray }
Write-OK "Marketplace '$Marketplace' registered"

# -------------------------------------------------------------
# 5. Install the plugin
# -------------------------------------------------------------
Write-Step "Installing the rto-rebrander plugin"

$installOutput = & claude plugin install "$PluginName@$Marketplace" 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Fail "Could not install the plugin. Output:"
    $installOutput | ForEach-Object { Write-Host "   $_" -ForegroundColor DarkGray }
    Write-Fail "Please email halil.houssein@aexinstitute.com.au with this screen."
    exit 1
}
$installOutput | ForEach-Object { Write-Host "   $_" -ForegroundColor DarkGray }
Write-OK "Plugin installed and enabled"

# -------------------------------------------------------------
# 6. Final instructions
# -------------------------------------------------------------
Write-Host ""
Write-Host "=============================================================" -ForegroundColor Yellow
Write-Host "  All done!" -ForegroundColor Yellow
Write-Host "=============================================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "ONE LAST STEP:" -ForegroundColor Yellow
Write-Host ""
Write-Host "  1. Close Claude Code Desktop completely"
Write-Host "  2. Open it again"
Write-Host "  3. The first time, you may see a one-time message about"
Write-Host "     'installing Python packages' - this is normal and quick."
Write-Host "  4. You can now ask Claude to rebrand documents!"
Write-Host ""
Write-Host "Example: just type or say something like:"
Write-Host "  'Rebrand the docs in C:\BSI Documents to aEX branding'" -ForegroundColor White
Write-Host ""
Write-Host "If anything goes wrong, email halil.houssein@aexinstitute.com.au"
Write-Host ""
