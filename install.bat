@echo off
REM ============================================================
REM   RTO Rebrander - One-Click Installer (aEX Institute)
REM
REM   This file is self-contained. The PowerShell logic that does
REM   the real work is embedded at the bottom of this same file.
REM   The batch wrapper extracts it to a temp file, runs it, and
REM   cleans up. No companion files needed.
REM
REM   How it works (for anyone reading the source):
REM     - The line "# === PS_BODY_BEGINS ===" near the bottom is
REM       a marker. Everything below it is PowerShell.
REM     - The batch portion reads its own file (%~f0), finds the
REM       marker line, and writes the PowerShell tail to a
REM       randomly-named temp file under %TEMP%.
REM     - It then invokes powershell.exe on that temp file and
REM       deletes the temp afterwards.
REM     - cmd.exe never sees the PowerShell content because of
REM       the "exit /b" below the cleanup.
REM ============================================================

setlocal
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

REM Run the extracted PowerShell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_TMP%"
set "PSEXIT=%ERRORLEVEL%"

REM Cleanup
del "%PS_TMP%" 2>nul

echo.
echo ============================================================
echo Installer finished. Press any key to close this window.
echo ============================================================
pause >nul
endlocal & exit /b %PSEXIT%

REM ============================================================
REM Everything below is PowerShell, not batch. cmd.exe never
REM reaches it because of the "exit /b" above.
REM ============================================================

# === PS_BODY_BEGINS ===
# =============================================================
#  RTO Rebrander - Installer (aEX Institute internal use)
# =============================================================
#  This script is normally embedded in install.bat and extracted
#  to a temp file at runtime. See install.bat for the wrapper.
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
        # Refresh PATH for this session so the new python is found below
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
# 2. Remove old skill-style install if it exists
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
# 3. Register the marketplace in Claude Code settings.json
# -------------------------------------------------------------
Write-Step "Registering the aEX marketplace with Claude Code"

$ClaudeDir   = Join-Path $env:USERPROFILE ".claude"
$SettingsPath = Join-Path $ClaudeDir "settings.json"
if (-not (Test-Path $ClaudeDir)) {
    New-Item -ItemType Directory -Path $ClaudeDir | Out-Null
}

# Load existing settings if any, else start fresh
if (Test-Path $SettingsPath) {
    try {
        $settings = Get-Content $SettingsPath -Raw | ConvertFrom-Json
    } catch {
        Write-Fail "Could not parse existing $SettingsPath. Please show this file to Halil and re-run."
        exit 1
    }
} else {
    $settings = New-Object PSObject
}

# Ensure extraKnownMarketplaces exists and add our entry
if (-not ($settings.PSObject.Properties.Name -contains "extraKnownMarketplaces")) {
    $settings | Add-Member -NotePropertyName "extraKnownMarketplaces" -NotePropertyValue (New-Object PSObject)
}
$existing = $settings.extraKnownMarketplaces.PSObject.Properties.Name
if ($existing -contains $Marketplace) {
    Write-Skip "Marketplace '$Marketplace' already registered"
} else {
    $entry = [PSCustomObject]@{
        source = [PSCustomObject]@{
            source = "github"
            repo   = $PluginRepo
        }
    }
    $settings.extraKnownMarketplaces | Add-Member -NotePropertyName $Marketplace -NotePropertyValue $entry
    Write-OK "Added '$Marketplace' marketplace pointing at $PluginRepo"
}

# Save settings.json (UTF-8, no BOM)
$json = $settings | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText($SettingsPath, $json, (New-Object System.Text.UTF8Encoding $false))
Write-OK "Saved $SettingsPath"

# -------------------------------------------------------------
# 4. Final instructions for the staff member
# -------------------------------------------------------------
Write-Host ""
Write-Host "=============================================================" -ForegroundColor Yellow
Write-Host "  Setup complete!" -ForegroundColor Yellow
Write-Host "=============================================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "ONE LAST STEP - please do this manually:" -ForegroundColor Yellow
Write-Host ""
Write-Host "  1. Open Claude Code Desktop"
Write-Host "  2. In the chat input area, type:"
Write-Host ""
Write-Host "       /plugin install $PluginName@$Marketplace" -ForegroundColor White
Write-Host ""
Write-Host "  3. Wait for it to finish installing"
Write-Host "  4. Close and re-open Claude Code Desktop"
Write-Host "  5. You can now ask Claude to rebrand documents!"
Write-Host ""
Write-Host "If anything goes wrong, email halil.houssein@aexinstitute.com.au"
Write-Host ""
