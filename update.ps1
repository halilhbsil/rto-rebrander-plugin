# =============================================================
#  RTO Rebrander - Updater (aEX Institute internal use)
# =============================================================
#  This PowerShell script is launched by update.bat. It:
#    1. Verifies Claude Code Desktop is installed (so `claude` CLI exists)
#    2. Pulls the latest plugin manifest from the GitHub marketplace
#    3. Installs the new plugin version locally
#    4. Reminds the staff member to restart Claude Code Desktop
#
#  Staff don't run this directly - they double-click update.bat
#  which sets the right execution policy and delegates to here.
# =============================================================

$ErrorActionPreference = "Stop"
$Marketplace = "aex-internal"
$PluginName  = "rto-rebrander"

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
if ($LASTEXITCODE -ne 0) {
    Write-Fail "Could not refresh the marketplace. Output was:"
    $mpOutput | ForEach-Object { Write-Host "   $_" -ForegroundColor DarkGray }
    Write-Fail ""
    Write-Fail "This usually means the rebrander hasn't been installed yet."
    Write-Fail "Please run install.bat first, then try this updater again."
    exit 1
}
Write-OK "Marketplace refreshed"

# -------------------------------------------------------------
# 3. Update the plugin to the latest version
# -------------------------------------------------------------
Write-Step "Installing the latest plugin version"

$updateOutput = & claude plugin update $PluginName 2>&1
$updateOutput | ForEach-Object { Write-Host "   $_" -ForegroundColor DarkGray }

if ($LASTEXITCODE -ne 0) {
    Write-Fail "Plugin update failed. See output above."
    Write-Fail "Email the screen contents to halil.houssein@aexinstitute.com.au"
    exit 1
}
Write-OK "Plugin updated to the latest version"

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
