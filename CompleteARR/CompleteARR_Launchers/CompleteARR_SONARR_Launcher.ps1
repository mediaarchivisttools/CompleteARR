#Requires -Version 7.0
<#
    CompleteARR_SONARR_Launcher.ps1
    ---------------------------------------------
    Orchestrates the CompleteARR SONARR suite:

      1) Runs the AutoSorter:
         .\CompleteARR_Scripts\CompleteARR_SONARR_AutoSorter.ps1

      2) Runs the main CompleteARR engine:
         .\CompleteARR_Scripts\CompleteARR_SONARR_SeriesEngine.ps1

    All scripts resolve paths relative to the CompleteARR root
    (the folder this launcher lives in), so the whole directory
    can be moved to any machine/location and still work.
#>

[CmdletBinding()]
param(
    # Optional override for the settings file. If omitted, the engine
    # will use CompleteARR_Settings.yml in the root folder.
    [string]$ConfigPath
)

# ------------------------------------------------------------
# Resolve paths
# ------------------------------------------------------------

$CompleteArrRoot        = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$CompleteArrRoot        = Split-Path -Path $CompleteArrRoot -Parent
$AutoSorterScriptPath   = Join-Path $CompleteArrRoot 'CompleteARR_Scripts\CompleteARR_SONARR_AutoSorter.ps1'
$EngineScriptPath       = Join-Path $CompleteArrRoot 'CompleteARR_Scripts\CompleteARR_SONARR_SeriesEngine.ps1'
$DefaultSettingsPath    = Join-Path $CompleteArrRoot 'CompleteARR_Settings' 'CompleteARR_SONARR_Settings.yml'

if (-not $ConfigPath) {
    $ConfigPath = $DefaultSettingsPath
}

Write-Host "============================================" -ForegroundColor Cyan
Write-Host " 🎬 CompleteARR SONARR Launcher - Starting Up!" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "📁 Working from folder: $CompleteArrRoot" -ForegroundColor Cyan
Write-Host "⚙️  Using settings file: $ConfigPath" -ForegroundColor Cyan
Write-Host ""
Write-Host "🔍 Checking your files and settings..." -ForegroundColor Cyan
Write-Host ""

# Quick checks so errors are obvious
if (-not (Test-Path -LiteralPath $AutoSorterScriptPath)) {
    Write-Host "ERROR: Could not find AutoSorter script at:" -ForegroundColor Red
    Write-Host "  $AutoSorterScriptPath" -ForegroundColor Red
    Read-Host "Press ENTER to close this window"
    exit 1
}

if (-not (Test-Path -LiteralPath $EngineScriptPath)) {
    Write-Host "ERROR: Could not find Engine script at:" -ForegroundColor Red
    Write-Host "  $EngineScriptPath" -ForegroundColor Red
    Read-Host "Press ENTER to close this window"
    exit 1
}

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    Write-Host "WARNING: Config file not found at:" -ForegroundColor Yellow
    Write-Host "  $ConfigPath" -ForegroundColor Yellow
    Write-Host "Engine may fail to start if it relies on this path." -ForegroundColor Yellow
    Write-Host ""
}

# ------------------------------------------------------------
# Run AutoSorter (same host, so logs appear in this window)
# ------------------------------------------------------------

$overallSuccess = $true

Write-Host "------------------------------------------------" -ForegroundColor Yellow
Write-Host " 🔄 STEP 1: Running AutoSorter" -ForegroundColor Yellow
Write-Host "------------------------------------------------" -ForegroundColor Yellow
Write-Host ""
Write-Host "📋 The AutoSorter will:" -ForegroundColor Yellow
Write-Host "   • Sort new shows into Family/Adult/Anime categories" -ForegroundColor Yellow
Write-Host "   • Use content ratings and genres to make decisions" -ForegroundColor Yellow
Write-Host "   • Move shows to the right 'Incomplete' folders" -ForegroundColor Yellow
Write-Host ""

try {
    # AutoSorter resolves its own config path based on the root folder.
    & $AutoSorterScriptPath
    Write-Host ""
    Write-Host "✅ AutoSorter completed successfully!" -ForegroundColor Green
    Write-Host "   Shows have been sorted into their categories." -ForegroundColor Green
}
catch {
    $overallSuccess = $false
    Write-Host ""
Write-Host "ERROR: CompleteARR SONARR AutoSorter failed with error:" -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "WARNING: Continuing to run CompleteARR Engine anyway." -ForegroundColor Yellow
}

# ------------------------------------------------------------
# Run Engine (pass -ConfigPath if available)
# ------------------------------------------------------------

Write-Host ""
Write-Host "------------------------------------------------" -ForegroundColor Yellow
Write-Host " 🔄 STEP 2: Running Series Engine" -ForegroundColor Yellow
Write-Host "------------------------------------------------" -ForegroundColor Yellow
Write-Host ""
Write-Host "📋 The Series Engine will:" -ForegroundColor Yellow
Write-Host "   • Check which shows are 'Complete' (all episodes downloaded)" -ForegroundColor Yellow
Write-Host "   • Move complete shows to 'Complete' folders (visible to users)" -ForegroundColor Yellow
Write-Host "   • Move incomplete shows back if new episodes are missing" -ForegroundColor Yellow
Write-Host "   • Manage special episode monitoring" -ForegroundColor Yellow
Write-Host ""

try {
    if (Test-Path -LiteralPath $ConfigPath) {
        & $EngineScriptPath -ConfigPath $ConfigPath
    }
    else {
        # Fallback: let the engine resolve its own default config.
        & $EngineScriptPath
    }

    Write-Host ""
    Write-Host "✅ Series Engine completed successfully!" -ForegroundColor Green
    Write-Host "   Shows have been promoted/demoted as needed." -ForegroundColor Green
}
catch {
    $overallSuccess = $false
    Write-Host ""
Write-Host "ERROR: CompleteARR SONARR Engine failed with error:" -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
}

# ------------------------------------------------------------
# Final message
# ------------------------------------------------------------

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host " 🎉 CompleteARR SONARR Finished!" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

if ($overallSuccess) {
    Write-Host "✅ All tasks completed successfully!" -ForegroundColor Green
    Write-Host "   Your shows are now organized and ready for users." -ForegroundColor Green
}
else {
    Write-Host "⚠️  Some tasks encountered errors." -ForegroundColor Yellow
    Write-Host "   Check the logs in the CompleteARR_Logs folder for details." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "📁 Log files are saved in: CompleteARR_Logs/" -ForegroundColor Cyan
Write-Host ""
Read-Host "Press ENTER to close this window:"
