# This script should be called remotely using:
# https://github.com/zeljkoavramovic/hlp4win11/blob/main/run.ps1
#
# This command fetches the run.ps1 bootstrap script from the internet and executes it directly in memory.
# The bootstrap script then downloads the main hlp4win11.ps1 installer script into the user's %TEMP% directory
# and launches it from there to perform the installation.

# --- Bootstrao Script ---
Write-Host "=========================================================="
Write-Host " Downloading hlp4win11 installation script from github... " -ForegroundColor Yellow
Write-Host "=========================================================="

Write-Host "Step 0: Downloading script hlp4win11.ps1 from github..."

$scriptUrl = "https://raw.githubusercontent.com/zeljkoavramovic/hlp4win11/main/hlp4win11.ps1"
# Use the system TEMP directory with a subfolder inside for better organization
$localDir = Join-Path $env:TEMP "hlp4win11"
New-Item -Path $localDir -ItemType Directory -Force
$localPath = Join-Path $localDir "hlp4win11.ps1"
Invoke-WebRequest -Uri $scriptUrl -OutFile $localPath

Write-Host "  Script successfully downloaded to $localDir - now trying to execute it...`n" -ForegroundColor Green

powershell.exe -ExecutionPolicy Bypass -File $localPath