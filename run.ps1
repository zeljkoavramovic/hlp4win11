# This script should be called remotely using:
# https://github.com/zeljkoavramovic/hlp4win11/blob/main/run.ps1
#
# This command fetches the run.ps1 bootstrap script from the internet and executes it directly in memory.
# The bootstrap script then downloads the main hlp4win11.ps1 installer script into the user's %TEMP% directory
# and launches it from there to perform the installation.

$scriptUrl = "https://raw.githubusercontent.com/zeljkoavramovic/hlp4win11/main/hlp4win11.ps1"

# Use the system TEMP directory with a subfolder inside for better organization
$localDir = Join-Path $env:TEMP "hlp4win11"
New-Item -Path $localDir -ItemType Directory -Force

$localPath = Join-Path $localDir "hlp4win11.ps1"

Invoke-WebRequest -Uri $scriptUrl -OutFile $localPath

powershell.exe -ExecutionPolicy Bypass -File $localPath