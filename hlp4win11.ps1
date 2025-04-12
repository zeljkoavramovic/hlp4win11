#Step 1: Copy this script somewhere to your root disk like c:\hlp4win11\hlp4win11.ps1 because it might not work from network drive if manual download is needed.
#Step 2: Open PowerShell as administrator and navigate to c:\hlp4win11 directory to execute the following command:   powershell.exe -ExecutionPolicy Bypass -File .\hlp4win11.ps1
#Step 3: If everything went well, you will be able to open HLP files on your Windows 11 machine.
#Step 4: Enjoy life! If some major Windows update overwrites legacy files then repeat steps 1-3.

#Requires -Version 3.0
#Requires -RunAsAdministrator

<#
.SYNOPSIS
Downloads the correct WinHlp32 MSU update (KB917607) for the system architecture,
extracts the required components, and installs them onto Windows 10/11 using simplified
file matching and fallback to English MUI files if necessary.

.DESCRIPTION
This script automates the process of restoring WinHlp32 functionality on modern Windows.
1. Checks for Administrator privileges and required PowerShell version.
2. Detects the system architecture (x64 or x86).
3. Fetches the Microsoft Download Center details page for the corresponding KB917607 MSU.
4. Parses the HTML to find the direct download link using regex.
5. Downloads the correct MSU file using Invoke-WebRequest. If download fails, prompts for manual download.
6. Expands the downloaded MSU and the CAB file within it.
7. Extracts the necessary WinHlp32.exe, ftsrch.dll, ftlx*.dll files and their corresponding
   MUI files. It attempts to use the system's default UI language first using simple path matching.
   If MUI files for the system language are not found, it falls back to finding English (en-US) MUI files.
8. Takes ownership and sets permissions on the target system files before replacing them,
   creating sequential backups with a .bkp extension (e.g., .01.bkp, .02.bkp). Target paths
   match original hlp.ps1 (files placed directly under %SystemRoot%, using en-US subdir if fallback occurred).
9. Cleans up temporary files.

MUST BE RUN AS ADMINISTRATOR.

.NOTES
Author: Zeljko Avramovic
Date:   April 12th, 2025
Version: 1.1
Requires PowerShell 3.0 or later.
Requires Administrator privileges.
Execution Policy may need adjustment (e.g., 'Set-ExecutionPolicy RemoteSigned' or run as 'powershell.exe -ExecutionPolicy Bypass -File .\hlp4win11.ps1').
Relies on the download link being present in a specific format within the Microsoft Download Center page's static HTML.
Website structure changes at Microsoft can break the download part of this script. If this is the case you can download file manually and rerun the script.
If your system language MUI files are not in the package, the WinHelp UI should appear in American English.
#>

# --- Script Configuration ---
$ScriptDir          = $PSScriptRoot # Directory where the script resides or is running from
$TempExtractDirBase = "WinHlp32_Install_Temp" # Base name for temporary extraction folder
$TempMsuDir         = Join-Path -Path $ScriptDir -ChildPath "${TempExtractDirBase}_MSU"
$TempCabDir         = Join-Path -Path $TempMsuDir -ChildPath "${TempExtractDirBase}_CAB"
$BackupExtension    = "bkp" # Backup extension for original system files

# --- URLs for the specific KB download pages ---
$downloadInfo = @{
    "x64" = @{
        Url = "https://www.microsoft.com/en-us/download/details.aspx?id=47671"
        Description = "KB917607 x64 (Win 8.1)"
        ExpectedFileName = "Windows8.1-KB917607-x64.msu"
        CabPattern = "Windows8.1-KB917607-x64*.cab" # Pattern to find the CAB inside MSU
    }
    "x86" = @{
        Url = "https://www.microsoft.com/en-us/download/details.aspx?id=47667"
        Description = "KB917607 x86 (Win 8.1)"
        ExpectedFileName = "Windows8.1-KB917607-x86.msu"
        CabPattern = "Windows8.1-KB917607-x86*.cab" # Pattern to find the CAB inside MSU
    }
}

# --- Global Variables ---
$Global:DownloadedMsuPath = $null # Will store the path to the successfully downloaded MSU
$Global:FailureCount = 0 # Track failures across stages

# --- Error Preference ---
# Stop on terminating errors to ensure try/catch blocks work as expected
$ErrorActionPreference = 'Stop'

# --- Helper Function for Replacing System Files (with Sequential Backup) ---
Function Replace-SystemFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SourceFile,

        [Parameter(Mandatory=$true)]
        [string]$DestinationFile,

        [Parameter(Mandatory=$true)]
        [string]$BackupExt # Pass the backup extension
    )

    Write-Verbose "Function Replace-SystemFile called with Source: '$SourceFile', Destination: '$DestinationFile'"
    $currentUserPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $UserPrincipal = $currentUserPrincipal.Identity.Name # Gets the correct format (e.g., DOMAIN\User or COMPUTERNAME\User)

    $DestinationBackup = $null # Will hold the actual backup name used
    $foundBackupSlot = $false

    Write-Host "  Processing destination: '$DestinationFile'"

    if (-not (Test-Path -Path $SourceFile -PathType Leaf)) {
        Write-Error "    Source file not found: '$SourceFile'"
        return $false # Indicate failure
    }
    Write-Verbose "    Source file '$SourceFile' exists."

    if (-not (Test-Path -Path $DestinationFile -PathType Leaf)) {
        Write-Warning "    Destination file does not exist: '$DestinationFile'. Will copy directly without backup."
        # Proceed to copy directly if destination doesn't exist
    } else {
        Write-Verbose "    Destination file '$DestinationFile' exists. Proceeding with backup."
        # --- Sequential Backup Logic ---
        Write-Host "    Searching for available backup slot (.01.$BackupExt - .99.$BackupExt)..."
        for ($i = 1; $i -le 99; $i++) {
            # Format number with leading zero (e.g., 01, 02) and add the custom extension
            $potentialBackupName = "{0}.{1:D2}.{2}" -f $DestinationFile, $i, $BackupExt
            if (-not (Test-Path -Path $potentialBackupName -PathType Leaf)) {
                $DestinationBackup = $potentialBackupName
                $foundBackupSlot = $true
                Write-Host "    Found available backup slot: '$DestinationBackup'" -ForegroundColor Green
                break # Exit the loop once a free slot is found
            }
        }

        if (-not $foundBackupSlot) {
            Write-Error "    Could not find an available backup slot (all .01.$BackupExt through .99.$BackupExt exist for '$DestinationFile'). Cannot rename original file."
            return $false # Indicate failure
        }
        # --- End Sequential Backup Logic ---

        # 1. Take Ownership
        Write-Host "    Attempting to take ownership of '$DestinationFile'..."
        Write-Verbose "    Executing: takeown.exe /F `"$DestinationFile`" /A"
        & takeown.exe /F $DestinationFile /A # Use /A for Administrators group ownership - more robust
        if ($LASTEXITCODE -ne 0) {
             Write-Warning "    Takeown /A failed (Code: $LASTEXITCODE). Trying takeown for current user '$UserPrincipal'..."
             Write-Verbose "    Executing: takeown.exe /F `"$DestinationFile`""
             & takeown.exe /F $DestinationFile
             if ($LASTEXITCODE -ne 0) {
                Write-Error "    Failed to take ownership of '$DestinationFile' (Code: $LASTEXITCODE)."
                return $false
            }
        }
        Write-Verbose "    Ownership taken successfully."

        # 2. Grant Full Control to Administrators group (usually better than specific user)
        Write-Host "    Granting Full Control to 'Administrators' on '$DestinationFile'..."
        $AdministratorsPrincipal = "BUILTIN\Administrators" # Use well-known SID alias
        Write-Verbose "    Executing: icacls.exe `"$DestinationFile`" /grant `"$AdministratorsPrincipal`:(F)`" /C"
        & icacls.exe $DestinationFile /grant "$AdministratorsPrincipal`:(F)" /C
        if ($LASTEXITCODE -ne 0) {
            # Sometimes granting to the specific user helps if Administrators fails
            Write-Warning "    ICACLS grant for Administrators failed (Code: $LASTEXITCODE). Attempting for '$UserPrincipal'..."
             Write-Verbose "    Executing: icacls.exe `"$DestinationFile`" /grant `"$UserPrincipal`:(F)`" /C"
             & icacls.exe $DestinationFile /grant "$UserPrincipal`:(F)" /C
             if ($LASTEXITCODE -ne 0) {
                Write-Warning "    Failed to grant Full Control permissions on '$DestinationFile' (Code: $LASTEXITCODE). Rename might still work if ownership is sufficient."
                # Don't return false here, attempt the rename anyway
            } else {
                 Write-Verbose "    Granted Full Control to '$UserPrincipal' successfully."
            }
        } else {
             Write-Verbose "    Granted Full Control to Administrators successfully."
        }


        # 3. Rename (Backup) using the found sequential name
        Write-Host "    Renaming '$DestinationFile' to '$DestinationBackup'..."
        try {
            Rename-Item -Path $DestinationFile -NewName $DestinationBackup -Force -ErrorAction Stop
            Write-Verbose "    Rename successful."
        } catch {
            Write-Error "    Failed to rename '$DestinationFile' to '$DestinationBackup'. Error: $($_.Exception.Message)"
            # Attempt to revert permissions/ownership? Maybe too complex. Log error clearly.
            Write-Error "    Manual check may be needed for '$DestinationFile'."
            return $false
        }
    } # End of 'else' block for existing destination file

    # 4. Copy New File
    Write-Host "    Copying '$SourceFile' to '$DestinationFile'..."
    try {
        Copy-Item -Path $SourceFile -Destination $DestinationFile -Force -ErrorAction Stop
        Write-Host "    File successfully copied." -ForegroundColor Green
    } catch {
        Write-Error "    Failed to copy '$SourceFile' to '$DestinationFile'. Error: $($_.Exception.Message)"
        # Attempt to restore backup if copy fails after rename
        if ($DestinationBackup -and (Test-Path -Path $DestinationBackup -PathType Leaf)) {
            Write-Warning "    Attempting to restore backup file '$DestinationBackup'..."
            try {
                 # Rename backup back to the original name
                 $originalFileName = Split-Path -Path $DestinationFile -Leaf
                 Rename-Item -Path $DestinationBackup -NewName $originalFileName -Force -ErrorAction Stop
                 Write-Warning "    Backup restored to '$originalFileName'."
            } catch {
                 Write-Error "    CRITICAL: Failed to restore backup '$DestinationBackup' after copy failure. Error: $($_.Exception.Message). Manual intervention required for '$DestinationFile'."
            }
        }
        return $false
    }

    return $true # Indicate success
}


# --- Main Script ---
Write-Host "=========================================================="
Write-Host " Starting hlp4win11 installation script for Windows 11 " -ForegroundColor Yellow
Write-Host "=========================================================="

# 1. Check for Admin privileges
Write-Host "Step 1: Checking for Administrator privileges..."
$currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-Not $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script must be run as Administrator."
    Write-Error "Please right-click the script and select 'Run as administrator'."
    Read-Host "Press Enter to exit"
    Exit 1
}
Write-Host "  Administrator privileges confirmed." -ForegroundColor Green

# 2. Detect Architecture
Write-Host "`nStep 2: Detecting system architecture..."
$architecture = $env:PROCESSOR_ARCHITECTURE
if ($architecture -eq 'AMD64') {
    $archKey = 'x64'
    Write-Host "  Detected architecture: x64" -ForegroundColor Green
} elseif ($architecture -eq 'x86') {
    $archKey = 'x86'
    Write-Host "  Detected architecture: x86" -ForegroundColor Green
} else {
    Write-Error "Unsupported architecture detected: $architecture"
    Read-Host "Press Enter to exit"
    Exit 1
}
$targetDownload = $downloadInfo[$archKey]

# 3. Download Phase
Write-Host "`nStep 3: Downloading required MSU file ($($targetDownload.Description))..."
$pageUrl = $targetDownload.Url
$expectedFileName = $targetDownload.ExpectedFileName
$htmlContent = $null
$downloadUrl = $null
$destinationPath = Join-Path -Path $ScriptDir -ChildPath $expectedFileName

# Check if file already exists (from previous attempt or manual download)
if (Test-Path -Path $destinationPath -PathType Leaf) {
    Write-Host "  MSU file '$expectedFileName' already exists in script directory. Skipping download." -ForegroundColor Cyan
    $Global:DownloadedMsuPath = $destinationPath
} else {
    Write-Host "  MSU file not found locally. Attempting automatic download..."
    try {
        # Fetch Download Page HTML
        $headers = @{
            'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        Write-Host "  Fetching details page: $pageUrl"
        $response = Invoke-WebRequest -Uri $pageUrl -UseBasicParsing -Headers $headers -TimeoutSec 120
        $htmlContent = $response.Content
        Write-Host "  HTML content fetched successfully." -ForegroundColor Green

        # Extract Download URL using Regex
        $escapedFileName = [regex]::Escape($expectedFileName)
        # Template uses single quotes: escape literal single quotes with '', use " directly. {0} is placeholder.
        $regex_pattern_template = '<a\s+[^>]*?href\s*=\s*[''"](?<Url>[^''"]*download\.microsoft\.com/[^''"]*/{0}[^''"]*)[''"][^>]*>'
        $regex_directlink = $regex_pattern_template -f $escapedFileName
        Write-Verbose "  Using Regex: $regex_directlink"

        # Find the first matching line/object
        $firstMatchInfo = $htmlContent | Select-String -Pattern $regex_directlink | Select-Object -First 1

        if ($firstMatchInfo -and $firstMatchInfo.Matches[0].Groups['Url'].Success) {
            # Extract URL from the named group of the first match on that line
            $downloadUrl = $firstMatchInfo.Matches[0].Groups['Url'].Value -replace '&amp;', '&'
            Write-Host "    Found download link: $downloadUrl" -ForegroundColor Green
        } else {
            # Handle case where pattern didn't match or capture group failed
            Write-Error "  Could not find download link pattern, or URL capture failed. Microsoft page structure might have changed."
            throw "Download link pattern not found or URL capture failed"
        }

        # Attempt download using Invoke-WebRequest
        Write-Host "  Attempting download..."
        Write-Host "    Source: $downloadUrl"
        Write-Host "    Destination: $destinationPath"
        # PowerShell 5+ shows progress automatically. For PS 3/4, this will just wait.
        Invoke-WebRequest -Uri $downloadUrl -OutFile $destinationPath -UseBasicParsing -TimeoutSec 300 # 5 minute timeout
        Write-Host "  Download completed successfully via Invoke-WebRequest." -ForegroundColor Green

        # Assign path only AFTER successful download
        $Global:DownloadedMsuPath = $destinationPath

    } catch {
        Write-Warning "`nAutomatic download failed: $($_.Exception.Message)"
        Write-Host "`nCould not automatically download the required MSU file."
        Write-Host "This might be due to network issues, Microsoft changing the download page, or other errors."
        Write-Host "`nPlease perform the following steps:"
        Write-Host "1. Manually download the correct package for your system:" -ForegroundColor Yellow
        Write-Host "   Architecture: $archKey"
        Write-Host "   Expected File: $expectedFileName"
        Write-Host "   Download Link: $pageUrl" -ForegroundColor Cyan
        Write-Host "   If link above is missing or broken, search Microsoft for 'KB917607 download' for Windows 8.1 ($archKey)."
        Write-Host "2. Save the downloaded file AS '$expectedFileName' in the *same directory* as this script:" -ForegroundColor Yellow
        Write-Host "   $ScriptDir" -ForegroundColor Cyan
        Write-Host "3. Rerun this script." -ForegroundColor Yellow
        Read-Host "Press Enter to exit after noting the instructions"
        Exit 1 # Exit after providing manual instructions
    }
} # End else block for needing download

# Check if we have a valid MSU path before proceeding (either pre-existing or downloaded)
if (-not $Global:DownloadedMsuPath -or -not (Test-Path -Path $Global:DownloadedMsuPath -PathType Leaf)) {
     Write-Error "`nCould not obtain the required MSU file '$expectedFileName'. Cannot proceed with installation."
     Read-Host "Press Enter to exit"
     Exit 1
}

Write-Host "`nMSU file '$expectedFileName' is ready at: $Global:DownloadedMsuPath" -ForegroundColor Cyan


# 4. Installation Phase
Write-Host "`n=========================================================="
Write-Host " Starting WinHlp32 Component Installation Phase " -ForegroundColor Yellow
Write-Host "=========================================================="

$installationSuccess = $true # Assume success initially for this phase

# Use a try/catch/finally block to ensure cleanup happens for installation steps
try {
    # 4a. Clean up previous temp directories if they exist
    Write-Host "Step 4a: Cleaning up old temporary directories (if any)..."
    if (Test-Path -Path $TempMsuDir) {
        Write-Verbose "  Removing existing directory: $TempMsuDir"
        Remove-Item -Path $TempMsuDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    # No need to check TempCabDir separately, it's inside TempMsuDir

    # 4b. Expand MSU
    Write-Host "Step 4b: Expanding MSU file '$($Global:DownloadedMsuPath)'..."
    try {
        New-Item -Path $TempMsuDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        Write-Verbose "  Executing: expand.exe -F:* `"$($Global:DownloadedMsuPath)`" `"$TempMsuDir`""
        & expand.exe -F:* $Global:DownloadedMsuPath $TempMsuDir
        if ($LASTEXITCODE -ne 0) { throw "expand.exe failed for MSU with exit code $LASTEXITCODE" }
        Write-Host "  MSU expanded successfully to '$TempMsuDir'" -ForegroundColor Green
    } catch {
        Write-Error "  Failed to expand MSU file. Error: $($_.Exception.Message)"
        $installationSuccess = $false
        throw # Stop execution here, cleanup will happen in finally
    }

    # 4c. Find and Expand CAB
    Write-Host "Step 4c: Finding and expanding relevant CAB file..."
    $cabPattern = $targetDownload.CabPattern
    Write-Verbose "  Searching for CAB file matching pattern '$cabPattern' in '$TempMsuDir'"
    $cabFileItem = Get-ChildItem -Path $TempMsuDir -Filter $cabPattern | Select-Object -First 1
    if (-not $cabFileItem) {
        Write-Error "  CAB file matching '$cabPattern' not found within expanded MSU content at '$TempMsuDir'."
        $installationSuccess = $false
        throw "Required CAB file not found." # Stop execution
    }
    $cabFilePath = $cabFileItem.FullName
    Write-Host "  Found CAB file: $cabFilePath"
    try {
        New-Item -Path $TempCabDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        Write-Verbose "  Executing: expand.exe -F:* `"$cabFilePath`" `"$TempCabDir`""
        & expand.exe -F:* $cabFilePath $TempCabDir
        if ($LASTEXITCODE -ne 0) { throw "expand.exe failed for CAB with exit code $LASTEXITCODE" }
        Write-Host "  CAB expanded successfully to '$TempCabDir'" -ForegroundColor Green
    } catch {
        Write-Error "  Failed to expand CAB file '$cabFilePath'. Error: $($_.Exception.Message)"
        $installationSuccess = $false
        throw # Stop execution
    }

    # 4d. Get System Language
    Write-Host "Step 4d: Getting default system UI language..."
    $muiLang = $null
    try {
        $muiLang = (Get-Culture).Name # e.g., "en-US"
        Write-Host "  Detected system language (from Get-Culture): $muiLang" -ForegroundColor Green
    } catch {
        Write-Warning "  Get-Culture failed. Error: $($_.Exception.Message)."
        try {
             $muiLang = (Get-WinSystemLocale).Name # Try original method as fallback
             Write-Host "  Detected system language (from Get-WinSystemLocale): $muiLang" -ForegroundColor Green
        } catch {
            Write-Error "  Failed to get system language using multiple methods. Error: $($_.Exception.Message)"
            Write-Warning "  Will attempt to proceed using 'en-US' as a fallback language."
            $muiLang = "en-US" # Fallback language if detection completely fails
        }
    }
    $fallbackMuiLang = "en-US" # Define the standard fallback

    # 4e. Get Architecture from MSU filename (like original hlp.ps1)
    Write-Host "Step 4e: Determining architecture from MSU filename..."
    $arch = if ($Global:DownloadedMsuPath -match 'x64') { "amd64" } else { "x86" }
    Write-Host "  Detected architecture string for path matching: $arch" -ForegroundColor Green

    # 4f. Find Source Files (Reverted Logic + Fallback) & Define Target Paths (Reverted Paths)
    Write-Host "Step 4f: Identifying required files and target locations (using original script logic + fallback)..."
    $systemRoot = $env:SystemRoot
    $sourceFiles = @{} # Hashtable to store found source paths
    $targetPaths = @{} # Hashtable to store target paths

    # Define file list and target locations (matching original hlp.ps1)
    $filesToProcess = @(
        @{ Name="WinHlp32 Exe MUI"; SourceFilter="winhlp32.exe.mui"; TargetSubPath="$muiLang\winhlp32.exe.mui"; NeedsLang=$true  }
        @{ Name="WinHlp32 Exe";    SourceFilter="winhlp32.exe";     TargetSubPath="winhlp32.exe";           NeedsLang=$false }
        @{ Name="FTSearch DLL MUI";SourceFilter="ftsrch.dll.mui";   TargetSubPath="$muiLang\ftsrch.dll.mui"; NeedsLang=$true  }
        @{ Name="FTSearch DLL";    SourceFilter="ftsrch.dll";       TargetSubPath="ftsrch.dll";             NeedsLang=$false } # Target reverted
        @{ Name="FTLasso JP DLL";  SourceFilter="ftlx0411.dll";     TargetSubPath="ftlx0411.dll";           NeedsLang=$false } # Target reverted
        @{ Name="FTLasso TH DLL";  SourceFilter="ftlx041e.dll";     TargetSubPath="ftlx041e.dll";           NeedsLang=$false } # Target reverted
    )

    $allFilesFound = $true
    Write-Host "  Searching for files in '$TempCabDir'..."
    foreach ($fileInfo in $filesToProcess) {
        $key = $fileInfo.Name
        $filter = $fileInfo.SourceFilter
        $targetSubPath = $fileInfo.TargetSubPath # Get the initial target subpath
        $foundItem = $null
        $usedLang = $muiLang # Assume we'll use the primary language
        $isFallback = $false

        Write-Verbose "  Searching for '$key' (Filter: '$filter')..."
        $search = Get-ChildItem -Path $TempCabDir -Recurse -Filter $filter -File -ErrorAction SilentlyContinue

        # --- Attempt 1: Find using primary language ---
        if ($fileInfo.NeedsLang) {
            Write-Verbose "    Attempt 1: Applying Where-Object { `$_.FullName -match '$arch' -and `$_.FullName -match '$muiLang' }"
            $foundItem = $search | Where-Object { $_.FullName -match $arch -and $_.FullName -match $muiLang } | Select-Object -First 1
        } else {
            Write-Verbose "    Applying Where-Object { `$_.FullName -match '$arch' }"
            $foundItem = $search | Where-Object { $_.FullName -match $arch } | Select-Object -First 1
        }

        # --- Attempt 2: Fallback if needed ---
        if (-not $foundItem -and $fileInfo.NeedsLang) {
            Write-Warning "    Could not find MUI file for '$key' using specific language '$muiLang'."
            Write-Warning "    Attempting fallback search using language '$fallbackMuiLang'..."
            $isFallback = $true
            $usedLang = $fallbackMuiLang # Switch to fallback language

            Write-Verbose "    Attempt 2: Applying Where-Object { `$_.FullName -match '$arch' -and `$_.FullName -match '$usedLang' }"
            $foundItem = $search | Where-Object { $_.FullName -match $arch -and $_.FullName -match $usedLang } | Select-Object -First 1

            if ($foundItem) {
                 Write-Host "    Found '$key' using fallback language '$usedLang'." -ForegroundColor Cyan
                 # Adjust target subpath to use the fallback language
                 $targetSubPath = $targetSubPath.Replace($muiLang, $usedLang)
            } else {
                 Write-Warning "    Fallback search also failed to find '$key' using language '$usedLang'."
            }
        }

        # --- Final Check and Assignment ---
        if ($foundItem) {
            $sourceFiles[$key] = $foundItem.FullName
            $targetPaths[$key] = Join-Path -Path $systemRoot -ChildPath $targetSubPath # Build full target path
            Write-Host "    Found '$key': Source='$($foundItem.FullName)', Target='$($targetPaths[$key])'" -ForegroundColor Green
        } else {
            # If still not found after initial attempt (and fallback if applicable)
            Write-Error "    ERROR: Could not find source file for '$key' (Filter: $filter, Arch: $arch, PrimaryLang: $muiLang, FallbackLang: $fallbackMuiLang)"
            $allFilesFound = $false
        }
    } # End foreach

    if (-not $allFilesFound) {
        Write-Error "  One or more required source files were not found in the expanded CAB content, even after fallback attempts."
        Write-Error "  Please check the CAB structure in '$TempCabDir'."
        $installationSuccess = $false
        throw "Missing required source files."
    }
    Write-Host "  All required source files located successfully (using fallbacks where necessary)." -ForegroundColor Green


    # 4g. Prepare Target Directories and Replace System Files
    Write-Host "`nStep 4g: Replacing system files (requires Administrator privileges)..."
    $currentSuccessCount = 0
    $currentFailureCount = 0

    # Create target directories if they don't exist
    # Get unique parent directories from the final target paths
    $dirsToEnsure = $targetPaths.Values | Split-Path -Parent | Select-Object -Unique
    foreach ($dir in $dirsToEnsure) {
        if (-not (Test-Path -Path $dir -PathType Container)) {
            Write-Host "  Creating missing target directory: $dir"
            try {
                 New-Item -Path $dir -ItemType Directory -Force | Out-Null
            } catch {
                 Write-Error "    Failed to create directory '$dir'. Error: $($_.Exception.Message)"
                 Write-Error "    File installation into this directory might fail."
                 $currentFailureCount++ # Count this as a failure upfront
            }
        } else {
            Write-Verbose "  Target directory exists: $dir"
        }
    }


    # Perform replacements using the helper function
    foreach ($key in $sourceFiles.Keys) { # Iterate only through keys we actually found sources for
        Write-Host "`n--- Processing '$key' ---"
        $source = $sourceFiles[$key]
        $target = $targetPaths[$key] # Get the corresponding target path

        if (Replace-SystemFile -SourceFile $source -DestinationFile $target -BackupExt $BackupExtension) {
            $currentSuccessCount++
        } else {
            $currentFailureCount++
            Write-Warning "   Replacement failed for '$key'. See errors above."
        }
    }

     # Update global failure count
    $Global:FailureCount += $currentFailureCount
    Write-Host "`nFile replacement summary for this phase: $currentSuccessCount succeeded, $currentFailureCount failed."
    if ($currentFailureCount -gt 0) {
        $installationSuccess = $false # Mark phase as failed if any replacement failed
    }


} catch {
     Write-Error "`nAn error occurred during the installation phase: $($_.Exception.Message)"
     Write-Error "Script execution of installation steps aborted."
     $installationSuccess = $false # Ensure phase is marked as failed on exceptions
     # Flow will continue to finally block
} finally {
    # 5. Cleanup Temporary Files (Always runs after installation attempt)
    Write-Host "`nStep 5: Cleaning up temporary extraction directory..."
    if (Test-Path -Path $TempMsuDir) {
        Write-Verbose "  Removing temporary directory: $TempMsuDir"
        Remove-Item -Path $TempMsuDir -Recurse -Force -ErrorAction SilentlyContinue
        if (Test-Path -Path $TempMsuDir) {
             Write-Warning "  Could not fully remove temporary directory '$TempMsuDir'. Manual cleanup might be needed."
        } else {
             Write-Host "  Temporary directory removed successfully." -ForegroundColor Green
        }
    } else {
        Write-Host "  No temporary directory ('$TempMsuDir') found to remove (might indicate earlier failure)."
    }
}


Write-Host      "=========================================================="
# Check combined failures ($Global:FailureCount) and phase success ($installationSuccess)
if ($Global:FailureCount -gt 0 -or -not $installationSuccess) {
     Write-Warning " WinHlp32 installation process completed with $($Global:FailureCount) error(s)."
     Write-Warning " Please review the output above for details."
     Write-Host "=========================================================="
     # Optional exit with error code
     # exit 1
} else {
     Write-Host " WinHlp32 installation completed successfully! " -ForegroundColor Green
     Write-Host " You should now be able to open .hlp files."
     Write-Host "(Note: If your system language files were not found, WinHelp UI will be in English)" -ForegroundColor Yellow
     Write-Host "=========================================================="
}

# Optional pause at the end
# Read-Host "Press Enter to exit"