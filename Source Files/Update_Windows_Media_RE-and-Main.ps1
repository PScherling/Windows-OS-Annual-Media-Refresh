<#
.SYNOPSIS
    Annual offline refresh of Windows 10/11 Enterprise LTSC or Windows Server 2022/2025 base installation media.

.DESCRIPTION
    This script automates the annual refresh of Windows 10/11 Enterprise LTSC and Windows Server 2022/2025
    installation media by performing offline servicing of the base OS and Windows Recovery Environment
    (WinRE).

    The workflow is designed for enterprise environments and follows these principles:
      - Year-based isolation of base media (no overwriting of previous years)
      - Deterministic, repeatable offline servicing using DISM
      - Injection of the latest cumulative update (LCU) only
      - SafeOS Dynamic Update servicing for WinRE
      - Optional .NET cumulative updates
      - Full lifecycle handling of ISO mounting, WIM mounting, cleanup, and validation
      - Safe re-runs after failure via completion marker logic
      - Hard protection against overwriting completed yearly refreshes

    The script is intended to be executed once per year (typically December) and produces
    a refreshed, production-ready installation media set suitable for WDS or other
    deployment solutions.

.LINK
    https://learn.microsoft.com/en-us/windows/deployment/update/media-dynamic-update
	https://github.com/PScherling
	
.NOTES
          FileName: Update_Windows_Media_RE-and-Main.ps1
          Solution: Windows 11 OS Based Annual Base Media Refresh
          Author: Patrick Scherling
          Contact: @Patrick Scherling
          Primary: @Patrick Scherling
          Created: 2025-12-23
          Modified: 2025-12-30

          Version - 0.0.1 - (2025-12-23) - Finalized functional version 1 (enterprise-safe, rerunnable, year-isolated, WinRE-compliant).
          Version - 0.0.2 - (2025-12-24) - Restructuring the handleing of "base" media directory and workflow
          Version - 0.0.3 - (2025-12-29) - Errorhandling for mounted windows images after foregoing failure
          Version - 0.0.4 - (2025-12-30) - Prompt if user executes script not in december


.EXAMPLE
	.\Update_Windows_Media_RE-and-Main.ps1

    Runs the annual offline servicing workflow using the most recent ISO and LCU available in the configured directories. 
    Requires administrative privileges.


Requires -RunAsAdministrator
#>

Clear-Host

#------------------------------------------------------------
# Safety & utilities
#------------------------------------------------------------
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "This script must be run as Administrator."
}


$ErrorActionPreference = "Stop"
$ScriptStartTime = Get-Date

#------------------------------------------------------------
# Paths
#------------------------------------------------------------
$Version        = "0.0.4"
$Year           = (Get-Date).Year
$BASE_PATH      = "D:\mediaRefresh"
$BASE_YEAR_PATH = "$BASE_PATH\base\$Year"
$LogDir         = "$BASE_PATH\log"
$ISO_DIR        = "$BASE_PATH\iso"
$WORK           = "$BASE_PATH\temp"

$MAIN_MOUNT     = "$WORK\MainOS"
$WINRE_MOUNT    = "$WORK\WinRE"

$LCU_DIR        = "$BASE_PATH\packages\CU"
$SAFEOS_DIR     = "$BASE_PATH\packages\SafeOS"
$DOTNET_DIR     = "$BASE_PATH\packages\DotNet"
$SSU_DIR        = "$BASE_PATH\packages\SSU"

#------------------------------------------------------------
# LOGGING
#------------------------------------------------------------
#function Get-TS { "{0:HH:mm:ss}" -f (Get-Date) }
#function Log($msg) { Write-Output "$(Get-TS): $msg" }

if(-not (Test-Path $LogDir)){
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}
$script:LogFile = Join-Path $LogDir ("mediaRefresh_{0}.log" -f (Get-Date -Format "yyyy-MM-dd_HH-mm-ss"))

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","OK","WARN","ERROR")] 
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"
    $line | Out-File -FilePath $script:LogFile -Append

    switch ($Level) {
        "INFO"  { Write-Host $line }
		"OK"    { Write-Host $line -ForegroundColor Green }
        "WARN"  { Write-Host $line -ForegroundColor Yellow }
        "ERROR" { Write-Host $line -ForegroundColor Red }
    }
}

#------------------------------------------------------------
# CONSOLE HEADER
#------------------------------------------------------------
Write-Host -ForegroundColor Cyan "
    +----+ +----+     
    |####| |####|     
    |####| |####|       WW   WW II NN   NN DDDDD   OOOOO  WW   WW  SSSS
    +----+ +----+       WW   WW II NNN  NN DD  DD OO   OO WW   WW SS
    +----+ +----+       WW W WW II NN N NN DD  DD OO   OO WW W WW  SSS
    |####| |####|       WWWWWWW II NN  NNN DD  DD OO   OO WWWWWWW    SS
    |####| |####|       WW   WW II NN   NN DDDDD   OOOO0  WW   WW SSSS
    +----+ +----+       
"
Write-Host "-----------------------------------------------------------------------------------"
Write-Host "              Windows Offline Media Refresh | Version $Version"
Write-Host "-----------------------------------------------------------------------------------"


#------------------------------------------------------------
# Pre-flight DISM cleanup (critical)
#------------------------------------------------------------
Write-Log "Script Version: $Version"
Write-Log "Target refresh year: $Year"
if ((Get-Date).Month -ne 12) {
    Write-Log "This script is intended to run refreshes in December." "WARN"
    do {
        $continue = Read-Host -Prompt " Do you want to continue (y/n)"
        $continue = $continue.ToLower()
    } until ($continue -in @("y","n"))

    if ($continue -eq "n") {
        Write-Log "Execution aborted by user." "WARN"
        exit
    }
    elseif($continue -eq "y"){
        Write-Log "Execution continues due to user input." "WARN"
    }
}

Write-Log "Pre-flight DISM cleanup"
dism /Cleanup-Wim | Out-Null


#------------------------------------------------------------
# Select inputs (strict)
#------------------------------------------------------------
$ISO = Get-ChildItem $ISO_DIR -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $ISO) { 
    Write-Log "No ISO found in $ISO_DIR" "ERROR"
    throw "No ISO found in this directory" 
}

$LCU = Get-ChildItem $LCU_DIR -Filter "*.msu" | Sort-Object LastWriteTime -Descending #| Select-Object -First 1
if (-not $LCU) { 
    Write-Log "No LCU found in $LCU_DIR" "ERROR"
    throw "No LCU found in this directory" 
}

$SafeOS = Get-ChildItem $SAFEOS_DIR -File -Include "*.cab" -ErrorAction SilentlyContinue
$DotNetCUs = Get-ChildItem $DOTNET_DIR -File -Include "*.msu" -ErrorAction SilentlyContinue
$SSUs = Get-ChildItem $SSU_DIR -File -Include "*.msu", "*.cab" -ErrorAction SilentlyContinue

Write-Log "ISO     : $($ISO.Name)"
Write-Log "LCU     : $($LCU.Name)"
Write-Log "SafeOS  : $($SafeOS.Count) package(s)"
Write-Log "DotNet  : $($DotNetCUs.Count) package(s)"
Write-Log "SSU     : $($SSUs.Count) package(s)"


$OLD_MEDIA      = "$BASE_YEAR_PATH\$($ISO.Name)\oldMedia"
$NEW_MEDIA      = "$BASE_YEAR_PATH\$($ISO.Name)\newMedia"
$SUCCESS_MARKER = "$BASE_YEAR_PATH\$($ISO.Name)\.refresh_completed"
$OldStamp       = "$BASE_YEAR_PATH\$($ISO.Name)\_RefreshInfo.json"


#------------------------------------------------------------
# Overwrite / rerun safety
#------------------------------------------------------------
<#
if (Test-Path $NEW_MEDIA -and -not (Test-Path $SUCCESS_MARKER)) {
    Log "Cleaning incomplete newMedia from previous run"
    Remove-Item $NEW_MEDIA -Recurse -Force -ErrorAction SilentlyContinue
}
elseif (-not (Test-Path $NEW_MEDIA) -and Test-Path "$($SUCCESS_MARKER)") {
    throw "Refresh for year $Year already completed successfully. Refusing to overwrite."
}
elseif (-not (Test-Path $NEW_MEDIA) -and -not (Test-Path $SUCCESS_MARKER)) {
    Log "This is a fresh run"
}
#>

# ISO consistency check
if (Test-Path $OldStamp) {
    $OldInfo = Get-Content $OldStamp | ConvertFrom-Json
    $CurrentISOHash = (Get-FileHash $ISO.FullName -Algorithm SHA256).Hash

    if ($OldInfo.ISO_SHA256 -ne $CurrentISOHash) {
        Write-Log "ISO mismatch detected for $($ISO.Name) for year $Year. Refusing to mix base media." "ERROR"
        throw "ISO mismatch detected for this iso for this year. Refusing to mix base media."
    }
}

if (Test-Path $SUCCESS_MARKER) {
    Write-Log "Refresh for $($ISO.Name) for year $Year already completed successfully. Refusing to overwrite." "ERROR"
    throw "Refresh for this iso for this year already completed successfully. Refusing to overwrite."
}

if (Test-Path $NEW_MEDIA) {
    Write-Log "Incomplete or failed run detected for $($ISO.Name) for year $Year. Cleaning newMedia." "WARN"
    Remove-Item $NEW_MEDIA -Recurse -Force -ErrorAction SilentlyContinue
}
else {
    Write-Log "This is a fresh run" "OK"
}
Write-Host "-----------------------------------------------------------------------------------"
#------------------------------------------------------------
# Prepare folders
#------------------------------------------------------------
Write-Log "Preparing directories"
foreach ($dir in @(
    $BASE_YEAR_PATH,
    $OLD_MEDIA,
    $NEW_MEDIA,
    $WORK,
    $MAIN_MOUNT,
    $WINRE_MOUNT,
    $LCU_DIR,
    $SAFEOS_DIR,
    $DOTNET_DIR,
    $SSU_DIR
)) {
    if (-not (Test-Path $dir)) {
        Write-Log "Creating directory $dir"
        New-Item -ItemType Directory -Path $dir | Out-Null
    }
}

#------------------------------------------------------------
# ISO lifecycle
#------------------------------------------------------------
$IsoMount = $null
try {
    if (Test-Path "$OLD_MEDIA\sources\install.wim") {
        Write-Log "oldMedia already exists for year $Year. Reusing existing base media." "WARN"
    }
    else {
        Write-Log "Mounting ISO"
        $IsoMount = Mount-DiskImage -ImagePath $ISO.FullName -PassThru
        $IsoDrive = ($IsoMount | Get-Volume).DriveLetter + ":"

        Write-Log "Copying ISO to oldMedia"
        robocopy "$IsoDrive\" "$OLD_MEDIA" /E /R:3 /W:3 /NFL /NDL /NJH /NJS /NP | Out-Null
    }

    Write-Log "Copying oldMedia to newMedia"
    robocopy "$OLD_MEDIA\" "$NEW_MEDIA" /E /R:3 /W:3 /NFL /NDL /NJH /NJS /NP | Out-Null

    Write-Log "Clearing read-only attributes"
    Get-ChildItem $NEW_MEDIA -Recurse -File | Where-Object IsReadOnly | ForEach-Object { $_.IsReadOnly = $false }
}
finally {
    if ($IsoMount) {
        Write-Log "Dismounting ISO: $($ISO.Name)"
        Dismount-DiskImage -ImagePath "$($ISO.FullName)" -ErrorAction SilentlyContinue | Out-Null
    }
}
Write-Host "-----------------------------------------------------------------------------------"
#------------------------------------------------------------
# Process install.wim images
#------------------------------------------------------------
$InstallWim = "$NEW_MEDIA\sources\install.wim"

# Defensive unmount before mounting
Write-Log "Checking for existing mounted WIMs"

$mountedInfo = dism /Get-MountedWimInfo

if ($mountedInfo -match "Mount Dir") {
    Write-Log "Detected existing mounted WIM(s). Forcing cleanup." "WARN"

    # Attempt graceful cleanup first
    dism /Cleanup-Wim | Out-Null

    # Re-check
    $mountedInfo = dism /Get-MountedWimInfo
    if ($mountedInfo -match "Mount Dir") {
        Write-Log "Stale mount still detected. Manual intervention may be required." "ERROR"
        throw "Unable to clear mounted WIM state"
    }
}

$Images = Get-WindowsImage -ImagePath $InstallWim

Remove-Item "$WORK\install_refreshed.wim" -ErrorAction SilentlyContinue

# Ensure clean mount directory
Write-Log "Checking for mounted windows images"
$mnt = Get-WindowsImage -Mounted
if($mnt){
    Write-Log "Stale mounted images found. We need to remove them..." WARN
    foreach($mount in $mnt){
        if($mount.ImagePath -like "$($BASE_YEAR_PATH)\*") {
            Write-Log "Try dismounting image from: $($mount.Path)"
            Dismount-WindowsImage -Path $mount.Path -Discard -ErrorAction Stop
        }
    }
}

if (Test-Path $WINRE_MOUNT) {
    Write-Log "Removing stale mount directory $WINRE_MOUNT" "WARN"
    Remove-Item $WINRE_MOUNT -Recurse -Force -ErrorAction SilentlyContinue
}

foreach ($Image in $Images) {
     # Per-image mount directory (critical)
    $MAIN_MOUNT = Join-Path $WORK "MainOS_Index_$($Image.ImageIndex)"

    # Ensure clean mount directory
    if (Test-Path $MAIN_MOUNT) {
        Write-Log "Removing stale mount directory $MAIN_MOUNT" "WARN"
        Remove-Item $MAIN_MOUNT -Recurse -Force -ErrorAction SilentlyContinue
    }

    $MainMounted = $false
    Write-Log "============================================="
    try {
        # Mounting OS
        New-Item -ItemType Directory -Path $MAIN_MOUNT | Out-Null
        Write-Log "Mounting Main OS: $($Image.ImageName) (Image Index $($Image.ImageIndex))"
        Mount-WindowsImage -ImagePath $InstallWim -Index $Image.ImageIndex -Path $MAIN_MOUNT -ErrorAction Stop
        $MainMounted = $true

        #----------------------------------------------------
        # WinRE (only once, first image)
        #----------------------------------------------------
        if ($Image.ImageIndex -eq 1 -and $SafeOS) {

            $WinREMounted = $false
            $WinREWim = "$WORK\winre.wim"

            try {
                Write-Log "Servicing WinRE"
                Copy-Item "$MAIN_MOUNT\Windows\System32\Recovery\winre.wim" $WinREWim -Force

                Mount-WindowsImage -ImagePath $WinREWim -Index 1 -Path $WINRE_MOUNT
                $WinREMounted = $true

                foreach ($pkg in $SafeOS) {
                    Write-Log "  Adding SafeOS: $($pkg.Name)"
                    Add-WindowsPackage -Path $WINRE_MOUNT -PackagePath $pkg.FullName #| Out-Null
                }

                Write-Log "Starting DISM image cleanup"
                DISM /image:$WINRE_MOUNT /cleanup-image /StartComponentCleanup #| Out-Null
            }
            finally {
                if ($WinREMounted) {
                    Write-Log "Dismounting WinRE"
                    Dismount-WindowsImage -Path $WINRE_MOUNT -Save -ErrorAction Stop
                }
            }

            Copy-Item $WinREWim "$MAIN_MOUNT\Windows\System32\Recovery\winre.wim" -Force
        }

        #----------------------------------------------------
        # Main OS servicing
        #----------------------------------------------------
        foreach ($lcufile in $LCU){
            Write-Log "Adding LCU: $($lcufile.Name)"
            Add-WindowsPackage -Path $MAIN_MOUNT -PackagePath $lcufile.FullName #| Out-Null
        }

        if ($DotNetCUs) {
            foreach ($pkg in $DotNetCUs) {
                Write-Log "Adding .NET CU: $($pkg.Name)"
                Add-WindowsPackage -Path $MAIN_MOUNT -PackagePath $pkg.FullName #| Out-Null
            }
        }
        Write-Log "Starting DISM image cleanup"
        DISM /image:$MAIN_MOUNT /cleanup-image /StartComponentCleanup #| Out-Null
    }
    finally {
        if ($MainMounted) {
            Write-Log "Dismounting Main OS: $($Image.ImageName) (Index $($Image.ImageIndex))"
            Dismount-WindowsImage -Path $MAIN_MOUNT -Save -ErrorAction Stop #| Out-Null
        }
    }

    Write-Log "Exporting refreshed image $($Image.ImageName) (Index $($Image.ImageIndex))"
    Export-WindowsImage -SourceImagePath $InstallWim -SourceIndex $Image.ImageIndex -DestinationImagePath "$WORK\install_refreshed.wim" -CheckIntegrity #| Out-Null

    Write-Log "Cleanup of mount directoty"
    Remove-Item $MAIN_MOUNT -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "-----------------------------------------------------------------------------------"
#------------------------------------------------------------
# Replace install.wim
#------------------------------------------------------------
Write-Log "Replacing install.wim"
Move-Item "$WORK\install_refreshed.wim" $InstallWim -Force
#------------------------------------------------------------
# Final cleanup
#------------------------------------------------------------
Write-Log "Final DISM cleanup"
dism /Cleanup-Wim | Out-Null

if (Test-Path $WORK) {
    Write-Log "Removing temp folder"
    Remove-Item $WORK -Recurse -Force -ErrorAction SilentlyContinue
}

$FinalImages = Get-WindowsImage -ImagePath "$NEW_MEDIA\sources\install.wim"
if ($FinalImages.Count -ne $Images.Count) {
    Write-Log "Final install.wim image count mismatch. Aborting completion marker." "ERROR"
    throw "Final install.wim image count mismatch. Aborting completion marker."
}
else{
    

    $ScriptEndTime = Get-Date
    $Duration = $ScriptEndTime - $ScriptStartTime
    Write-Log ("Total runtime: {0:hh\:mm\:ss}" -f $Duration)

    $Stamp = @{
        Year        = $Year
        ISO         = $ISO.Name
        ISO_SHA256  = (Get-FileHash $ISO.FullName -Algorithm SHA256).Hash
        LCU         = $LCU.Name
        Completed   = (Get-Date)
        Runtime = ("{0:hh\:mm\:ss}" -f $Duration)
    }
    
    try{
        $Stamp | ConvertTo-Json -Depth 3 | Set-Content "$BASE_YEAR_PATH\$($ISO.Name)\RefreshInfo.json" -Encoding UTF8
    }
    catch{
        Write-Log "FATAL-ERROR: RefreshInfo file could not be created." "ERROR"
        throw "RefreshInfo file could not be created."
    }

    try{
        Write-Log "Marking refresh as completed"
        New-Item -ItemType File -Path $SUCCESS_MARKER -Force | Out-Null
    }
    catch{
        Write-Log "FATAL-ERROR: Success marker file could not be created." "ERROR"
        throw "Success marker file could not be created."
    }

    Write-Log "Media refresh completed successfully" "OK"

}

