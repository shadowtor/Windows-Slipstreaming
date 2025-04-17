<#
.SYNOPSIS
    Slipstreams all drivers from a single folder and cumulative update packages into install.wim,
    and integrates the same drivers into boot.wim for WinPE connectivity.

.DESCRIPTION
    • Updates the Windows install.wim image:
        - Integrates every driver under $AdditionalDriversPath.
        - Slipstreams **all** MSU packages found in Downloads\InstallUpdates (including CU and .NET updates).
        - Optionally verifies integrations before committing.
    • Updates the Windows boot.wim image:
        - Integrates the same drivers for network support during Setup/WinPE.
    • Automatically downloads the update if a URL is given, unblocks it, and includes any additional MSUs in the folder.

.PARAMETER InstallWimPath
    Path to the install.wim file (e.g., ".\Downloads\install.wim").

.PARAMETER EditionIndex
    Index of the Windows edition in install.wim to update (default: 6).

.PARAMETER CumulativeUpdatePath
    Local path to a specific MSU file to include in addition to others in Downloads\InstallUpdates.

.PARAMETER CumulativeUpdateUrl
    URL to download a single MSU file into Downloads\InstallUpdates.

.PARAMETER ProcessAllIndexes
    Switch: process all indexes (1–11) in install.wim instead of a single index.

.PARAMETER VerifyIntegration
    Switch: after integrating drivers/update, list packages and pause for manual verification.

.PARAMETER BootWimPath
    Path to the boot.wim file (default: ".\Downloads\boot.wim").

.PARAMETER AdditionalDriversPath
    Folder containing all driver files/subfolders to import (default: ".\drivers").

.EXAMPLE
    .\SlipstreamUpdate.ps1 -InstallWimPath "C:\ISO\sources\install.wim" `
        -CumulativeUpdateUrl "https://example.com/windows11-kbXYZ.msu" `
        -BootWimPath "C:\ISO\sources\boot.wim" `
        -AdditionalDriversPath "C:\Drivers" -VerifyIntegration

.NOTES
    • Run in an elevated PowerShell session.
    • Backup original WIMs before making changes.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]$InstallWimPath,

    [Parameter(Mandatory=$false)]
    [int]$EditionIndex = 6,

    [Parameter(Mandatory=$false)]
    [string]$CumulativeUpdatePath,

    [Parameter(Mandatory=$false)]
    [string]$CumulativeUpdateUrl,

    [Parameter(Mandatory=$false)]
    [switch]$ProcessAllIndexes,

    [Parameter(Mandatory=$false)]
    [switch]$VerifyIntegration,

    [Parameter(Mandatory=$false)]
    [string]$BootWimPath = ".\Downloads\boot.wim",

    [Parameter(Mandatory=$false)]
    [string]$AdditionalDriversPath = ".\drivers"
)

# Always show verbose messages
$VerbosePreference = 'Continue'

#===========================================================================
# SECTION 1: Gather cumulative update packages
#===========================================================================

if (-not $CumulativeUpdatePath -and -not $CumulativeUpdateUrl) {
    Write-Error "Provide either -CumulativeUpdatePath or -CumulativeUpdateUrl (optional) to seed Downloads\InstallUpdates."
    exit 1
}

$updateFolder = Join-Path (Get-Location) "Downloads\InstallUpdates"
if (-not (Test-Path $updateFolder)) {
    New-Item -ItemType Directory -Force -Path $updateFolder | Out-Null
}

if ($CumulativeUpdateUrl) {
    try {
        $uri      = [Uri]$CumulativeUpdateUrl
        $fileName = Split-Path $uri.AbsolutePath -Leaf
        $dest     = Join-Path $updateFolder $fileName
        Write-Verbose "$(Get-Date -Format T) → Downloading update from '$CumulativeUpdateUrl'..."
        Invoke-WebRequest -Uri $CumulativeUpdateUrl -OutFile $dest -UseBasicParsing
        Write-Verbose "→ Downloaded to '$dest'."
    }
    catch {
        Write-Error "Download failed: $_"
        exit 1
    }
}

if ($CumulativeUpdatePath) {
    if (-not (Test-Path $CumulativeUpdatePath)) {
        Write-Error "Specified update path not found: '$CumulativeUpdatePath'"
        exit 1
    }
    Copy-Item -Path $CumulativeUpdatePath -Destination $updateFolder -Force
}

# Unblock all MSUs
Get-ChildItem -Path $updateFolder -Filter *.msu | ForEach-Object {
    try {
        Unblock-File -Path $_.FullName -ErrorAction Stop
        Write-Verbose "→ Unblocked '$($_.Name)'"
    }
    catch {
        Write-Verbose "Could not unblock '$($_.Name)': $_"
    }
}

$UpdatePackages = Get-ChildItem -Path $updateFolder -Filter *.msu | Select-Object -ExpandProperty FullName
if (-not $UpdatePackages) {
    Write-Error "No MSU packages found in '$updateFolder'."
    exit 1
}

#===========================================================================
# SECTION 2: Function to update a single install.wim index
#===========================================================================

function Update-Index {
    param([int]$IndexNumber)

    $mountDir = Join-Path (Get-Location) "InstallTemp_$IndexNumber"
    if (Test-Path $mountDir) { Remove-Item $mountDir -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $mountDir | Out-Null

    Write-Verbose "$(Get-Date -Format T) → Mounting install.wim index $IndexNumber..."
    Dism /Mount-WIM /WimFile:"$InstallWimPath" /Index:$IndexNumber /MountDir:"$mountDir"
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to mount install.wim index $IndexNumber."
        return $false
    }

    Write-Verbose "$(Get-Date -Format T) → Adding drivers from '$AdditionalDriversPath'..."
    Dism /Image:"$mountDir" /Add-Driver /Driver:"$AdditionalDriversPath" /Recurse
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Driver import failed; retrying with /ForceUnsigned..."
        Dism /Image:"$mountDir" /Add-Driver /Driver:"$AdditionalDriversPath" /Recurse /ForceUnsigned
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Driver import still failed for install.wim index $IndexNumber."
            Dism /Unmount-WIM /MountDir:"$mountDir" /Discard
            return $false
        }
    }

    foreach ($pkg in $UpdatePackages) {
        Write-Verbose "$(Get-Date -Format T) → Applying update '$([IO.Path]::GetFileName($pkg))'..."
        Dism /Image:"$mountDir" /Add-Package /PackagePath:"$pkg"
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Failed to apply update '$pkg' to index $IndexNumber."
            Dism /Unmount-WIM /MountDir:"$mountDir" /Discard
            return $false
        }
    }

    if ($VerifyIntegration) {
        Write-Verbose "$(Get-Date -Format T) → Verifying integrated packages..."
        Dism /Image:"$mountDir" /Get-Packages | Write-Output
        Read-Host "Press Enter to commit changes (or Ctrl+C to abort)"
    }

    Write-Verbose "$(Get-Date -Format T) → Committing install.wim index $IndexNumber..."
    Dism /Unmount-WIM /MountDir:"$mountDir" /Commit
    Remove-Item -Path $mountDir -Recurse -Force

    Write-Verbose "$(Get-Date -Format T) → install.wim index $IndexNumber completed."
    return $true
}

#===========================================================================
# SECTION 3: Function to update boot.wim
#===========================================================================

function Update-BootIndex {
    param([int]$IndexNumber)

    $bootMount = Join-Path (Get-Location) "BootMount_$IndexNumber"
    if (Test-Path $bootMount) { Remove-Item $bootMount -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $bootMount | Out-Null

    Write-Verbose "$(Get-Date -Format T) → Mounting boot.wim index $IndexNumber..."
    Dism /Mount-WIM /WimFile:"$BootWimPath" /Index:$IndexNumber /MountDir:"$bootMount"
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to mount boot.wim index $IndexNumber."
        return $false
    }

    Write-Verbose "$(Get-Date -Format T) → Adding drivers from '$AdditionalDriversPath' (boot)..."
    Dism /Image:"$bootMount" /Add-Driver /Driver:"$AdditionalDriversPath" /Recurse
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Boot driver import failed; retrying with /ForceUnsigned..."
        Dism /Image:"$bootMount" /Add-Driver /Driver:"$AdditionalDriversPath" /Recurse /ForceUnsigned
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Failed to add drivers to boot.wim index $IndexNumber."
            Dism /Unmount-WIM /MountDir:"$bootMount" /Discard
            return $false
        }
    }

    Write-Verbose "$(Get-Date -Format T) → Committing boot.wim index $IndexNumber..."
    Dism /Unmount-WIM /MountDir:"$bootMount" /Commit
    Remove-Item -Path $bootMount -Recurse -Force

    Write-Verbose "$(Get-Date -Format T) → boot.wim index $IndexNumber completed."
    return $true
}

#===========================================================================
# SECTION 4: Main processing with progress bars
#===========================================================================

# install.wim
if ($ProcessAllIndexes) {
    $total = 11
    for ($i = 1; $i -le $total; $i++) {
        $pct = [math]::Round(($i / $total) * 100)
        Write-Progress -Activity "Updating install.wim" -Status "Index $i of $total" -PercentComplete $pct
        if (-not (Update-Index -IndexNumber $i)) { exit 1 }
    }
} else {
    Write-Progress -Activity "Updating install.wim" -Status "Index $EditionIndex of 1" -PercentComplete 100
    if (-not (Update-Index -IndexNumber $EditionIndex)) { exit 1 }
}
Write-Progress -Activity "install.wim update" -Completed

# boot.wim (usually two indexes)
for ($b = 1; $b -le 2; $b++) {
    $pct = [math]::Round(($b / 2) * 100)
    Write-Progress -Activity "Updating boot.wim" -Status "Index $b of 2" -PercentComplete $pct
    if (-not (Update-BootIndex -IndexNumber $b)) { exit 1 }
}
Write-Progress -Activity "boot.wim update" -Completed

Write-Output "All slipstream operations completed successfully."
