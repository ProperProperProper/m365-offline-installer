#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Downloads the full offline Microsoft 365 Apps for Enterprise installer and installs it.

.PARAMETER WorkDir
    Root folder for ODT and Office source files. Default: C:\M365Offline

.PARAMETER Channel
    Update channel: Current | MonthlyEnterprise | SemiAnnual. Default: Current

.PARAMETER Language
    Language ID. Default: en-us

.PARAMETER Architecture
    64 or 32. Default: 64

.PARAMETER ExcludeApps
    Comma-separated app IDs to exclude e.g. "Access,Publisher,Groove"

.PARAMETER DownloadOnly
    Download source files only, do not install.

.PARAMETER InstallOnly
    Install from existing WorkDir, skip download.

.PARAMETER Unattended
    Skip all confirmation prompts (for automation).
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $WorkDir      = "$env:USERPROFILE\M365Offline",
    [ValidateSet("Current","MonthlyEnterprise","SemiAnnual")]
    [string] $Channel      = "Current",
    [string] $Language     = "en-us",
    [ValidateSet("64","32")]
    [string] $Architecture = "64",
    [string] $ExcludeApps = "",
    [switch] $DownloadOnly,
    [switch] $InstallOnly,
    [switch] $Unattended
)

$ErrorActionPreference = "Stop"
$ProgressPreference    = "SilentlyContinue"

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
$OdtDir     = Join-Path $WorkDir "ODT"
$OfficeDir  = Join-Path $WorkDir "OfficeSource"
$OdtExe     = Join-Path $OdtDir  "officedeploymenttool.exe"
$SetupExe   = Join-Path $OdtDir  "setup.exe"
$ConfigDown = Join-Path $OdtDir  "config_download.xml"
$ConfigInst = Join-Path $OdtDir  "config_install.xml"
$LogDir     = $WorkDir

# Minimum expected size of a complete Office source download (bytes)
$MinSourceBytes = 2GB

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Write-Step    { param($Msg) Write-Host "`n==> $Msg" -ForegroundColor Cyan }
function Write-Ok      { param($Msg) Write-Host "    [OK]   $Msg" -ForegroundColor Green }
function Write-Warn    { param($Msg) Write-Host "    [WARN] $Msg" -ForegroundColor Yellow }
function Write-Fail    { param($Msg) Write-Host "    [FAIL] $Msg" -ForegroundColor Red }
function Write-Info    { param($Msg) Write-Host "    [INFO] $Msg" -ForegroundColor Gray }

function Format-Bytes {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N1} MB" -f ($Bytes / 1MB) }
    return "{0} bytes" -f $Bytes
}

function Assert-ExitCode {
    param($Code, $Desc)
    # Treat null as 0 — can occur when process handle doesn't expose code; size check is the real gate
    if ($null -ne $Code -and $Code -ne 0) { throw "$Desc failed with exit code $Code" }
}

function Start-TrackedProcess {
    param([string]$FilePath, [string]$Arguments)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName        = $FilePath
    $psi.Arguments       = $Arguments
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true
    return [System.Diagnostics.Process]::Start($psi)
}

function Confirm-Step {
    param($Prompt)
    if ($Unattended) { return }
    Write-Host ""
    $ans = Read-Host "$Prompt [Y/N]"
    if ($ans -notmatch "^[Yy]") { throw "Aborted by user." }
}

function Get-FolderSize {
    param($Path)
    if (-not (Test-Path $Path)) { return 0 }
    (Get-ChildItem $Path -Recurse -ErrorAction SilentlyContinue |
     Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
}

function Build-ExcludeXml {
    if (-not $ExcludeApps) { return "" }
    ($ExcludeApps -split ",") | ForEach-Object {
        "      <ExcludeApp ID=`"$($_.Trim())`" />`n"
    }
}

function Write-XmlFile {
    param($Path, $Content)
    [System.IO.File]::WriteAllText($Path, $Content, [System.Text.Encoding]::UTF8)
}

# ---------------------------------------------------------------------------
# Banner
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "==========================================================" -ForegroundColor White
Write-Host "  Microsoft 365 Apps for Enterprise - Offline Installer"   -ForegroundColor White
Write-Host "==========================================================" -ForegroundColor White
Write-Info "WorkDir     : $WorkDir"
Write-Info "Channel     : $Channel"
Write-Info "Architecture: ${Architecture}-bit"
Write-Info "Language    : $Language"
if ($ExcludeApps) { Write-Info "Excluding   : $ExcludeApps" }
if ($DownloadOnly) { Write-Info "Mode        : DOWNLOAD ONLY" }
elseif ($InstallOnly) { Write-Info "Mode        : INSTALL ONLY" }
else { Write-Info "Mode        : DOWNLOAD + INSTALL" }

# ---------------------------------------------------------------------------
# PRE-FLIGHT CHECKS
# ---------------------------------------------------------------------------
Write-Step "Pre-flight checks..."

# Check 1: Not already installed
$officeReg = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" `
    -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -match "Microsoft 365|Office 365|Microsoft Office" }

if ($officeReg) {
    Write-Warn "Existing Office installation detected:"
    $officeReg | ForEach-Object { Write-Warn "  $($_.DisplayName) $($_.DisplayVersion)" }
    Confirm-Step "Office is already installed. Continue anyway?"
} else {
    Write-Ok "No existing Office installation found."
}

# Check 2: Disk space
$drive     = Split-Path $WorkDir -Qualifier
$disk      = Get-PSDrive ($drive.TrimEnd(":")) -ErrorAction SilentlyContinue
$freeBytes = if ($disk) { $disk.Free } else { 0 }
$needed    = if (-not $InstallOnly) { 10GB } else { 4GB }

Write-Info "Free space on ${drive}: $(Format-Bytes $freeBytes)  (need ~$(Format-Bytes $needed))"
if ($freeBytes -lt $needed) {
    throw "Insufficient disk space. Need $(Format-Bytes $needed), have $(Format-Bytes $freeBytes) free on $drive."
}
Write-Ok "Disk space OK."

# Check 3: Internet reachability (skip for install-only)
if (-not $InstallOnly) {
    Write-Info "Testing connectivity to Microsoft..."
    $ping = Test-Connection "download.microsoft.com" -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $ping) {
        Write-Warn "Could not ping download.microsoft.com."
        Confirm-Step "Internet connectivity check failed. Continue anyway?"
    } else {
        Write-Ok "Microsoft download servers reachable."
    }
}

# Check 4: Existing source files (install-only)
if ($InstallOnly) {
    $existingSize = Get-FolderSize $OfficeDir
    if ($existingSize -lt $MinSourceBytes) {
        throw "OfficeSource at '$OfficeDir' is $(Format-Bytes $existingSize) - too small. Run the download phase first."
    }
    Write-Ok "Office source files found ($(Format-Bytes $existingSize))."
}

Write-Ok "All pre-flight checks passed."

# ---------------------------------------------------------------------------
# Confirm before starting
# ---------------------------------------------------------------------------
Confirm-Step "Ready to begin. Proceed?"

# ---------------------------------------------------------------------------
# 1. Create directories
# ---------------------------------------------------------------------------
Write-Step "Creating working directories..."
New-Item -ItemType Directory -Force -Path $OdtDir, $OfficeDir | Out-Null
Write-Ok "ODT dir    : $OdtDir"
Write-Ok "Source dir : $OfficeDir"

# ---------------------------------------------------------------------------
# 2. Download and extract ODT
# ---------------------------------------------------------------------------
if (-not $InstallOnly) {
    Write-Step "Downloading Office Deployment Tool..."

    $odtPageUrl = "https://www.microsoft.com/en-us/download/details.aspx?id=49117"
    Write-Info "Fetching download URL from Microsoft..."

    try {
        $page   = Invoke-WebRequest -Uri $odtPageUrl -UseBasicParsing
        $odtUrl = ($page.Links |
                   Where-Object { $_.href -match "download\.microsoft\.com.*officedeploymenttool.*\.exe" } |
                   Select-Object -First 1).href
    } catch {
        throw "Could not reach Microsoft download page: $_"
    }

    if (-not $odtUrl) {
        throw "Could not parse ODT download URL from $odtPageUrl - the page layout may have changed."
    }

    Write-Ok "URL found: $odtUrl"
    Write-Info "Downloading ODT executable..."
    Invoke-WebRequest -Uri $odtUrl -OutFile $OdtExe -UseBasicParsing

    if (-not (Test-Path $OdtExe)) {
        throw "ODT file not found at $OdtExe after download."
    }
    $odtSize = (Get-Item $OdtExe).Length
    if ($odtSize -lt 1MB) {
        throw "ODT file is only $(Format-Bytes $odtSize) - download may have failed."
    }
    Write-Ok "Downloaded ODT: $(Format-Bytes $odtSize)"

    # Extract ODT
    Write-Info "Extracting ODT..."
    $proc = Start-TrackedProcess -FilePath $OdtExe -Arguments "/quiet /extract:`"$OdtDir`""
    $proc.WaitForExit()
    Assert-ExitCode $proc.ExitCode "ODT extraction"

    if (-not (Test-Path $SetupExe)) {
        throw "setup.exe not found in $OdtDir after extraction."
    }
    $setupSize = (Get-Item $SetupExe).Length
    Write-Ok "setup.exe extracted: $(Format-Bytes $setupSize)"
}

# Verify setup.exe present
if (-not (Test-Path $SetupExe)) {
    throw "setup.exe not found at $SetupExe. Run without -InstallOnly first."
}

# ---------------------------------------------------------------------------
# 3. Download Office source files
# ---------------------------------------------------------------------------
if (-not $InstallOnly) {
    Write-Step "Writing download config..."

    $excludeXml = Build-ExcludeXml
    $dlXml  = "<Configuration>`n"
    $dlXml += "  <Add OfficeClientEdition=`"$Architecture`"`n"
    $dlXml += "       Channel=`"$Channel`"`n"
    $dlXml += "       SourcePath=`"$OfficeDir`"`n"
    $dlXml += "       AllowCdnFallback=`"FALSE`">`n"
    $dlXml += "    <Product ID=`"O365ProPlusRetail`">`n"
    $dlXml += "      <Language ID=`"$Language`" />`n"
    $dlXml += $excludeXml
    $dlXml += "    </Product>`n"
    $dlXml += "  </Add>`n"
    $dlXml += "  <Logging Level=`"Standard`" Path=`"$LogDir`" />`n"
    $dlXml += "</Configuration>"
    Write-XmlFile $ConfigDown $dlXml
    Write-Ok "Download config written."

    Confirm-Step "About to download ~3-5 GB of Office source files to '$OfficeDir'. Continue?"

    Write-Step "Downloading Office source files (this will take a while)..."
    Write-Info "Channel: $Channel | Product: Microsoft 365 Apps for Enterprise | ${Architecture}-bit | $Language"
    Write-Info "Progress shown every 15 seconds. Window is NOT frozen - ODT is downloading in the background."
    Write-Host ""

    $maxAttempts = 3
    $dlSuccess   = $false

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        if ($attempt -gt 1) {
            Write-Warn "Retry $attempt of $maxAttempts..."
            Start-Sleep -Seconds 5
        }

        $dlProc = Start-TrackedProcess -FilePath $SetupExe -Arguments "/download `"$ConfigDown`""

        while (-not $dlProc.HasExited) {
            $currentSize = Get-FolderSize $OfficeDir
            Write-Host "`r    Downloaded so far: $(Format-Bytes $currentSize)    " -NoNewline
            Start-Sleep -Seconds 15
        }
        Write-Host ""

        # WaitForExit() flushes the exit code - HasExited alone can leave it null
        $dlProc.WaitForExit()
        $dlExitCode = $dlProc.ExitCode

        if ($dlExitCode -eq 0) {
            $dlSuccess = $true
            break
        }

        # Find ODT log for diagnosis
        $odtLog = Get-ChildItem $LogDir, $env:TEMP, "C:\Windows\Temp" -Filter "*.log" -Recurse -ErrorAction SilentlyContinue |
                  Where-Object { $_.Name -match "Setup|Office|ODT" } |
                  Sort-Object LastWriteTime -Descending | Select-Object -First 1

        Write-Fail "setup.exe /download exited with code $dlExitCode on attempt $attempt."
        if ($odtLog) {
            Write-Warn "Last 20 lines of $($odtLog.FullName):"
            Get-Content $odtLog.FullName -Tail 20 -ErrorAction SilentlyContinue |
                ForEach-Object { Write-Host "    $_" -ForegroundColor DarkYellow }
        } else {
            Write-Warn "No ODT log found. Check %TEMP% and C:\Windows\Temp for Setup*.log files."
        }

        if ($attempt -eq $maxAttempts) {
            throw "Office source download failed after $maxAttempts attempts. Last exit code: $dlExitCode"
        }
    }

    # Verify download size
    $sourceSize = Get-FolderSize $OfficeDir
    Write-Info "Total downloaded: $(Format-Bytes $sourceSize)"
    if ($sourceSize -lt $MinSourceBytes) {
        throw "Downloaded source is only $(Format-Bytes $sourceSize) - expected at least $(Format-Bytes $MinSourceBytes). Download may be incomplete."
    }
    Write-Ok "Source download complete and verified ($(Format-Bytes $sourceSize))."

    if ($DownloadOnly) {
        Write-Host ""
        Write-Host "==> Download complete." -ForegroundColor Green
        Write-Host "    Copy '$WorkDir' to the target machine, then run:" -ForegroundColor Green
        Write-Host "    .\Install-M365Offline.ps1 -InstallOnly -WorkDir <path>" -ForegroundColor Yellow
        exit 0
    }
}

# ---------------------------------------------------------------------------
# 4. Install
# ---------------------------------------------------------------------------
Write-Step "Writing install config..."

$excludeXml = Build-ExcludeXml
$instXml  = "<Configuration>`n"
$instXml += "  <Add OfficeClientEdition=`"$Architecture`"`n"
$instXml += "       Channel=`"$Channel`"`n"
$instXml += "       SourcePath=`"$OfficeDir`"`n"
$instXml += "       AllowCdnFallback=`"FALSE`">`n"
$instXml += "    <Product ID=`"O365ProPlusRetail`">`n"
$instXml += "      <Language ID=`"$Language`" />`n"
$instXml += $excludeXml
$instXml += "    </Product>`n"
$instXml += "  </Add>`n"
$instXml += "  <Updates Enabled=`"FALSE`" />`n"
$instXml += "  <Display Level=`"Full`" AcceptEULA=`"TRUE`" />`n"
$instXml += "  <Logging Level=`"Standard`" Path=`"$LogDir`" />`n"
$instXml += "  <Property Name=`"AUTOACTIVATE`" Value=`"1`" />`n"
$instXml += "</Configuration>"
Write-XmlFile $ConfigInst $instXml
Write-Ok "Install config written."

Confirm-Step "Ready to install Microsoft 365 Apps for Enterprise. This will take 5-15 minutes. Proceed?"

Write-Step "Installing Microsoft 365 Apps for Enterprise..."
Write-Info "Installation is running silently. Please wait..."

$instProc = Start-TrackedProcess -FilePath $SetupExe -Arguments "/configure `"$ConfigInst`""
$instProc.WaitForExit()
Assert-ExitCode $instProc.ExitCode "Office installation"

# ---------------------------------------------------------------------------
# 5. Post-install verification
# ---------------------------------------------------------------------------
Write-Step "Verifying installation..."

# Check registry
$installed = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" `
    -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -match "Microsoft 365|Office 365" }

if ($installed) {
    $installed | ForEach-Object {
        Write-Ok "Registry: $($_.DisplayName) $($_.DisplayVersion)"
    }
} else {
    Write-Warn "No Office entry found in Uninstall registry. Checking file system..."
}

# Check key executables
$officeExes = @(
    "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
    "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
)
$foundExes = 0
foreach ($exe in $officeExes) {
    if (Test-Path $exe) {
        Write-Ok "Found: $exe"
        $foundExes++
    } else {
        Write-Warn "Not found: $exe"
    }
}

if ($foundExes -eq 0 -and -not $installed) {
    Write-Fail "Could not verify installation. Check the ODT log files in $LogDir for errors."
    exit 1
}

# Check log for errors
$odtLog = Get-ChildItem $LogDir -Filter "*.log" -ErrorAction SilentlyContinue |
          Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($odtLog) {
    $errors = Select-String -Path $odtLog.FullName -Pattern "FAILED|ERROR" -ErrorAction SilentlyContinue
    if ($errors) {
        Write-Warn "Errors found in $($odtLog.Name):"
        $errors | Select-Object -First 10 | ForEach-Object { Write-Warn "  $($_.Line.Trim())" }
    } else {
        Write-Ok "No errors in ODT log ($($odtLog.Name))."
    }
}

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "  Microsoft 365 Apps installed successfully." -ForegroundColor Green
Write-Host "  Log files: $LogDir" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
