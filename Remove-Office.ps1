#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Fully removes Microsoft Office / Microsoft 365 Apps from Windows.

.DESCRIPTION
    Automates every step from Microsoft's official manual uninstall guide:
    https://support.microsoft.com/en-US/Office/lifecycle/manually-uninstall-office

    Covers Click-to-Run (M365/Office 2016+), MSI, and Microsoft Store installs.
    Safe to run even if Office was already removed from Add/Remove Programs —
    it will still clean up all leftover files, folders, services, and registry keys.

.PARAMETER BackupPath
    Where to export registry backups before deletion.
    Default: %USERPROFILE%\OfficeRemovalBackup

.PARAMETER SkipBackup
    Skip the registry backup step.

.PARAMETER Force
    Skip all confirmation prompts.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $BackupPath = "$env:USERPROFILE\OfficeRemovalBackup",
    [switch] $SkipBackup,
    [switch] $Force
)

$ErrorActionPreference = "Continue"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Write-Step { param($Msg) Write-Host "`n==> $Msg" -ForegroundColor Cyan }
function Write-Ok   { param($Msg) Write-Host "    [OK]   $Msg" -ForegroundColor Green }
function Write-Warn { param($Msg) Write-Host "    [WARN] $Msg" -ForegroundColor Yellow }
function Write-Info { param($Msg) Write-Host "    [INFO] $Msg" -ForegroundColor Gray }
function Write-Fail { param($Msg) Write-Host "    [FAIL] $Msg" -ForegroundColor Red }

function Confirm-Action {
    param($Prompt)
    if ($Force) { return }
    $ans = Read-Host "`n$Prompt [Y/N]"
    if ($ans -notmatch "^[Yy]") { throw "Aborted by user." }
}

function Remove-RegKey {
    param([string]$Path)
    if (Test-Path $Path) {
        try   { Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop; Write-Ok "Deleted key: $Path" }
        catch { Write-Warn "Could not delete $Path`: $_" }
    } else {
        Write-Info "Not found (skip): $Path"
    }
}

function Remove-Dir {
    param([string]$Path)
    if (Test-Path $Path) {
        try   { Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop; Write-Ok "Deleted: $Path" }
        catch { Write-Warn "Could not delete $Path`: $_" }
    } else {
        Write-Info "Not found (skip): $Path"
    }
}

function Remove-File {
    param([string]$Path)
    if (Test-Path $Path) {
        try   { Remove-Item -Path $Path -Force -ErrorAction Stop; Write-Ok "Deleted file: $Path" }
        catch { Write-Warn "Could not delete $Path`: $_" }
    } else {
        Write-Info "Not found (skip): $Path"
    }
}

function Export-RegKey {
    param([string]$HivePath, [string]$OutFile)
    if (Test-Path $HivePath) {
        reg export $HivePath $OutFile /y 2>$null | Out-Null
        if (Test-Path $OutFile) { Write-Ok "Backed up: $HivePath" }
    }
}

function Run-MsiExec {
    param([string]$ProductCode)
    Write-Info "MsiExec /X $ProductCode"
    $p = Start-Process "msiexec.exe" -ArgumentList "/X $ProductCode /qn /norestart" -Wait -PassThru -NoNewWindow
    if ($p.ExitCode -eq 0 -or $p.ExitCode -eq 1605) {
        Write-Ok "Uninstalled or not present: $ProductCode"
    } else {
        Write-Warn "msiexec exited $($p.ExitCode) for $ProductCode (may already be gone)"
    }
}

# ---------------------------------------------------------------------------
# Banner
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "==========================================================" -ForegroundColor White
Write-Host "  Office / Microsoft 365 - Full Manual Removal"            -ForegroundColor White
Write-Host "==========================================================" -ForegroundColor White
Write-Host "  Ref: support.microsoft.com - manually-uninstall-office"  -ForegroundColor DarkGray

# ---------------------------------------------------------------------------
# Detect architecture
# ---------------------------------------------------------------------------
$is64OS  = [Environment]::Is64BitOperatingSystem
$is64Off = Test-Path "${env:ProgramFiles}\Microsoft Office"
Write-Info "OS: $(if ($is64OS) { '64-bit' } else { '32-bit' })"

# ---------------------------------------------------------------------------
# Detect what is installed
# ---------------------------------------------------------------------------
Write-Step "Detecting Office installations..."

$uninstallPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*")
if ($is64OS) { $uninstallPaths += "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" }

$officeProducts = Get-ItemProperty $uninstallPaths -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -match "Microsoft (365|Office|OneNote|Visio|Project)" }

if ($officeProducts) {
    Write-Warn "Found installed Office products:"
    $officeProducts | ForEach-Object { Write-Warn "  $($_.DisplayName) $($_.DisplayVersion)" }
} else {
    Write-Info "No Office products found in Add/Remove Programs."
    Write-Info "This is normal after a partial or manual uninstall."
    Write-Info "All file, folder, service, and registry cleanup will still run."
}

$isC2R    = (Get-Service "ClickToRunSvc" -ErrorAction SilentlyContinue) -or (Test-Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun")
$isMSI    = $officeProducts | Where-Object { $_.UninstallString -match "msiexec" }
$isStore  = Get-AppxPackage -Name "Microsoft.Office.Desktop" -ErrorAction SilentlyContinue

Write-Info "Click-to-Run detected : $(if ($isC2R)   { 'YES' } else { 'no' })"
Write-Info "MSI install detected  : $(if ($isMSI)   { 'YES' } else { 'no' })"
Write-Info "Store app detected    : $(if ($isStore) { 'YES' } else { 'no' })"

Confirm-Action "All Office components will be removed. This cannot be undone. Proceed?"

# ---------------------------------------------------------------------------
# REGISTRY BACKUP (before touching anything)
# ---------------------------------------------------------------------------
if (-not $SkipBackup) {
    Write-Step "Backing up registry keys to $BackupPath..."
    New-Item -ItemType Directory -Force -Path $BackupPath | Out-Null

    @(
        @{ H = "HKCU\Software\Microsoft\Office";                     F = "HKCU_Office.reg" },
        @{ H = "HKLM\SOFTWARE\Microsoft\Office";                     F = "HKLM_Office.reg" },
        @{ H = "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office";         F = "HKLM_Wow_Office.reg" },
        @{ H = "HKLM\SOFTWARE\Microsoft\AppVISV";                    F = "HKLM_AppVISV.reg" },
        @{ H = "HKLM\SYSTEM\CurrentControlSet\Services\ose";         F = "HKLM_ose.reg" },
        @{ H = "HKLM\SYSTEM\CurrentControlSet\Services\ClickToRunSvc"; F = "HKLM_C2RSvc.reg" }
    ) | ForEach-Object { Export-RegKey -HivePath $_.H -OutFile (Join-Path $BackupPath $_.F) }

    Write-Ok "Backup complete: $BackupPath"
}

# ===========================================================================
# CLICK-TO-RUN STEPS
# Ref: https://support.microsoft.com/en-US/Office/lifecycle/manually-uninstall-office
# ===========================================================================

# ---------------------------------------------------------------------------
# C2R Step 1 — Delete Microsoft Office 16 from C:\Program Files
# ---------------------------------------------------------------------------
Write-Step "[C2R Step 1] Remove Microsoft Office 16 from Program Files..."
Remove-Dir "$env:ProgramFiles\Microsoft Office 16"

# ---------------------------------------------------------------------------
# C2R Step 2 — Remove scheduled tasks
# ---------------------------------------------------------------------------
Write-Step "[C2R Step 2] Remove Office scheduled tasks..."

@(
    "Office Automatic Updates",
    "Office Subscription Maintenance",
    "Office ClickToRun Service Monitor",
    "OfficeTelemetryAgentLogOn2016",
    "OfficeTelemetryAgentFallBack2016"
) | ForEach-Object {
    $task = Get-ScheduledTask -TaskPath "\Microsoft\Office\" -TaskName $_ -ErrorAction SilentlyContinue
    if ($task) {
        Unregister-ScheduledTask -TaskPath "\Microsoft\Office\" -TaskName $_ -Confirm:$false -ErrorAction SilentlyContinue
        Write-Ok "Removed task: $_"
    } else {
        Write-Info "Task not found (skip): $_"
    }
}

# ---------------------------------------------------------------------------
# C2R Step 3 — End Office processes
# ---------------------------------------------------------------------------
Write-Step "[C2R Step 3] Terminate Office processes..."

@("OfficeClickToRun","OfficeC2RClient","AppVShNotify") | ForEach-Object {
    $procs = Get-Process -Name $_ -ErrorAction SilentlyContinue
    if ($procs) {
        $procs | Stop-Process -Force -ErrorAction SilentlyContinue
        Write-Ok "Stopped: $_ ($($procs.Count) instance(s))"
    } else {
        Write-Info "Not running: $_"
    }
}

# setup*.exe from Office directories
Get-Process -Name "setup" -ErrorAction SilentlyContinue |
    Where-Object { $_.Path -match "Microsoft Office|ClickToRun" } |
    ForEach-Object { Stop-Process $_ -Force -ErrorAction SilentlyContinue; Write-Ok "Stopped: setup.exe (PID $($_.Id))" }

# ---------------------------------------------------------------------------
# C2R Step 4 — Delete ClickToRunSvc service
# ---------------------------------------------------------------------------
Write-Step "[C2R Step 4] Delete ClickToRunSvc service..."

$svc = Get-Service "ClickToRunSvc" -ErrorAction SilentlyContinue
if ($svc) {
    Stop-Service "ClickToRunSvc" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    sc.exe delete ClickToRunSvc | Out-Null
    Write-Ok "ClickToRunSvc deleted."
} else {
    Write-Info "ClickToRunSvc not found (skip)."
}

# ---------------------------------------------------------------------------
# C2R Step 5 — Delete Office files from all listed locations
# ---------------------------------------------------------------------------
Write-Step "[C2R Step 5] Delete Office files..."

Remove-Dir "$env:ProgramFiles\Microsoft Office 16"
Remove-Dir "$env:ProgramFiles\Microsoft Office"
Remove-Dir "${env:ProgramFiles(x86)}\Microsoft Office"
Remove-Dir "$env:CommonProgramFiles\Microsoft Shared\ClickToRun"
Remove-Dir "$env:ProgramData\Microsoft\ClickToRun"
# Page lists this as a file, not a folder
Remove-File "$env:ProgramData\Microsoft\Office\ClickToRunPackagerLocker"

# ---------------------------------------------------------------------------
# C2R Step 6 — Delete registry subkeys
# ---------------------------------------------------------------------------
Write-Step "[C2R Step 6] Delete Click-to-Run registry keys..."

Remove-RegKey "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun"
Remove-RegKey "HKLM:\SOFTWARE\Microsoft\AppVISV"
Remove-RegKey "HKCU:\Software\Microsoft\Office"

# Uninstall entry: HKLM\...\Uninstall\Microsoft Office <Edition> - en-us
$uninstallBase = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
Get-ChildItem $uninstallBase -ErrorAction SilentlyContinue |
    Where-Object { (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).DisplayName -match "Microsoft (365|Office)" } |
    ForEach-Object { Remove-RegKey $_.PSPath }

if ($is64OS) {
    $uninstallBaseWow = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    Get-ChildItem $uninstallBaseWow -ErrorAction SilentlyContinue |
        Where-Object { (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).DisplayName -match "Microsoft (365|Office)" } |
        ForEach-Object { Remove-RegKey $_.PSPath }
}

# ---------------------------------------------------------------------------
# C2R Step 7 — Delete Start menu shortcuts
# ---------------------------------------------------------------------------
Write-Step "[C2R Step 7] Delete Start menu shortcuts..."

$startMenuRoot = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs"
Remove-Dir "$startMenuRoot\Microsoft Office 2016 Tools"
Remove-Dir "$startMenuRoot\Microsoft Office"

Get-ChildItem $startMenuRoot -Filter "*.lnk" -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match "Word|Excel|PowerPoint|Outlook|OneNote|Access|Publisher|Visio|Project" } |
    ForEach-Object { Remove-File $_.FullName }

# ---------------------------------------------------------------------------
# C2R Step 8 — Uninstall licensing and extensibility components via MsiExec
# ---------------------------------------------------------------------------
Write-Step "[C2R Step 8] Uninstall licensing/extensibility components..."

if ($is64OS -and -not $is64Off) {
    # x86 Office on x64 Windows
    Write-Info "Architecture: x86 Office on x64 Windows"
    Run-MsiExec "{90160000-008F-0000-1000-0000000FF1CE}"
    Run-MsiExec "{90160000-008C-0000-0000-0000000FF1CE}"
    Run-MsiExec "{90160000-008C-0409-0000-0000000FF1CE}"
} elseif (-not $is64OS) {
    # x86 Office on x86 Windows
    Write-Info "Architecture: x86 Office on x86 Windows"
    Run-MsiExec "{90160000-007E-0000-0000-0000000FF1CE}"
    Run-MsiExec "{90160000-008C-0000-0000-0000000FF1CE}"
    Run-MsiExec "{90160000-008C-0409-0000-0000000FF1CE}"
} else {
    # x64 Office on x64 Windows
    Write-Info "Architecture: x64 Office on x64 Windows"
    Run-MsiExec "{90160000-007E-0000-1000-0000000FF1CE}"
    Run-MsiExec "{90160000-008C-0000-1000-0000000FF1CE}"
    Run-MsiExec "{90160000-008C-0409-1000-0000000FF1CE}"
}

# ===========================================================================
# MSI INSTALLATION STEPS
# ===========================================================================

# ---------------------------------------------------------------------------
# MSI Step 1 — Uninstall .msi packages from %windir%\Installer
# (Page: sort by Subject, find "Microsoft Office 2016", right-click Uninstall)
# Automated via registry product codes instead of file subject scanning
# ---------------------------------------------------------------------------
Write-Step "[MSI Step 1] Uninstall remaining MSI Office packages..."

$msiProducts = Get-ItemProperty $uninstallPaths -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -match "Microsoft (365|Office)" -and $_.UninstallString -match "msiexec" }

if ($msiProducts) {
    foreach ($prod in $msiProducts) {
        if ($prod.PSChildName -match "^\{") {
            Write-Info "Uninstalling: $($prod.DisplayName)"
            Run-MsiExec $prod.PSChildName
        }
    }
} else {
    Write-Info "No MSI Office packages found in registry."
}

# ---------------------------------------------------------------------------
# MSI Step 2 — Stop Office Source Engine service
# ---------------------------------------------------------------------------
Write-Step "[MSI Step 2] Stop Office Source Engine service..."

$ose = Get-Service "ose" -ErrorAction SilentlyContinue
if ($ose) {
    Stop-Service "ose" -Force -ErrorAction SilentlyContinue
    Write-Ok "Office Source Engine (ose) stopped."
} else {
    Write-Info "Office Source Engine not found (skip)."
}

# ---------------------------------------------------------------------------
# MSI Step 3 — Delete MSI installation folders
# ---------------------------------------------------------------------------
Write-Step "[MSI Step 3] Delete MSI installation folders..."

Remove-Dir "$env:CommonProgramFiles\Microsoft Shared\Office16"
Remove-Dir "$env:CommonProgramFiles\Microsoft Shared\Source Engine"
if ($is64OS) {
    Remove-Dir "${env:CommonProgramFiles(x86)}\Microsoft Shared\Office16"
    Remove-Dir "${env:CommonProgramFiles(x86)}\Microsoft Shared\Source Engine"
}
Remove-Dir "$env:ProgramFiles\Microsoft Office\Office16"

# MSOCache: folders containing "0FF1CE}" in the name
Get-ChildItem "C:\MSOCache\All Users" -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match "0FF1CE\}" } |
    ForEach-Object { Remove-Dir $_.FullName }

# ---------------------------------------------------------------------------
# MSI Step 4 — Delete specific installation files (NOT whole folders)
# Normal.dotm and Welcome to Word.dotx from Templates
# Building blocks.dotx from Document Building Blocks
# ---------------------------------------------------------------------------
Write-Step "[MSI Step 4] Delete specific Office template and building block files..."

Remove-File "$env:APPDATA\Microsoft\Templates\Normal.dotm"
Remove-File "$env:APPDATA\Microsoft\Templates\Welcome to Word.dotx"

# Building blocks — language subfolder varies, search for the file
Get-ChildItem "$env:APPDATA\Microsoft\Document Building Blocks" -Recurse -Filter "Building blocks.dotx" -ErrorAction SilentlyContinue |
    ForEach-Object { Remove-File $_.FullName }

# ---------------------------------------------------------------------------
# MSI Step 5 — Delete MSI registry subkeys
# ---------------------------------------------------------------------------
Write-Step "[MSI Step 5] Delete MSI registry keys..."

# Current user
Remove-RegKey "HKCU:\Software\Microsoft\Office\16.0"

# Machine (both native and WOW paths on x64)
Remove-RegKey "HKLM:\SOFTWARE\Microsoft\Office\16.0"
if ($is64OS) { Remove-RegKey "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0" }

# Delivery/SourceEngine (*0FF1CE}-*)
$deliveryBase = "HKLM:\SOFTWARE\Microsoft\Office\Delivery\SourceEngine\Downloads"
if (Test-Path $deliveryBase) {
    Get-ChildItem $deliveryBase -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match "0FF1CE" } |
        ForEach-Object { Remove-RegKey $_.PSPath }
}

# Uninstall entries (*0FF1CE*)
foreach ($base in $uninstallPaths) {
    $basePath = $base.TrimEnd("\*")
    if (Test-Path $basePath) {
        Get-ChildItem $basePath -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match "0FF1CE" } |
            ForEach-Object { Remove-RegKey $_.PSPath }
    }
}

# Installer UpgradeCodes (*F01FEC)
$upgradeBase = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes"
if (Test-Path $upgradeBase) {
    Get-ChildItem $upgradeBase -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match "F01FEC" } |
        ForEach-Object { Remove-RegKey $_.PSPath }
}

# UserData (*F01FEC)
$userData = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products"
if (Test-Path $userData) {
    Get-ChildItem $userData -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match "F01FEC" } |
        ForEach-Object { Remove-RegKey $_.PSPath }
}

# Office Source Engine service key
Remove-RegKey "HKLM:\SYSTEM\CurrentControlSet\Services\ose"

# HKCR Installer entries (*F01FEC)
foreach ($hkcrBase in @("HKCR:\Installer\Features","HKCR:\Installer\Products","HKCR:\Installer\UpgradeCodes")) {
    if (Test-Path $hkcrBase) {
        Get-ChildItem $hkcrBase -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match "F01FEC" } |
            ForEach-Object { Remove-RegKey $_.PSPath }
    }
}

# Win32Assemblies (*Office16*)
$win32Asm = "HKCR:\Installer\Win32Assemblies"
if (Test-Path $win32Asm) {
    Get-ChildItem $win32Asm -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match "Office16" } |
        ForEach-Object { Remove-RegKey $_.PSPath }
}

# ===========================================================================
# STORE APP
# ===========================================================================
Write-Step "[Store] Remove Microsoft Store Office app (if present)..."

$storeApp = Get-AppxPackage -Name "Microsoft.Office.Desktop" -ErrorAction SilentlyContinue
if ($storeApp) {
    try {
        $storeApp | Remove-AppxPackage -ErrorAction Stop
        Write-Ok "Store Office app removed."
        # Verify
        $check = Get-AppxPackage -Name "Microsoft.Office.Desktop" -ErrorAction SilentlyContinue
        if (-not $check) { Write-Ok "Verified: Store app is gone." }
        else              { Write-Warn "Store app may still be present." }
    } catch {
        Write-Warn "Could not remove Store app: $_"
    }
} else {
    Write-Info "No Store Office app found (skip)."
}

# ===========================================================================
# DONE
# ===========================================================================
Write-Host ""
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "  Office removal complete."                                 -ForegroundColor Green
if (-not $SkipBackup) {
    Write-Host "  Registry backup: $BackupPath"                        -ForegroundColor Green
}
Write-Host "  A restart is strongly recommended."                      -ForegroundColor Yellow
Write-Host "==========================================================" -ForegroundColor Green

if (-not $Force) {
    $restart = Read-Host "`nRestart now? [Y/N]"
    if ($restart -match "^[Yy]") { Restart-Computer -Force }
}
