# Microsoft 365 Apps — Offline Installer

> Download the full Microsoft 365 Apps for Enterprise installer once, deploy to as many machines as you like — **no internet required at install time.**

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)
![Offline](https://img.shields.io/badge/Install-100%25%20Offline-brightgreen)

---

## Why this exists

The official Microsoft 365 installer phones home constantly — even when you point it at local source files it can silently fall back to Microsoft's CDN, pull updates mid-install, and fail on machines with no internet.

This script fixes that. It uses the **Office Deployment Tool (ODT)** correctly:

- `AllowCdnFallback="FALSE"` — hard blocks any CDN fallback
- `Updates Enabled="FALSE"` — no internet calls during install
- Verified source integrity check before install begins

---

## Features

- **One download, unlimited deploys** — grab the ~3.5 GB source on any internet-connected machine, copy the folder, install everywhere
- **Live download progress** — shows downloaded size every 15 seconds so you know it isn't frozen
- **Auto-retry** — retries the download up to 3 times on network hiccups
- **Pre-flight checks** — validates disk space, internet reachability, and existing Office installs before touching anything
- **Post-install verification** — checks registry entries and key executables to confirm success
- **ODT log dump on failure** — automatically finds and prints the ODT log tail so you know exactly what went wrong
- **Fully parameterised** — channel, language, architecture, excluded apps, custom download path

---

## Requirements

- Windows 10/11
- PowerShell 5.1+
- **Run as Administrator**
- ~10 GB free disk space on the download machine
- ~4 GB free disk space on install-only machines

---

## Step-by-Step Guide

### Scenario A — Download and install on the same machine

**Step 1 — Get the scripts**

Clone the repo or download the ZIP from GitHub (Code → Download ZIP) and extract it.

```powershell
git clone https://github.com/ProperProperProper/m365-offline-installer.git
```

**Step 2 — Open PowerShell as Administrator**

Right-click the Start button → **Terminal (Admin)** or **Windows PowerShell (Admin)**.

**Step 3 — Navigate to the folder**

```powershell
cd "C:\path\to\m365-offline-installer"
```

**Step 4 — Unblock the script for this session**

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
```

**Step 5 — Run the installer**

```powershell
.\Install-M365Offline.ps1
```

The script downloads ~3.5 GB then installs Office. Allow 20–40 minutes total. Progress is printed every 15 seconds during the download so you can confirm it is not frozen.

---

### Scenario B — Download once, deploy offline to many machines

**Steps 1–4** — same as Scenario A on the internet-connected machine.

**Step 5 — Download only (do not install yet)**

```powershell
.\Install-M365Offline.ps1 -DownloadOnly -WorkDir "D:\M365Offline"
```

**Step 6 — Copy files to the target machine**

Copy the entire `D:\M365Offline` folder to each target via USB drive, network share, or any other transfer method. No internet access is needed on the target.

**Step 7 — Install from local files on the target machine**

Open PowerShell as Administrator on the target, then:

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
.\Install-M365Offline.ps1 -InstallOnly -WorkDir "D:\M365Offline"
```

---

### Scenario C — Silent / unattended install (automation, MDM, RMM)

```powershell
.\Install-M365Offline.ps1 -Unattended
```

Skips all Y/N confirmation prompts. Combine with `-DownloadOnly` or `-InstallOnly` as needed.

---

## Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-WorkDir` | `%USERPROFILE%\M365Offline` | Root folder for ODT and downloaded source files |
| `-Channel` | `Current` | Update channel: `Current`, `MonthlyEnterprise`, `SemiAnnual` |
| `-Language` | `en-us` | Language ID (IETF format) |
| `-Architecture` | `64` | Office bitness: `64` or `32` |
| `-ExcludeApps` | *(none)* | Comma-separated apps to exclude e.g. `"Access,Publisher,Groove"` |
| `-DownloadOnly` | `$false` | Download source files, do not install |
| `-InstallOnly` | `$false` | Install from existing source files, skip download |
| `-Unattended` | `$false` | Skip all Y/N confirmation prompts |

---

## What Gets Installed

**Microsoft 365 Apps for Enterprise** (`O365ProPlusRetail`):

| App | Included |
|-----|----------|
| Word | Yes |
| Excel | Yes |
| PowerPoint | Yes |
| Outlook | Yes |
| Teams | Yes |
| OneNote | Yes |
| Access | Yes |
| Publisher | Yes |
| OneDrive | Yes |

To exclude apps you don't need:

```powershell
.\Install-M365Offline.ps1 -ExcludeApps "Access,Publisher,Groove"
```

---

## How It Works

```
1. Pre-flight checks
   - Existing Office detection
   - Disk space validation
   - Internet connectivity test (download phase only)

2. Download ODT
   - Fetches latest setup.exe from Microsoft automatically
   - No hardcoded version URLs

3. Download Office source (~3.5 GB)
   - Saved to WorkDir\OfficeSource
   - CDN fallback disabled — local only

4. Install
   - Reads from local source only
   - Updates disabled during install
   - EULA auto-accepted

5. Post-install verification
   - Registry check for installed products
   - Key executable presence check (Word, Excel, Outlook)
   - ODT log scan for errors
```

---

## Troubleshooting

**Download stalls or retries**
The ODT download can pause between packages. The script retries automatically up to 3 times. If all 3 fail, check your proxy/firewall settings for `download.microsoft.com` and `officecdn.microsoft.com`.

**Exit code errors**
On failure the script automatically finds and prints the last 20 lines of the ODT log from `%TEMP%` so you can see the root cause.

**"Insufficient disk space"**
You need ~10 GB free for download+install, ~4 GB for install-only.

**Existing Office detected**
The script warns you if Office is already installed and asks to confirm before continuing. Use `-Unattended` to auto-confirm.

---

---

## Removal

A companion script automates every step from Microsoft's [manual uninstall guide](https://support.microsoft.com/en-US/Office/lifecycle/manually-uninstall-office):

```powershell
# Run as Administrator
.\Remove-Office.ps1
```

Works even if Office was already partially removed from Add/Remove Programs — cleans up all leftover folders, registry keys, services, scheduled tasks, and shortcuts. Registry keys are backed up to `%USERPROFILE%\OfficeRemovalBackup` before deletion.

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-BackupPath` | `%USERPROFILE%\OfficeRemovalBackup` | Where to save registry backups |
| `-SkipBackup` | `$false` | Skip the registry backup step |
| `-Force` | `$false` | Skip all confirmation prompts |

---

## Related Projects

- [win11-zero-touch-usb](https://github.com/PahtrikProper/win11-zero-touch-usb) — Bootable Windows 11 USB with auto OEM driver injection
- [PrinterMigration](https://github.com/PahtrikProper/PrinterMigration) — Intune GP-to-IP printer migration toolkit

---

## License

MIT — free to use, modify, and distribute.
