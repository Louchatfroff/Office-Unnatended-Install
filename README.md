# Office Unattended Installation and Activation

Scripts for automatic (unattended) installation and activation of Microsoft Office.

Based on [ohook by asdcorp](https://github.com/asdcorp/ohook) for activation.

## Quick Start (Recommended)

Open PowerShell and run:

```powershell
irm https://office-unnatended.vercel.app | iex
```

Or with the direct GitHub URL:

```powershell
irm https://raw.githubusercontent.com/Louchatfroff/Office-Unnatended-Install/main/start.ps1 | iex
```

> **Note:** The XML configuration is automatically generated in the temp folder during each run. No additional files are required.

## Local Usage

### Full Automatic Installation

1. **Right-click** on `Office-Install-Activate.cmd`
2. Select **"Run as administrator"**
3. Wait for installation and activation to complete

### Interactive Menu Version

1. **Right-click** on `Office-Menu-Install.cmd`
2. Select **"Run as administrator"**
3. Choose the desired option from the menu

### Activation Only

If Office is already installed:

```batch
:: Detailed version with logs
Ohook-Activate.cmd

:: Silent version
Ohook-Activate-Silent.cmd

:: Silent version with logs
Ohook-Activate-Silent.cmd /log
```

## Disabling Telemetry

The `Disable-Telemetry.cmd` script disables:

**Telemetry:**
- Windows (DiagTrack, diagnostic data, advertising)
- Microsoft Office (feedback, customer data, LinkedIn)
- Microsoft Edge (metrics, experiments, SmartScreen)

**Recommendations and widgets:**
- Windows 11 Widgets
- Start Menu suggestions
- Explorer recommendations
- Bing search in taskbar
- Copilot
- Teams chat in taskbar

### URLs to Configure

In the installation scripts, configure these variables:

```batch
set "OHOOK_SCRIPT_URL=https://raw.githubusercontent.com/YOUR_USER/YOUR_REPO/main/Ohook-Activate.cmd"
set "TELEMETRY_SCRIPT_URL=https://raw.githubusercontent.com/YOUR_USER/YOUR_REPO/main/Disable-Telemetry.cmd"
```

## How Ohook Works

Ohook works by placing a custom `sppc.dll` file in the Office folder. This file intercepts activation verification calls and always returns that Office is activated.

**Advantages:**
- Does not modify Windows system files
- Survives Office updates
- Does not require a KMS server

**Files used:**
- `sppc64.dll` (64-bit) - SHA256: `393a1fa26deb3663854e41f2b687c188a9eacd87b23f17ea09422c4715cb5a9f`
- `sppc32.dll` (32-bit) - SHA256: `09865ea5993215965e8f27a74b8a41d15fd0f60f5f404cb7a8b3c7757acdab02`

## Configuration

### Change Office Version

Edit `Office-Install-Activate.cmd` and modify these variables:

```batch
:: Office Edition
set "OFFICE_PRODUCT=O365ProPlusRetail"

:: Other options:
:: O365ProPlusRetail    - Office 365 ProPlus (recommended)
:: ProPlus2021Volume    - Office 2021 LTSC Professional Plus
:: ProPlus2019Volume    - Office 2019 Professional Plus
:: Standard2021Volume   - Office 2021 Standard
```

### Change Language

```batch
set "OFFICE_LANG=en-us"

:: Other languages:
:: fr-fr (French)
:: de-de (German)
:: es-es (Spanish)
:: it-it (Italian)
```

### Change Architecture

```batch
set "OFFICE_ARCH=64"

:: Options: 64 or 32
```

### Exclude Applications

```batch
set "EXCLUDE_APPS=Publisher,Access,OneDrive,Teams"

:: Available applications:
:: Access, Excel, OneDrive, OneNote, Outlook
:: PowerPoint, Publisher, Word, Teams
```

### Change Ohook DLL URL

In `Ohook-Activate.cmd` or `Ohook-Activate-Silent.cmd`:

```batch
set "OHOOK_VERSION=0.5"
set "DLL64_URL=https://github.com/asdcorp/ohook/releases/download/%OHOOK_VERSION%/sppc64.dll"
set "DLL32_URL=https://github.com/asdcorp/ohook/releases/download/%OHOOK_VERSION%/sppc32.dll"
```

## Requirements

- Windows 10/11 (or Windows Server 2016+)
- Internet connection
- Administrator privileges
- PowerShell 5.0+

## Supported Products

| Product | ID |
|---------|-----|
| Office 365 ProPlus | `O365ProPlusRetail` |
| Office 2021 Pro Plus | `ProPlus2021Volume` |
| Office 2021 Standard | `Standard2021Volume` |
| Office 2019 Pro Plus | `ProPlus2019Volume` |
| Office 2019 Standard | `Standard2019Volume` |
| Visio 2021 | `VisioPro2021Volume` |
| Project 2021 | `ProjectPro2021Volume` |

## Troubleshooting

### Activation Fails

1. Close all Office applications
2. Rerun the activation script as administrator
3. Restart the computer and try again

### Check Activation Status

```cmd
cscript "%ProgramFiles%\Microsoft Office\root\Office16\OSPP.VBS" /dstatus
```

### Clean Reinstall

1. Uninstall Office via Windows Settings
2. Use [Office Removal Tool](https://aka.ms/SaRA-OfficeUninstallFromPC)
3. Rerun the installation script

### DLLs Won't Download

Check that GitHub is not blocked on your network. You can manually download the DLLs from:
- https://github.com/asdcorp/ohook/releases

## Sources

- [ohook by asdcorp](https://github.com/asdcorp/ohook) - Activation DLL source code
- [Office Deployment Tool](https://docs.microsoft.com/deployoffice/overview-office-deployment-tool) - Office installation tool
- [MAS Documentation](https://massgrave.dev/ohook) - Ohook documentation

## Disclaimer

This script is provided for educational purposes. Use it responsibly and in compliance with the laws in your country.
