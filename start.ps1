<#
.SYNOPSIS
    Office Unattended Installation and Activation Script
.DESCRIPTION
    Installs and activates Microsoft Office using Ohook method
    Usage: irm https://office-unnatended.vercel.app | iex
.NOTES
    Based on ohook by asdcorp - https://github.com/asdcorp/ohook
#>

# Configuration - Base URL for scripts
$BaseURL = "https://raw.githubusercontent.com/Louchatfroff/Office-Unnatended-Install/main"

$host.UI.RawUI.WindowTitle = "Office Install & Activate"

# Check Admin
function Test-Admin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Elevate if needed
if (-not (Test-Admin)) {
    Write-Host "`n[!] This script requires Administrator privileges." -ForegroundColor Yellow
    Write-Host "[*] Relaunching as Administrator...`n" -ForegroundColor Cyan

    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; iex ((New-Object Net.WebClient).DownloadString('$BaseURL/start.ps1'))`"" -Verb RunAs
    exit
}

# Main
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Office Unattended Install and Activate" -ForegroundColor White
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

try {
    Write-Host "[*] Downloading installation script..." -ForegroundColor Cyan
    $scriptPath = "$env:TEMP\Office-Install-Activate.cmd"
    $url = "$BaseURL/Office-Install-Activate.cmd"

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $ProgressPreference = 'Continue'

    # Download with Invoke-WebRequest (reliable with built-in progress)
    Invoke-WebRequest -Uri $url -OutFile $scriptPath -UseBasicParsing

    if (Test-Path $scriptPath) {
        Write-Host "[OK] Script downloaded" -ForegroundColor Green
        Write-Host "[*] Launching installation...`n" -ForegroundColor Cyan

        # Execute CMD script
        Start-Process cmd.exe -ArgumentList "/c `"$scriptPath`"" -Wait

        # Cleanup
        Remove-Item $scriptPath -Force -ErrorAction SilentlyContinue
    } else {
        Write-Host ""
        Write-Host "[ERROR] Download failed" -ForegroundColor Red
        Write-Host ""
        Write-Host "         POSSIBLE CAUSES:" -ForegroundColor Yellow
        Write-Host "         - No Internet connection"
        Write-Host "         - GitHub is unreachable or blocked"
        Write-Host "         - Firewall/Proxy blocking connection"
        Write-Host ""
        Write-Host "         SOLUTIONS:" -ForegroundColor Yellow
        Write-Host "         1. Check your Internet connection"
        Write-Host "         2. Try accessing github.com"
        Write-Host "         3. Temporarily disable firewall"
        Write-Host ""
    }
} catch {
    Write-Host ""
    Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "         POSSIBLE CAUSES:" -ForegroundColor Yellow
    Write-Host "         - Network error"
    Write-Host "         - PowerShell blocked by security policy"
    Write-Host "         - Antivirus interference"
    Write-Host ""
    Write-Host "         SOLUTIONS:" -ForegroundColor Yellow
    Write-Host "         1. Retry the command"
    Write-Host "         2. Run: Set-ExecutionPolicy Bypass -Scope Process"
    Write-Host "         3. Temporarily disable antivirus"
    Write-Host ""
}
