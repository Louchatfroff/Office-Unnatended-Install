<#
.SYNOPSIS
    Office Unattended Installation and Activation Script
.DESCRIPTION
    Installs and activates Microsoft Office using Ohook method
    Usage: irm https://YOUR-URL.vercel.app | iex
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
    Write-Host "`n[!] Ce script necessite les droits Administrateur." -ForegroundColor Yellow
    Write-Host "[*] Relancement en tant qu'Administrateur...`n" -ForegroundColor Cyan

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
    Write-Host "[*] Telechargement du script d'installation..." -ForegroundColor Cyan
    $scriptPath = "$env:TEMP\Office-Install-Activate.cmd"
    $url = "$BaseURL/Office-Install-Activate.cmd"

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Download with progress
    $wc = New-Object Net.WebClient
    $downloadComplete = $false

    Register-ObjectEvent -InputObject $wc -EventName DownloadProgressChanged -Action {
        $percent = $EventArgs.ProgressPercentage
        $received = [math]::Round($EventArgs.BytesReceived / 1KB, 0)
        Write-Host -NoNewline "`r       Progress: $percent% - $received KB"
    } | Out-Null

    Register-ObjectEvent -InputObject $wc -EventName DownloadFileCompleted -Action {
        $global:downloadComplete = $true
    } | Out-Null

    $wc.DownloadFileAsync([Uri]$url, $scriptPath)

    while (-not $global:downloadComplete) {
        Start-Sleep -Milliseconds 100
    }
    Write-Host ""

    if (Test-Path $scriptPath) {
        Write-Host "[OK] Script telecharge" -ForegroundColor Green
        Write-Host "[*] Lancement de l'installation...`n" -ForegroundColor Cyan

        # Execute CMD script
        Start-Process cmd.exe -ArgumentList "/c `"$scriptPath`"" -Wait

        # Cleanup
        Remove-Item $scriptPath -Force -ErrorAction SilentlyContinue
    } else {
        Write-Host "[ERREUR] Echec du telechargement" -ForegroundColor Red
    }
} catch {
    Write-Host "[ERREUR] $($_.Exception.Message)" -ForegroundColor Red
}
