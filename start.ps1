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

    # Download with dynamic progress bar
    $wc = New-Object Net.WebClient
    $global:downloadComplete = $false

    Register-ObjectEvent -InputObject $wc -EventName DownloadProgressChanged -Action {
        $percent = $EventArgs.ProgressPercentage
        $received = [math]::Round($EventArgs.BytesReceived / 1KB, 0)

        # Dynamic width based on window size
        $width = $Host.UI.RawUI.WindowSize.Width - 25
        if ($width -lt 10) { $width = 10 }

        $done = [math]::Floor($width * $percent / 100)
        $left = $width - $done
        $bar = '[' + ('=' * $done) + (' ' * $left) + ']'

        Write-Host -NoNewline "`r       $bar $percent% ($received KB)"
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
        Write-Host ""
        Write-Host "[ERREUR] Echec du telechargement" -ForegroundColor Red
        Write-Host ""
        Write-Host "         CAUSES POSSIBLES:" -ForegroundColor Yellow
        Write-Host "         - Pas de connexion Internet"
        Write-Host "         - GitHub est inaccessible ou bloque"
        Write-Host "         - Pare-feu/Proxy bloquant la connexion"
        Write-Host ""
        Write-Host "         SOLUTIONS:" -ForegroundColor Yellow
        Write-Host "         1. Verifiez votre connexion Internet"
        Write-Host "         2. Essayez d'acceder a github.com"
        Write-Host "         3. Desactivez temporairement le pare-feu"
        Write-Host ""
    }
} catch {
    Write-Host ""
    Write-Host "[ERREUR] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "         CAUSES POSSIBLES:" -ForegroundColor Yellow
    Write-Host "         - Erreur reseau"
    Write-Host "         - PowerShell bloque par la politique de securite"
    Write-Host "         - Antivirus interferant"
    Write-Host ""
    Write-Host "         SOLUTIONS:" -ForegroundColor Yellow
    Write-Host "         1. Reessayez la commande"
    Write-Host "         2. Executez: Set-ExecutionPolicy Bypass -Scope Process"
    Write-Host "         3. Desactivez temporairement l'antivirus"
    Write-Host ""
}
