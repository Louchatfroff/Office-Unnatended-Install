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

# Menu
function Show-Menu {
    Clear-Host
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host "       OFFICE UNATTENDED INSTALLATION AND ACTIVATION" -ForegroundColor White
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "       [" -NoNewline; Write-Host "1" -ForegroundColor Yellow -NoNewline; Write-Host "] Office 365 ProPlus (Recommande)"
    Write-Host "       [" -NoNewline; Write-Host "2" -ForegroundColor Yellow -NoNewline; Write-Host "] Office 2021 LTSC Professional Plus"
    Write-Host "       [" -NoNewline; Write-Host "3" -ForegroundColor Yellow -NoNewline; Write-Host "] Office 2019 Professional Plus"
    Write-Host ""
    Write-Host "       [" -NoNewline; Write-Host "4" -ForegroundColor Yellow -NoNewline; Write-Host "] Activer Office uniquement (deja installe)"
    Write-Host "       [" -NoNewline; Write-Host "5" -ForegroundColor Yellow -NoNewline; Write-Host "] Desactiver telemetrie et recommandations"
    Write-Host "       [" -NoNewline; Write-Host "6" -ForegroundColor Yellow -NoNewline; Write-Host "] Verifier le statut d'activation"
    Write-Host ""
    Write-Host "       [" -NoNewline; Write-Host "0" -ForegroundColor Red -NoNewline; Write-Host "] Quitter"
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host ""
}

# Download and run CMD script
function Invoke-RemoteScript {
    param([string]$ScriptName)

    try {
        Write-Host "[*] Telechargement de $ScriptName..." -ForegroundColor Cyan
        $scriptPath = "$env:TEMP\$ScriptName"

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        (New-Object Net.WebClient).DownloadFile("$BaseURL/$ScriptName", $scriptPath)

        if (Test-Path $scriptPath) {
            Write-Host "[OK] Script telecharge" -ForegroundColor Green

            # Execute CMD script
            $process = Start-Process cmd.exe -ArgumentList "/c `"$scriptPath`"" -Wait -PassThru

            # Cleanup
            Remove-Item $scriptPath -Force -ErrorAction SilentlyContinue
        } else {
            Write-Host "[ERREUR] Echec du telechargement" -ForegroundColor Red
        }
    } catch {
        Write-Host "[ERREUR] $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Install Office
function Install-Office {
    param(
        [string]$Product,
        [string]$Channel,
        [string]$PidKey = ""
    )

    $tempDir = "$env:TEMP\OfficeInstall_$((Get-Random))"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    $odtPath = "$tempDir\setup.exe"
    $configPath = "$tempDir\config.xml"

    # Download ODT
    Write-Host "`n[*] Telechargement de Office Deployment Tool..." -ForegroundColor Cyan
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        (New-Object Net.WebClient).DownloadFile("https://officecdn.microsoft.com/pr/wsus/setup.exe", $odtPath)
        Write-Host "[OK] ODT telecharge" -ForegroundColor Green
    } catch {
        Write-Host "[ERREUR] Echec du telechargement de ODT: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    # Create config
    Write-Host "[*] Creation de la configuration..." -ForegroundColor Cyan

    $pidKeyAttr = if ($PidKey) { " PIDKEY=`"$PidKey`"" } else { "" }

    $configContent = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="$Channel">
    <Product ID="$Product"$pidKeyAttr>
      <Language ID="fr-fr" />
      <Language ID="MatchOS" />
      <ExcludeApp ID="Publisher" />
      <ExcludeApp ID="Access" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="Teams" />
    </Product>
  </Add>
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="SharedComputerLicensing" Value="0" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@

    $configContent | Out-File -FilePath $configPath -Encoding UTF8
    Write-Host "[OK] Configuration creee" -ForegroundColor Green

    # Download Office
    Write-Host "[*] Telechargement des fichiers Office..." -ForegroundColor Cyan
    Write-Host "    Cela peut prendre plusieurs minutes..." -ForegroundColor Gray

    Push-Location $tempDir
    $downloadProcess = Start-Process -FilePath $odtPath -ArgumentList "/download `"$configPath`"" -Wait -PassThru -NoNewWindow
    Pop-Location

    Write-Host "[OK] Telechargement termine" -ForegroundColor Green

    # Install Office
    Write-Host "[*] Installation de Microsoft Office..." -ForegroundColor Cyan
    Write-Host "    Veuillez patienter 5-15 minutes..." -ForegroundColor Gray

    $installProcess = Start-Process -FilePath $odtPath -ArgumentList "/configure `"$configPath`"" -Wait -PassThru -NoNewWindow

    Write-Host "[OK] Installation terminee" -ForegroundColor Green

    # Wait
    Start-Sleep -Seconds 20

    # Activate
    Write-Host "`n[*] Activation avec Ohook..." -ForegroundColor Cyan
    Invoke-RemoteScript -ScriptName "Ohook-Activate.cmd"

    # Disable telemetry
    Write-Host "`n[*] Desactivation de la telemetrie..." -ForegroundColor Cyan
    Invoke-RemoteScript -ScriptName "Disable-Telemetry.cmd"

    # Cleanup
    Write-Host "[*] Nettoyage..." -ForegroundColor Cyan
    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "[OK] Nettoyage termine" -ForegroundColor Green

    Write-Host "`n============================================" -ForegroundColor Green
    Write-Host "   Installation et Activation Terminees!" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Green
}

# Check activation status
function Get-ActivationStatus {
    Write-Host "`n[*] Verification du statut d'activation..." -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan

    $osppPaths = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\OSPP.VBS",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\OSPP.VBS",
        "$env:ProgramFiles\Microsoft Office\Office16\OSPP.VBS",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OSPP.VBS"
    )

    $found = $false
    foreach ($path in $osppPaths) {
        if (Test-Path $path) {
            & cscript //nologo $path /dstatus
            $found = $true
            break
        }
    }

    if (-not $found) {
        Write-Host "[!] Office ne semble pas etre installe ou OSPP.VBS introuvable" -ForegroundColor Yellow
    }

    Write-Host "============================================" -ForegroundColor Cyan
}

# Main loop
do {
    Show-Menu
    $choice = Read-Host "  Votre choix [0-6]"

    switch ($choice) {
        "1" {
            Write-Host "`n[*] Installation de Office 365 ProPlus..." -ForegroundColor Cyan
            Install-Office -Product "O365ProPlusRetail" -Channel "Current"
            Read-Host "`nAppuyez sur Entree pour continuer"
        }
        "2" {
            Write-Host "`n[*] Installation de Office 2021 LTSC..." -ForegroundColor Cyan
            Install-Office -Product "ProPlus2021Volume" -Channel "PerpetualVL2021" -PidKey "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"
            Read-Host "`nAppuyez sur Entree pour continuer"
        }
        "3" {
            Write-Host "`n[*] Installation de Office 2019..." -ForegroundColor Cyan
            Install-Office -Product "ProPlus2019Volume" -Channel "PerpetualVL2019" -PidKey "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"
            Read-Host "`nAppuyez sur Entree pour continuer"
        }
        "4" {
            Invoke-RemoteScript -ScriptName "Ohook-Activate.cmd"
            Read-Host "`nAppuyez sur Entree pour continuer"
        }
        "5" {
            Invoke-RemoteScript -ScriptName "Disable-Telemetry.cmd"
            Read-Host "`nAppuyez sur Entree pour continuer"
        }
        "6" {
            Get-ActivationStatus
            Read-Host "`nAppuyez sur Entree pour continuer"
        }
        "0" {
            Write-Host "`n  Au revoir!`n" -ForegroundColor Cyan
            break
        }
        default {
            Write-Host "`n  [!] Choix invalide" -ForegroundColor Red
            Start-Sleep -Seconds 1
        }
    }
} while ($choice -ne "0")
