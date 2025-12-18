@echo off
:: ============================================================================
:: Office Unattended Installation and Activation Script
:: Based on Microsoft Activation Scripts (MAS) - https://massgrave.dev
:: ============================================================================

setlocal EnableDelayedExpansion
title Office Unattended Install and Activate

:: ============================================================================
:: CONFIGURATION
:: ============================================================================

set "OFFICE_PRODUCT=O365ProPlusRetail"
set "OFFICE_ARCH=64"
set "OFFICE_LANG=fr-fr"
set "OFFICE_CHANNEL=Current"
set "EXCLUDE_APPS=Publisher,Access,OneDrive,Teams"

set "WORK_DIR=%TEMP%\OfficeInstall_%RANDOM%"
set "ODT_EXE=%WORK_DIR%\setup.exe"
set "CONFIG_FILE=%WORK_DIR%\config.xml"

set "OHOOK_SCRIPT_URL=https://raw.githubusercontent.com/Louchatfroff/Office-Unnatended-Install/main/Ohook-Activate.cmd"
set "TELEMETRY_SCRIPT_URL=https://raw.githubusercontent.com/Louchatfroff/Office-Unnatended-Install/main/Disable-Telemetry.cmd"

:: ============================================================================
:: CHECK ADMIN PRIVILEGES
:: ============================================================================

:check_admin
echo.
echo ============================================
echo   Office Unattended Install and Activate
echo ============================================
echo.

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERREUR] Ce script necessite les droits Administrateur.
    echo.
    echo          SOLUTION:
    echo          1. Clic droit sur le script
    echo          2. Selectionnez "Executer en tant qu'administrateur"
    echo.
    echo          OU ouvrez PowerShell en admin et executez:
    echo          irm https://office-unnatended.vercel.app ^| iex
    echo.
    pause
    exit /b 1
)

echo [OK] Droits administrateur confirmes
echo.

:: ============================================================================
:: CREATE TEMP DIRECTORY
:: ============================================================================

:create_dirs
echo [INFO] Creation du dossier temporaire...
if not exist "%WORK_DIR%" mkdir "%WORK_DIR%"
if not exist "%WORK_DIR%" (
    echo [ERREUR] Impossible de creer le dossier temporaire.
    echo.
    echo          CAUSES POSSIBLES:
    echo          - Disque plein
    echo          - Permissions insuffisantes sur %%TEMP%%
    echo          - Antivirus bloquant la creation
    echo.
    echo          SOLUTIONS:
    echo          1. Liberez de l'espace disque
    echo          2. Verifiez les permissions du dossier %%TEMP%%
    echo          3. Desactivez temporairement l'antivirus
    echo.
    pause
    exit /b 1
)
echo [OK] Dossier: %WORK_DIR%
echo.

:: ============================================================================
:: DOWNLOAD OFFICE DEPLOYMENT TOOL
:: ============================================================================

:download_odt
echo [INFO] Telechargement de Office Deployment Tool...

set "ODT_URL=https://officecdn.microsoft.com/pr/wsus/setup.exe"

powershell -Command ^
    "$ProgressPreference = 'SilentlyContinue';" ^
    "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;" ^
    "$url = '%ODT_URL%';" ^
    "$out = '%ODT_EXE%';" ^
    "try {" ^
    "    $wc = New-Object Net.WebClient;" ^
    "    $wc.Headers.Add('User-Agent', 'Mozilla/5.0');" ^
    "    $total = 0;" ^
    "    $wc.DownloadProgressChanged += {" ^
    "        $pct = $_.ProgressPercentage;" ^
    "        $rcv = [math]::Round($_.BytesReceived/1MB, 2);" ^
    "        $width = $Host.UI.RawUI.WindowSize.Width - 30;" ^
    "        if ($width -lt 10) { $width = 10 };" ^
    "        $done = [math]::Floor($width * $pct / 100);" ^
    "        $left = $width - $done;" ^
    "        $bar = '[' + ('=' * $done) + ('>' * [math]::Min(1, $left)) + (' ' * [math]::Max(0, $left - 1)) + ']';" ^
    "        Write-Host -NoNewline \"`r       $bar $pct%% ($rcv MB)   \";" ^
    "    };" ^
    "    $wc.DownloadFileCompleted += { $global:done = $true };" ^
    "    $global:done = $false;" ^
    "    $wc.DownloadFileAsync([Uri]$url, $out);" ^
    "    while (-not $global:done) { Start-Sleep -Milliseconds 50 };" ^
    "    Write-Host '';" ^
    "} catch { Write-Host \"Erreur: $($_.Exception.Message)\"; exit 1 }"

if not exist "%ODT_EXE%" (
    echo.
    echo [ERREUR] Echec du telechargement de Office Deployment Tool.
    echo.
    echo          CAUSES POSSIBLES:
    echo          - Pas de connexion Internet
    echo          - Serveur Microsoft indisponible
    echo          - Pare-feu/Proxy bloquant la connexion
    echo          - Antivirus bloquant le telechargement
    echo.
    echo          SOLUTIONS:
    echo          1. Verifiez votre connexion Internet
    echo          2. Desactivez temporairement le pare-feu/antivirus
    echo          3. Si vous etes sur un reseau d'entreprise, contactez votre admin
    echo          4. Reessayez plus tard
    echo.
    echo          URL: %ODT_URL%
    echo.
    pause
    exit /b 1
)

echo [OK] Office Deployment Tool telecharge
echo.

:: ============================================================================
:: CREATE CONFIGURATION FILE
:: ============================================================================

:create_config
echo [INFO] Creation du fichier de configuration...

echo ^<Configuration^>> "%CONFIG_FILE%"
echo   ^<Add OfficeClientEdition="%OFFICE_ARCH%" Channel="%OFFICE_CHANNEL%"^>>> "%CONFIG_FILE%"
echo     ^<Product ID="%OFFICE_PRODUCT%"^>>> "%CONFIG_FILE%"
echo       ^<Language ID="%OFFICE_LANG%" /^>>> "%CONFIG_FILE%"
echo       ^<Language ID="MatchOS" /^>>> "%CONFIG_FILE%"
echo       ^<ExcludeApp ID="Publisher" /^>>> "%CONFIG_FILE%"
echo       ^<ExcludeApp ID="Access" /^>>> "%CONFIG_FILE%"
echo       ^<ExcludeApp ID="OneDrive" /^>>> "%CONFIG_FILE%"
echo       ^<ExcludeApp ID="Teams" /^>>> "%CONFIG_FILE%"
echo     ^</Product^>>> "%CONFIG_FILE%"
echo   ^</Add^>>> "%CONFIG_FILE%"
echo   ^<Property Name="SharedComputerLicensing" Value="0" /^>>> "%CONFIG_FILE%"
echo   ^<Property Name="FORCEAPPSHUTDOWN" Value="TRUE" /^>>> "%CONFIG_FILE%"
echo   ^<Property Name="DeviceBasedLicensing" Value="0" /^>>> "%CONFIG_FILE%"
echo   ^<Property Name="SCLCacheOverride" Value="0" /^>>> "%CONFIG_FILE%"
echo   ^<Updates Enabled="TRUE" /^>>> "%CONFIG_FILE%"
echo   ^<RemoveMSI /^>>> "%CONFIG_FILE%"
echo   ^<Display Level="None" AcceptEULA="TRUE" /^>>> "%CONFIG_FILE%"
echo ^</Configuration^>>> "%CONFIG_FILE%"

if not exist "%CONFIG_FILE%" (
    echo [ERREUR] Impossible de creer le fichier de configuration.
    echo.
    echo          CAUSES POSSIBLES:
    echo          - Disque plein
    echo          - Permissions insuffisantes
    echo.
    pause
    exit /b 1
)

echo [OK] Configuration creee
echo.

:: ============================================================================
:: DOWNLOAD OFFICE FILES
:: ============================================================================

:download_office
echo [INFO] Telechargement des fichiers Office...
echo        Cela peut prendre 5-30 minutes selon votre connexion.
echo        Veuillez patienter...
echo.

cd /d "%WORK_DIR%"
"%ODT_EXE%" /download "%CONFIG_FILE%"

if %errorlevel% neq 0 (
    echo.
    echo [ATTENTION] Le telechargement a peut-etre rencontre des problemes.
    echo             Code de retour: %errorlevel%
    echo.
    echo             CAUSES POSSIBLES:
    echo             - Connexion Internet interrompue
    echo             - Espace disque insuffisant (besoin de ~4 Go)
    echo             - Timeout du serveur
    echo.
    echo             Le script va tenter de continuer...
    echo.
)

echo [OK] Telechargement Office termine
echo.

:: ============================================================================
:: INSTALL OFFICE
:: ============================================================================

:install_office
echo [INFO] Installation de Microsoft Office...
echo        Cela peut prendre 5-15 minutes.
echo        NE FERMEZ PAS cette fenetre.
echo.

"%ODT_EXE%" /configure "%CONFIG_FILE%"

if %errorlevel% neq 0 (
    echo.
    echo [ERREUR] L'installation d'Office a echoue.
    echo          Code d'erreur: %errorlevel%
    echo.
    echo          CAUSES POSSIBLES:
    echo          - Une version d'Office est deja installee
    echo          - Fichiers d'installation corrompus
    echo          - Espace disque insuffisant
    echo          - Applications Office ouvertes pendant l'installation
    echo.
    echo          SOLUTIONS:
    echo          1. Fermez toutes les applications Office
    echo          2. Desinstallez les anciennes versions d'Office
    echo          3. Utilisez l'outil de desinstallation Microsoft:
    echo             https://aka.ms/SaRA-OfficeUninstallFromPC
    echo          4. Liberez de l'espace disque (min 4 Go)
    echo          5. Relancez ce script
    echo.
    pause
    exit /b 1
)

echo.
echo [OK] Microsoft Office installe avec succes
echo.

echo [INFO] Finalisation de l'installation...
timeout /t 30 /nobreak >nul

:: ============================================================================
:: ACTIVATE OFFICE USING OHOOK
:: ============================================================================

:activate_office
echo [INFO] Activation d'Office avec Ohook...
echo.

call :download_with_progress "%OHOOK_SCRIPT_URL%" "%TEMP%\Ohook-Activate.cmd" "script Ohook"

if exist "%TEMP%\Ohook-Activate.cmd" (
    echo [OK] Script telecharge
    call "%TEMP%\Ohook-Activate.cmd"
    del /f /q "%TEMP%\Ohook-Activate.cmd" 2>nul
) else (
    echo [ERREUR] Impossible de telecharger le script d'activation.
    echo.
    echo          CAUSES POSSIBLES:
    echo          - GitHub est inaccessible
    echo          - Pare-feu bloquant raw.githubusercontent.com
    echo          - URL incorrecte
    echo.
    echo          SOLUTION MANUELLE:
    echo          1. Telechargez manuellement Ohook-Activate.cmd depuis:
    echo             %OHOOK_SCRIPT_URL%
    echo          2. Executez-le en tant qu'administrateur
    echo.
)

echo.
echo [OK] Processus d'activation termine
echo.

:: ============================================================================
:: VERIFY ACTIVATION
:: ============================================================================

:verify_activation
echo [INFO] Verification du statut d'activation...
echo.

set "OSPP_FOUND=0"
for %%p in (
    "%ProgramFiles%\Microsoft Office\root\Office16\OSPP.VBS"
    "%ProgramFiles(x86)%\Microsoft Office\root\Office16\OSPP.VBS"
    "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS"
    "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS"
) do (
    if exist "%%~p" (
        cscript //nologo "%%~p" /dstatus 2>nul
        set "OSPP_FOUND=1"
        goto :after_verify
    )
)

:after_verify
if "%OSPP_FOUND%"=="0" (
    echo [INFO] Impossible de verifier le statut - OSPP.VBS non trouve.
    echo        Cela peut etre normal si Office vient d'etre installe.
)
echo.

:: ============================================================================
:: CLEANUP
:: ============================================================================

:cleanup
echo [INFO] Nettoyage des fichiers temporaires...
rd /s /q "%WORK_DIR%" 2>nul
echo [OK] Nettoyage termine
echo.

:: ============================================================================
:: DISABLE TELEMETRY
:: ============================================================================

:disable_telemetry
echo [INFO] Desactivation de la telemetrie...
echo.

call :download_with_progress "%TELEMETRY_SCRIPT_URL%" "%TEMP%\Disable-Telemetry.cmd" "script telemetrie"

if exist "%TEMP%\Disable-Telemetry.cmd" (
    echo [OK] Script telecharge
    call "%TEMP%\Disable-Telemetry.cmd"
    del /f /q "%TEMP%\Disable-Telemetry.cmd" 2>nul
) else (
    echo [ATTENTION] Impossible de telecharger le script de telemetrie.
    echo             La telemetrie n'a pas ete desactivee.
    echo             Vous pouvez le faire manuellement plus tard.
)

echo.

:: ============================================================================
:: FINISH
:: ============================================================================

:finish
echo ============================================
echo   Installation et Activation Terminees!
echo ============================================
echo.
echo Produit: %OFFICE_PRODUCT%
echo Architecture: %OFFICE_ARCH%-bit
echo Langue: %OFFICE_LANG%
echo.
echo Si vous rencontrez des problemes:
echo   - Activation: https://massgrave.dev/troubleshoot
echo   - Ohook: https://github.com/asdcorp/ohook
echo.
pause
exit /b 0

:: ============================================================================
:: FUNCTION: Download with progress bar
:: ============================================================================

:download_with_progress
set "DL_URL=%~1"
set "DL_OUT=%~2"
set "DL_NAME=%~3"

echo [INFO] Telechargement du %DL_NAME%...

powershell -Command ^
    "$ProgressPreference = 'SilentlyContinue';" ^
    "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;" ^
    "$url = '%DL_URL%';" ^
    "$out = '%DL_OUT%';" ^
    "try {" ^
    "    $wc = New-Object Net.WebClient;" ^
    "    $wc.DownloadProgressChanged += {" ^
    "        $pct = $_.ProgressPercentage;" ^
    "        $rcv = [math]::Round($_.BytesReceived/1KB, 0);" ^
    "        $width = $Host.UI.RawUI.WindowSize.Width - 25;" ^
    "        if ($width -lt 10) { $width = 10 };" ^
    "        $done = [math]::Floor($width * $pct / 100);" ^
    "        $left = $width - $done;" ^
    "        $bar = '[' + ('=' * $done) + (' ' * $left) + ']';" ^
    "        Write-Host -NoNewline \"`r       $bar $pct%% ($rcv KB)\";" ^
    "    };" ^
    "    $wc.DownloadFileCompleted += { $global:done = $true };" ^
    "    $global:done = $false;" ^
    "    $wc.DownloadFileAsync([Uri]$url, $out);" ^
    "    while (-not $global:done) { Start-Sleep -Milliseconds 50 };" ^
    "    Write-Host '';" ^
    "} catch { Write-Host \"Erreur: $($_.Exception.Message)\" }" 2>nul

goto :eof
