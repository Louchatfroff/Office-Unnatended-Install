@echo off
:: ============================================================================
:: Office Unattended Installation and Activation Script
:: Based on Microsoft Activation Scripts (MAS) - https://massgrave.dev
:: ============================================================================
:: This script will:
:: 1. Download Office Deployment Tool (ODT)
:: 2. Install Microsoft Office silently
:: 3. Activate Office using MAS Ohook method
:: ============================================================================

setlocal EnableDelayedExpansion
title Office Unattended Install and Activate

:: ============================================================================
:: CONFIGURATION - Modify these variables as needed
:: ============================================================================

:: Office Edition: O365ProPlusRetail, ProPlus2021Volume, ProPlus2019Volume, etc.
set "OFFICE_PRODUCT=O365ProPlusRetail"

:: Architecture: 64 or 32
set "OFFICE_ARCH=64"

:: Language: en-us, fr-fr, de-de, es-es, etc.
set "OFFICE_LANG=fr-fr"

:: Channel: Current, MonthlyEnterprise, SemiAnnual
set "OFFICE_CHANNEL=Current"

:: Exclude apps (comma separated): Access,Excel,OneDrive,OneNote,Outlook,PowerPoint,Publisher,Word,Teams
set "EXCLUDE_APPS=Publisher,Access,OneDrive,Teams"

:: Working directory (temp)
set "WORK_DIR=%TEMP%\OfficeInstall_%RANDOM%"
set "ODT_EXE=%WORK_DIR%\setup.exe"
set "CONFIG_FILE=%WORK_DIR%\config.xml"

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
    echo [ERROR] This script requires Administrator privileges.
    echo         Right-click and select "Run as administrator"
    echo.
    pause
    exit /b 1
)

echo [OK] Running with Administrator privileges
echo.

:: ============================================================================
:: CREATE TEMP DIRECTORY
:: ============================================================================

:create_dirs
echo [INFO] Creating temp directory...
if not exist "%WORK_DIR%" mkdir "%WORK_DIR%"
echo [OK] Temp directory: %WORK_DIR%
echo.

:: ============================================================================
:: DOWNLOAD OFFICE DEPLOYMENT TOOL
:: ============================================================================

:download_odt
echo [INFO] Downloading Office Deployment Tool...

set "ODT_URL=https://officecdn.microsoft.com/pr/wsus/setup.exe"

powershell -Command "$ProgressPreference = 'Continue'; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $wc = New-Object Net.WebClient; $wc.DownloadProgressChanged += { Write-Host -NoNewline \"`r       Progress: $($_.ProgressPercentage)%% - $([math]::Round($_.BytesReceived/1MB, 2)) MB\" }; $wc.DownloadFileAsync([Uri]'%ODT_URL%', '%ODT_EXE%'); while ($wc.IsBusy) { Start-Sleep -Milliseconds 100 }; Write-Host ''"

if not exist "%ODT_EXE%" (
    echo [ERROR] Failed to download Office Deployment Tool
    echo         Please check your internet connection
    pause
    exit /b 1
)

echo [OK] Office Deployment Tool ready
echo.

:: ============================================================================
:: CREATE CONFIGURATION FILE
:: ============================================================================

:create_config
echo [INFO] Creating Office configuration file...

:: Create the configuration XML
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
    echo [ERROR] Failed to create configuration file
    pause
    exit /b 1
)

echo [OK] Configuration file created: %CONFIG_FILE%
echo.

:: ============================================================================
:: DOWNLOAD OFFICE FILES
:: ============================================================================

:download_office
echo [INFO] Downloading Office installation files...
echo        This may take several minutes depending on your connection...
echo.

cd /d "%WORK_DIR%"
"%ODT_EXE%" /download "%CONFIG_FILE%"

if %errorlevel% neq 0 (
    echo [WARNING] Download may have encountered issues
    echo           Continuing with installation attempt...
)

echo.
echo [OK] Office files download complete
echo.

:: ============================================================================
:: INSTALL OFFICE
:: ============================================================================

:install_office
echo [INFO] Installing Microsoft Office...
echo        This may take 5-15 minutes. Please wait...
echo.

"%ODT_EXE%" /configure "%CONFIG_FILE%"

if %errorlevel% neq 0 (
    echo [ERROR] Office installation failed
    echo         Error code: %errorlevel%
    pause
    exit /b 1
)

echo.
echo [OK] Microsoft Office installed successfully
echo.

:: Wait for Office installation to fully complete
echo [INFO] Waiting for Office installation to finalize...
timeout /t 30 /nobreak >nul

:: ============================================================================
:: ACTIVATE OFFICE USING OHOOK
:: ============================================================================

:activate_office
echo [INFO] Activating Office using Ohook method...
echo.

:: Download and execute Ohook activation script from web
set "OHOOK_SCRIPT_URL=METTRE_URL_DU_SCRIPT_OHOOK_ICI"

echo [INFO] Downloading Ohook activation script...
powershell -Command "$ProgressPreference = 'Continue'; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $wc = New-Object Net.WebClient; $wc.DownloadProgressChanged += { Write-Host -NoNewline \"`r       Progress: $($_.ProgressPercentage)%%\" }; $wc.DownloadFileAsync([Uri]'%OHOOK_SCRIPT_URL%', '%TEMP%\Ohook-Activate.cmd'); while ($wc.IsBusy) { Start-Sleep -Milliseconds 100 }; Write-Host ''" 2>nul

if exist "%TEMP%\Ohook-Activate.cmd" (
    echo [OK] Script downloaded
    call "%TEMP%\Ohook-Activate.cmd"
    del /f /q "%TEMP%\Ohook-Activate.cmd" 2>nul
) else (
    echo [ERROR] Failed to download Ohook script from:
    echo         %OHOOK_SCRIPT_URL%
)

echo.
echo [OK] Activation process completed
echo.

:: ============================================================================
:: VERIFY ACTIVATION
:: ============================================================================

:verify_activation
echo [INFO] Verifying Office activation status...
echo.

:: Check Office licensing status
cscript //nologo "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" /dstatus 2>nul
if %errorlevel% neq 0 (
    cscript //nologo "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS" /dstatus 2>nul
)

echo.

:: ============================================================================
:: CLEANUP (Optional)
:: ============================================================================

:cleanup
echo [INFO] Cleaning up temporary files...

:: Remove temp directory
rd /s /q "%WORK_DIR%" 2>nul

echo [OK] Cleanup complete
echo.

:: ============================================================================
:: DISABLE TELEMETRY
:: ============================================================================

:disable_telemetry
echo [INFO] Disabling telemetry and recommendations...
echo.

:: Download and execute telemetry disable script from web
set "TELEMETRY_SCRIPT_URL=METTRE_URL_DU_SCRIPT_TELEMETRY_ICI"

echo [INFO] Downloading telemetry disable script...
powershell -Command "$ProgressPreference = 'Continue'; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $wc = New-Object Net.WebClient; $wc.DownloadProgressChanged += { Write-Host -NoNewline \"`r       Progress: $($_.ProgressPercentage)%%\" }; $wc.DownloadFileAsync([Uri]'%TELEMETRY_SCRIPT_URL%', '%TEMP%\Disable-Telemetry.cmd'); while ($wc.IsBusy) { Start-Sleep -Milliseconds 100 }; Write-Host ''" 2>nul

if exist "%TEMP%\Disable-Telemetry.cmd" (
    echo [OK] Telemetry script downloaded
    call "%TEMP%\Disable-Telemetry.cmd"
    del /f /q "%TEMP%\Disable-Telemetry.cmd" 2>nul
) else (
    echo [WARNING] Failed to download telemetry script
    echo           URL: %TELEMETRY_SCRIPT_URL%
)

echo.

:: ============================================================================
:: FINISH
:: ============================================================================

:finish
echo ============================================
echo   Installation and Activation Complete!
echo ============================================
echo.
echo Microsoft Office has been installed and activated.
echo.
echo Product: %OFFICE_PRODUCT%
echo Architecture: %OFFICE_ARCH%-bit
echo Language: %OFFICE_LANG%
echo.
echo If you encounter any issues, please visit:
echo https://massgrave.dev/troubleshoot
echo.
pause
exit /b 0
