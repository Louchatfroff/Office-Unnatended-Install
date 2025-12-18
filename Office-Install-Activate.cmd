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
set "OFFICE_LANG=en-us"
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
    echo [ERROR] This script requires Administrator privileges.
    echo.
    echo         SOLUTION:
    echo         1. Right-click on the script
    echo         2. Select "Run as administrator"
    echo.
    echo         OR open PowerShell as admin and run:
    echo         irm https://office-unnatended.vercel.app ^| iex
    echo.
    pause
    exit /b 1
)

echo [OK] Administrator privileges confirmed
echo.

:: ============================================================================
:: CREATE TEMP DIRECTORY
:: ============================================================================

:create_dirs
echo [INFO] Creating temp directory...
if not exist "%WORK_DIR%" mkdir "%WORK_DIR%"
if not exist "%WORK_DIR%" (
    echo [ERROR] Failed to create temp directory.
    echo.
    echo         POSSIBLE CAUSES:
    echo         - Disk full
    echo         - Insufficient permissions on %%TEMP%%
    echo         - Antivirus blocking creation
    echo.
    echo         SOLUTIONS:
    echo         1. Free up disk space
    echo         2. Check permissions on %%TEMP%% folder
    echo         3. Temporarily disable antivirus
    echo.
    pause
    exit /b 1
)
echo [OK] Directory: %WORK_DIR%
echo.

:: ============================================================================
:: DOWNLOAD OFFICE DEPLOYMENT TOOL
:: ============================================================================

:download_odt
echo [INFO] Downloading Office Deployment Tool...

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
    "} catch { Write-Host \"Error: $($_.Exception.Message)\"; exit 1 }"

if not exist "%ODT_EXE%" (
    echo.
    echo [ERROR] Failed to download Office Deployment Tool.
    echo.
    echo         POSSIBLE CAUSES:
    echo         - No Internet connection
    echo         - Microsoft server unavailable
    echo         - Firewall/Proxy blocking connection
    echo         - Antivirus blocking download
    echo.
    echo         SOLUTIONS:
    echo         1. Check your Internet connection
    echo         2. Temporarily disable firewall/antivirus
    echo         3. If on corporate network, contact your admin
    echo         4. Try again later
    echo.
    echo         URL: %ODT_URL%
    echo.
    pause
    exit /b 1
)

echo [OK] Office Deployment Tool downloaded
echo.

:: ============================================================================
:: CREATE CONFIGURATION FILE
:: ============================================================================

:create_config
echo [INFO] Creating configuration file...

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
    echo [ERROR] Failed to create configuration file.
    echo.
    echo         POSSIBLE CAUSES:
    echo         - Disk full
    echo         - Insufficient permissions
    echo.
    pause
    exit /b 1
)

echo [OK] Configuration created
echo.

:: ============================================================================
:: DOWNLOAD OFFICE FILES
:: ============================================================================

:download_office
echo [INFO] Downloading Office files...
echo        This may take 5-30 minutes depending on your connection.
echo        Please wait...
echo.

cd /d "%WORK_DIR%"
"%ODT_EXE%" /download "%CONFIG_FILE%"

if %errorlevel% neq 0 (
    echo.
    echo [WARNING] Download may have encountered issues.
    echo           Return code: %errorlevel%
    echo.
    echo           POSSIBLE CAUSES:
    echo           - Internet connection interrupted
    echo           - Insufficient disk space (need ~4 GB)
    echo           - Server timeout
    echo.
    echo           Script will attempt to continue...
    echo.
)

echo [OK] Office download complete
echo.

:: ============================================================================
:: INSTALL OFFICE
:: ============================================================================

:install_office
echo [INFO] Installing Microsoft Office...
echo        This may take 5-15 minutes.
echo        DO NOT CLOSE this window.
echo.

"%ODT_EXE%" /configure "%CONFIG_FILE%"

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Office installation failed.
    echo         Error code: %errorlevel%
    echo.
    echo         POSSIBLE CAUSES:
    echo         - An Office version is already installed
    echo         - Corrupted installation files
    echo         - Insufficient disk space
    echo         - Office applications open during installation
    echo.
    echo         SOLUTIONS:
    echo         1. Close all Office applications
    echo         2. Uninstall previous Office versions
    echo         3. Use Microsoft uninstall tool:
    echo            https://aka.ms/SaRA-OfficeUninstallFromPC
    echo         4. Free up disk space (min 4 GB)
    echo         5. Re-run this script
    echo.
    pause
    exit /b 1
)

echo.
echo [OK] Microsoft Office installed successfully
echo.

echo [INFO] Finalizing installation...
timeout /t 30 /nobreak >nul

:: ============================================================================
:: ACTIVATE OFFICE USING OHOOK
:: ============================================================================

:activate_office
echo [INFO] Activating Office with Ohook...
echo.

call :download_with_progress "%OHOOK_SCRIPT_URL%" "%TEMP%\Ohook-Activate.cmd" "Ohook script"

if exist "%TEMP%\Ohook-Activate.cmd" (
    echo [OK] Script downloaded
    call "%TEMP%\Ohook-Activate.cmd"
    del /f /q "%TEMP%\Ohook-Activate.cmd" 2>nul
) else (
    echo [ERROR] Failed to download activation script.
    echo.
    echo         POSSIBLE CAUSES:
    echo         - GitHub is inaccessible
    echo         - Firewall blocking raw.githubusercontent.com
    echo         - Incorrect URL
    echo.
    echo         MANUAL SOLUTION:
    echo         1. Manually download Ohook-Activate.cmd from:
    echo            %OHOOK_SCRIPT_URL%
    echo         2. Run it as administrator
    echo.
)

echo.
echo [OK] Activation process complete
echo.

:: ============================================================================
:: VERIFY ACTIVATION
:: ============================================================================

:verify_activation
echo [INFO] Verifying activation status...
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
    echo [INFO] Cannot verify status - OSPP.VBS not found.
    echo        This may be normal if Office was just installed.
)
echo.

:: ============================================================================
:: CLEANUP
:: ============================================================================

:cleanup
echo [INFO] Cleaning up temporary files...
rd /s /q "%WORK_DIR%" 2>nul
echo [OK] Cleanup complete
echo.

:: ============================================================================
:: DISABLE TELEMETRY
:: ============================================================================

:disable_telemetry
echo [INFO] Disabling telemetry...
echo.

call :download_with_progress "%TELEMETRY_SCRIPT_URL%" "%TEMP%\Disable-Telemetry.cmd" "telemetry script"

if exist "%TEMP%\Disable-Telemetry.cmd" (
    echo [OK] Script downloaded
    call "%TEMP%\Disable-Telemetry.cmd"
    del /f /q "%TEMP%\Disable-Telemetry.cmd" 2>nul
) else (
    echo [WARNING] Failed to download telemetry script.
    echo           Telemetry was not disabled.
    echo           You can do this manually later.
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
echo Product: %OFFICE_PRODUCT%
echo Architecture: %OFFICE_ARCH%-bit
echo Language: %OFFICE_LANG%
echo.
echo If you encounter any issues:
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

echo [INFO] Downloading %DL_NAME%...

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
    "} catch { Write-Host \"Error: $($_.Exception.Message)\" }" 2>nul

goto :eof
