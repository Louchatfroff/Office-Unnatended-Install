@echo off
:: ============================================================================
:: Office Installation Menu - Interactive Version
:: Based on Microsoft Activation Scripts (MAS) - https://massgrave.dev
:: ============================================================================

setlocal EnableDelayedExpansion
title Office Installation Menu

:: Check admin
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Please run this script as Administrator.
    pause
    exit /b 1
)

:menu
cls
echo.
echo  ============================================================
echo      OFFICE UNATTENDED INSTALLATION AND ACTIVATION
echo  ============================================================
echo.
echo      [1] Office 365 ProPlus (Recommended)
echo      [2] Office 2021 LTSC Professional Plus
echo      [3] Office 2019 Professional Plus
echo      [4] Install with custom config file
echo.
echo      [5] Activate Office only (already installed)
echo      [6] Check activation status
echo      [7] Disable telemetry and recommendations
echo.
echo      [0] Exit
echo.
echo  ============================================================
echo.

set /p choice="  Your choice [0-7]: "

if "%choice%"=="1" goto install_365
if "%choice%"=="2" goto install_2021
if "%choice%"=="3" goto install_2019
if "%choice%"=="4" goto install_custom
if "%choice%"=="5" goto activate_only
if "%choice%"=="6" goto check_status
if "%choice%"=="7" goto disable_telemetry
if "%choice%"=="0" goto end

echo  [ERROR] Invalid choice
timeout /t 2 >nul
goto menu

:: ============================================================================
:: OFFICE 365 INSTALLATION
:: ============================================================================
:install_365
cls
echo.
echo [INFO] Installing Office 365 ProPlus...
echo.

set "PRODUCT=O365ProPlusRetail"
set "CHANNEL=Current"
set "PIDKEY="
goto do_install

:: ============================================================================
:: OFFICE 2021 INSTALLATION
:: ============================================================================
:install_2021
cls
echo.
echo [INFO] Installing Office 2021 LTSC Professional Plus...
echo.

set "PRODUCT=ProPlus2021Volume"
set "CHANNEL=PerpetualVL2021"
set "PIDKEY=FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"
goto do_install

:: ============================================================================
:: OFFICE 2019 INSTALLATION
:: ============================================================================
:install_2019
cls
echo.
echo [INFO] Installing Office 2019 Professional Plus...
echo.

set "PRODUCT=ProPlus2019Volume"
set "CHANNEL=PerpetualVL2019"
set "PIDKEY=NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"
goto do_install

:: ============================================================================
:: CUSTOM INSTALLATION
:: ============================================================================
:install_custom
cls
echo.
echo [INFO] Install with custom configuration file
echo.
set /p CONFIG_PATH="  Path to XML file: "

if not exist "%CONFIG_PATH%" (
    echo [ERROR] File not found: %CONFIG_PATH%
    pause
    goto menu
)

set "USE_CUSTOM=1"
goto do_install

:: ============================================================================
:: COMMON INSTALLATION LOGIC
:: ============================================================================
:do_install
set "WORK_DIR=%TEMP%\OfficeInstall_%RANDOM%"
set "ODT_EXE=%WORK_DIR%\setup.exe"
set "CONFIG_FILE=%WORK_DIR%\config.xml"

:: Create temp directory
if not exist "%WORK_DIR%" mkdir "%WORK_DIR%"

:: Download ODT
echo [INFO] Downloading Office Deployment Tool...
powershell -Command ^
    "$ProgressPreference = 'SilentlyContinue';" ^
    "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;" ^
    "$url = 'https://officecdn.microsoft.com/pr/wsus/setup.exe';" ^
    "$out = '%ODT_EXE%';" ^
    "try {" ^
    "    $wc = New-Object Net.WebClient;" ^
    "    $wc.Headers.Add('User-Agent', 'Mozilla/5.0');" ^
    "    $wc.DownloadProgressChanged += {" ^
    "        $pct = $_.ProgressPercentage;" ^
    "        $rcv = [math]::Round($_.BytesReceived/1MB, 2);" ^
    "        $width = $Host.UI.RawUI.WindowSize.Width - 30;" ^
    "        if ($width -lt 10) { $width = 10 };" ^
    "        $done = [math]::Floor($width * $pct / 100);" ^
    "        $left = $width - $done;" ^
    "        $bar = '[' + ('=' * $done) + (' ' * $left) + ']';" ^
    "        Write-Host -NoNewline \"`r       $bar $pct%% ($rcv MB)\";" ^
    "    };" ^
    "    $wc.DownloadFileCompleted += { $global:done = $true };" ^
    "    $global:done = $false;" ^
    "    $wc.DownloadFileAsync([Uri]$url, $out);" ^
    "    while (-not $global:done) { Start-Sleep -Milliseconds 50 };" ^
    "    Write-Host '';" ^
    "} catch { Write-Host \"Error: $($_.Exception.Message)\"; exit 1 }"

if not exist "%ODT_EXE%" (
    echo.
    echo [ERROR] Failed to download ODT
    echo.
    echo         POSSIBLE CAUSES:
    echo         - No Internet connection
    echo         - Microsoft server unavailable
    echo         - Firewall/Proxy blocking connection
    echo.
    echo         SOLUTIONS:
    echo         1. Check your Internet connection
    echo         2. Temporarily disable firewall/antivirus
    echo         3. Try again later
    echo.
    pause
    goto menu
)
echo [OK] ODT downloaded
echo.

:: Create or use config
if defined USE_CUSTOM (
    copy "%CONFIG_PATH%" "%CONFIG_FILE%" >nul
) else (
    call :create_config
)

:: Download Office
echo [INFO] Downloading Office files...
echo        This may take several minutes...
echo.
cd /d "%WORK_DIR%"
"%ODT_EXE%" /download "%CONFIG_FILE%"
echo.
echo [OK] Download complete
echo.

:: Install Office
echo [INFO] Installing Microsoft Office...
echo        Please wait 5-15 minutes...
echo.
"%ODT_EXE%" /configure "%CONFIG_FILE%"
echo.
echo [OK] Installation complete
echo.

:: Wait
timeout /t 20 /nobreak >nul

:: Activate
call :run_ohook
echo.
echo [OK] Activation complete
echo.

:: Disable telemetry
call :run_telemetry_disable
echo.
echo [OK] Telemetry disabled
echo.

:: Cleanup temp directory
echo [INFO] Cleaning up temporary files...
rd /s /q "%WORK_DIR%" 2>nul
echo [OK] Cleanup complete
echo.

pause
goto menu

:: ============================================================================
:: CREATE CONFIG FILE
:: ============================================================================
:create_config
echo [INFO] Creating configuration file...

set "PIDKEY_LINE="
if defined PIDKEY set "PIDKEY_LINE= PIDKEY=\"%PIDKEY%\""

(
echo ^<Configuration^>
echo   ^<Add OfficeClientEdition="64" Channel="%CHANNEL%"^>
echo     ^<Product ID="%PRODUCT%"%PIDKEY_LINE%^>
echo       ^<Language ID="en-us" /^>
echo       ^<Language ID="MatchOS" /^>
echo       ^<ExcludeApp ID="Publisher" /^>
echo       ^<ExcludeApp ID="Access" /^>
echo       ^<ExcludeApp ID="OneDrive" /^>
echo       ^<ExcludeApp ID="Teams" /^>
echo     ^</Product^>
echo   ^</Add^>
echo   ^<Property Name="FORCEAPPSHUTDOWN" Value="TRUE" /^>
echo   ^<Property Name="SharedComputerLicensing" Value="0" /^>
echo   ^<Updates Enabled="TRUE" /^>
echo   ^<RemoveMSI /^>
echo   ^<Display Level="None" AcceptEULA="TRUE" /^>
echo ^</Configuration^>
) > "%CONFIG_FILE%"

echo [OK] Configuration created
goto :eof

:: ============================================================================
:: ACTIVATE ONLY
:: ============================================================================
:activate_only
cls
echo.
echo [INFO] Activating Office with Ohook...
echo.

call :run_ohook

echo.
echo [OK] Activation complete
echo.
pause
goto menu

:: ============================================================================
:: RUN OHOOK FROM WEB
:: ============================================================================
:run_ohook
echo [INFO] Downloading Ohook script...

set "OHOOK_SCRIPT_URL=https://raw.githubusercontent.com/Louchatfroff/Office-Unnatended-Install/main/Ohook-Activate.cmd"

powershell -Command ^
    "$ProgressPreference = 'SilentlyContinue';" ^
    "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;" ^
    "$url = '%OHOOK_SCRIPT_URL%';" ^
    "$out = '%TEMP%\Ohook-Activate.cmd';" ^
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

if exist "%TEMP%\Ohook-Activate.cmd" (
    echo [OK] Script downloaded
    call "%TEMP%\Ohook-Activate.cmd"
    del /f /q "%TEMP%\Ohook-Activate.cmd" 2>nul
) else (
    echo.
    echo [ERROR] Failed to download Ohook script
    echo.
    echo         POSSIBLE CAUSES:
    echo         - GitHub is unreachable
    echo         - Firewall blocking raw.githubusercontent.com
    echo.
    echo         SOLUTIONS:
    echo         1. Check your Internet connection
    echo         2. Try accessing github.com
    echo         3. Temporarily disable firewall
    echo.
    echo         URL: %OHOOK_SCRIPT_URL%
)
goto :eof

:: ============================================================================
:: DISABLE TELEMETRY (menu option)
:: ============================================================================
:disable_telemetry
cls
echo.
echo [INFO] Disabling telemetry and recommendations...
echo.

call :run_telemetry_disable

echo.
echo [OK] Telemetry and recommendations disabled
echo.
pause
goto menu

:: ============================================================================
:: RUN TELEMETRY DISABLE FROM WEB
:: ============================================================================
:run_telemetry_disable
echo [INFO] Downloading disable script...

set "TELEMETRY_SCRIPT_URL=https://raw.githubusercontent.com/Louchatfroff/Office-Unnatended-Install/main/Disable-Telemetry.cmd"

powershell -Command ^
    "$ProgressPreference = 'SilentlyContinue';" ^
    "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;" ^
    "$url = '%TELEMETRY_SCRIPT_URL%';" ^
    "$out = '%TEMP%\Disable-Telemetry.cmd';" ^
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

if exist "%TEMP%\Disable-Telemetry.cmd" (
    echo [OK] Script downloaded
    call "%TEMP%\Disable-Telemetry.cmd"
    del /f /q "%TEMP%\Disable-Telemetry.cmd" 2>nul
) else (
    echo.
    echo [ERROR] Failed to download script
    echo.
    echo         POSSIBLE CAUSES:
    echo         - GitHub is unreachable
    echo         - Firewall blocking raw.githubusercontent.com
    echo.
    echo         SOLUTIONS:
    echo         1. Check your Internet connection
    echo         2. Temporarily disable firewall
    echo.
    echo         URL: %TELEMETRY_SCRIPT_URL%
)
goto :eof

:: ============================================================================
:: CHECK STATUS
:: ============================================================================
:check_status
cls
echo.
echo [INFO] Checking Office activation status...
echo.
echo ============================================================

:: Try both paths
set "OSPP_FOUND=0"
if exist "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" (
    cscript //nologo "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" /dstatus
    set "OSPP_FOUND=1"
)
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS" (
    cscript //nologo "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS" /dstatus
    set "OSPP_FOUND=1"
)

if "%OSPP_FOUND%"=="0" (
    echo [WARNING] Office does not appear to be installed or OSPP.VBS not found
)

echo.
echo ============================================================
echo.
pause
goto menu

:: ============================================================================
:: END
:: ============================================================================
:end
echo.
echo  Goodbye!
echo.
exit /b 0
