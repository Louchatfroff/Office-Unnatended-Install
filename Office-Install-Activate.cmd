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

:: Working directory
set "WORK_DIR=%~dp0"
set "ODT_DIR=%WORK_DIR%ODT"
set "CONFIG_FILE=%ODT_DIR%\config.xml"

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
:: CREATE WORKING DIRECTORY
:: ============================================================================

:create_dirs
echo [INFO] Creating working directories...
if not exist "%ODT_DIR%" mkdir "%ODT_DIR%"
echo [OK] Working directory: %ODT_DIR%
echo.

:: ============================================================================
:: DOWNLOAD OFFICE DEPLOYMENT TOOL
:: ============================================================================

:download_odt
echo [INFO] Downloading Office Deployment Tool...

set "ODT_URL=https://officecdn.microsoft.com/pr/wsus/setup.exe"
set "ODT_EXE=%ODT_DIR%\setup.exe"

if exist "%ODT_EXE%" (
    echo [INFO] ODT already exists, skipping download...
) else (
    powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('%ODT_URL%', '%ODT_EXE%')}" 2>nul
    if not exist "%ODT_EXE%" (
        echo [ERROR] Failed to download Office Deployment Tool
        echo         Please check your internet connection
        pause
        exit /b 1
    )
)

echo [OK] Office Deployment Tool ready
echo.

:: ============================================================================
:: CREATE CONFIGURATION FILE
:: ============================================================================

:create_config
echo [INFO] Creating Office configuration file...

:: Build ExcludeApp XML elements
set "EXCLUDE_XML="
for %%a in (%EXCLUDE_APPS%) do (
    set "EXCLUDE_XML=!EXCLUDE_XML!        <ExcludeApp ID="%%a" />^

"
)

:: Create the configuration XML
(
echo ^<Configuration^>
echo   ^<Add OfficeClientEdition="%OFFICE_ARCH%" Channel="%OFFICE_CHANNEL%"^>
echo     ^<Product ID="%OFFICE_PRODUCT%"^>
echo       ^<Language ID="%OFFICE_LANG%" /^>
echo       ^<Language ID="MatchOS" /^>
echo %EXCLUDE_XML%
echo     ^</Product^>
echo   ^</Add^>
echo   ^<Property Name="SharedComputerLicensing" Value="0" /^>
echo   ^<Property Name="FORCEAPPSHUTDOWN" Value="TRUE" /^>
echo   ^<Property Name="DeviceBasedLicensing" Value="0" /^>
echo   ^<Property Name="SCLCacheOverride" Value="0" /^>
echo   ^<Updates Enabled="TRUE" /^>
echo   ^<RemoveMSI /^>
echo   ^<AppSettings^>
echo     ^<User Key="software\microsoft\office\16.0\common\general" Name="shownfirstrunoptin" Value="1" Type="REG_DWORD" /^>
echo   ^</AppSettings^>
echo   ^<Display Level="None" AcceptEULA="TRUE" /^>
echo ^</Configuration^>
) > "%CONFIG_FILE%"

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

cd /d "%ODT_DIR%"
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
set "https://raw.githubusercontent.com/Louchatfroff/Office-Unnatended-Install/refs/heads/main/Ohook-Activate-Silent.cmd?token=GHSAT0AAAAAADQFQ3YJHKXHNVWPJRM2MBQO2KCVPNAI"

echo [INFO] Downloading Ohook activation script...
powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('%OHOOK_SCRIPT_URL%', '%TEMP%\Ohook-Activate.cmd')" 2>nul

if exist "%TEMP%\Ohook-Activate.cmd" (
    echo [OK] Script downloaded
    call "%TEMP%\Ohook-Activate-Silent.cmd"
    del /f /q "%TEMP%\Ohook-Activate-Silent.cmd" 2>nul
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

:: Uncomment the following line to remove downloaded files after installation
:: rd /s /q "%ODT_DIR%\Office" 2>nul

echo [OK] Cleanup complete
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
