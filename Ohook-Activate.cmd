@echo off
:: ============================================================================
:: Ohook Office Activation Script - Standalone Version
:: Based on ohook by asdcorp - https://github.com/asdcorp/ohook
:: ============================================================================
:: This script activates Microsoft Office using the Ohook method
:: It places a custom sppc.dll file in the Office folder
:: ============================================================================

setlocal EnableDelayedExpansion
title Ohook Office Activation

:: ============================================================================
:: CONFIGURATION - DLL URLs (modify if needed)
:: ============================================================================
set "OHOOK_VERSION=0.5"
set "OHOOK_URL_BASE=https://github.com/asdcorp/ohook/releases/download/%OHOOK_VERSION%"
set "DLL32_URL=%OHOOK_URL_BASE%/sppc32.dll"
set "DLL64_URL=%OHOOK_URL_BASE%/sppc64.dll"

:: SHA256 checksums for verification
set "DLL32_HASH=09865ea5993215965e8f27a74b8a41d15fd0f60f5f404cb7a8b3c7757acdab02"
set "DLL64_HASH=393a1fa26deb3663854e41f2b687c188a9eacd87b23f17ea09422c4715cb5a9f"

:: Temp directory
set "TEMP_DIR=%TEMP%\ohook_temp"
set "DLL32_PATH=%TEMP_DIR%\sppc32.dll"
set "DLL64_PATH=%TEMP_DIR%\sppc64.dll"

:: ============================================================================
:: ADMIN CHECK
:: ============================================================================
:check_admin
echo.
echo ============================================
echo   Ohook Office Activation Script
echo   Version: %OHOOK_VERSION%
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
:: OFFICE DETECTION
:: ============================================================================
:detect_office
echo [INFO] Detecting Office installations...
echo.

set "OFFICE_FOUND=0"
set "OFFICE_PATHS="

:: Office Click-to-Run paths
for %%a in (
    "%ProgramFiles%\Microsoft Office\root\Office16"
    "%ProgramFiles(x86)%\Microsoft Office\root\Office16"
    "%ProgramFiles%\Microsoft Office\root\Office15"
    "%ProgramFiles(x86)%\Microsoft Office\root\Office15"
) do (
    if exist "%%~a\OSPP.VBS" (
        echo [FOUND] %%~a
        set "OFFICE_FOUND=1"
        set "OFFICE_PATHS=!OFFICE_PATHS!%%~a;"
    )
)

:: Office MSI paths
for %%a in (
    "%ProgramFiles%\Microsoft Office\Office16"
    "%ProgramFiles(x86)%\Microsoft Office\Office16"
    "%ProgramFiles%\Microsoft Office\Office15"
    "%ProgramFiles(x86)%\Microsoft Office\Office15"
) do (
    if exist "%%~a\OSPP.VBS" (
        echo [FOUND] %%~a
        set "OFFICE_FOUND=1"
        set "OFFICE_PATHS=!OFFICE_PATHS!%%~a;"
    )
)

:: Check registry for additional paths
for /f "tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v InstallationPath 2^>nul') do (
    set "C2R_PATH=%%b"
    if defined C2R_PATH (
        for %%v in (Office16 Office15) do (
            if exist "!C2R_PATH!\root\%%v\OSPP.VBS" (
                echo [FOUND] !C2R_PATH!\root\%%v
                set "OFFICE_FOUND=1"
                set "OFFICE_PATHS=!OFFICE_PATHS!!C2R_PATH!\root\%%v;"
            )
        )
    )
)

echo.
if "%OFFICE_FOUND%"=="0" (
    echo [ERROR] No Office installation detected.
    echo.
    echo         POSSIBLE CAUSES:
    echo         - Office is not installed on this computer
    echo         - Office is installed in a non-standard location
    echo         - Office installation is corrupted
    echo.
    echo         SOLUTIONS:
    echo         1. Install Office using Office-Install-Activate.cmd
    echo         2. Download Office from office.com
    echo         3. If Office is installed, try to repair it:
    echo            Settings ^> Apps ^> Microsoft Office ^> Modify ^> Repair
    echo.
    pause
    exit /b 1
)

:: ============================================================================
:: DOWNLOAD DLLs
:: ============================================================================
:download_dlls
echo [INFO] Downloading Ohook files...
echo.

if not exist "%TEMP_DIR%" mkdir "%TEMP_DIR%"

:: Download 64-bit DLL
echo [INFO] Downloading sppc64.dll...
powershell -Command ^
    "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;" ^
    "$ProgressPreference = 'Continue';" ^
    "try {" ^
    "    Invoke-WebRequest -Uri '%DLL64_URL%' -OutFile '%DLL64_PATH%' -UseBasicParsing;" ^
    "} catch {" ^
    "    Write-Host \"Error: $($_.Exception.Message)\";" ^
    "    exit 1;" ^
    "}"

if not exist "%DLL64_PATH%" (
    echo.
    echo [ERROR] Failed to download sppc64.dll
    echo.
    echo         POSSIBLE CAUSES:
    echo         - No Internet connection
    echo         - GitHub is inaccessible or blocked
    echo         - Firewall/Proxy blocking connection
    echo         - Antivirus blocking download
    echo.
    echo         SOLUTIONS:
    echo         1. Check your Internet connection
    echo         2. Try accessing github.com in a browser
    echo         3. Temporarily disable firewall/antivirus
    echo         4. If on corporate network, contact admin
    echo.
    echo         URL: %DLL64_URL%
    echo.
    pause
    exit /b 1
)
echo [OK] sppc64.dll downloaded

:: Download 32-bit DLL
echo [INFO] Downloading sppc32.dll...
powershell -Command ^
    "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;" ^
    "$ProgressPreference = 'Continue';" ^
    "try {" ^
    "    Invoke-WebRequest -Uri '%DLL32_URL%' -OutFile '%DLL32_PATH%' -UseBasicParsing;" ^
    "} catch {" ^
    "    Write-Host \"Error: $($_.Exception.Message)\";" ^
    "    exit 1;" ^
    "}"

if not exist "%DLL32_PATH%" (
    echo.
    echo [ERROR] Failed to download sppc32.dll
    echo.
    echo         POSSIBLE CAUSES:
    echo         - No Internet connection
    echo         - GitHub is inaccessible or blocked
    echo         - Firewall/Proxy blocking connection
    echo         - Antivirus blocking download
    echo.
    echo         SOLUTIONS:
    echo         1. Check your Internet connection
    echo         2. Try accessing github.com in a browser
    echo         3. Temporarily disable firewall/antivirus
    echo         4. If on corporate network, contact admin
    echo.
    echo         URL: %DLL32_URL%
    echo.
    pause
    exit /b 1
)
echo [OK] sppc32.dll downloaded
echo.

:: ============================================================================
:: HASH VERIFICATION (optional but recommended)
:: ============================================================================
:verify_hash
echo [INFO] Verifying file integrity...

for /f "skip=1 tokens=* delims=" %%a in ('certutil -hashfile "%DLL64_PATH%" SHA256 2^>nul') do (
    set "COMPUTED_HASH=%%a"
    goto :check_hash64
)
:check_hash64
set "COMPUTED_HASH=%COMPUTED_HASH: =%"
if /i "%COMPUTED_HASH%"=="%DLL64_HASH%" (
    echo [OK] sppc64.dll - Hash valid
) else (
    echo [WARNING] sppc64.dll - Different hash (may be a newer version^)
)

for /f "skip=1 tokens=* delims=" %%a in ('certutil -hashfile "%DLL32_PATH%" SHA256 2^>nul') do (
    set "COMPUTED_HASH=%%a"
    goto :check_hash32
)
:check_hash32
set "COMPUTED_HASH=%COMPUTED_HASH: =%"
if /i "%COMPUTED_HASH%"=="%DLL32_HASH%" (
    echo [OK] sppc32.dll - Hash valid
) else (
    echo [WARNING] sppc32.dll - Different hash (may be a newer version^)
)
echo.

:: ============================================================================
:: DLL INSTALLATION
:: ============================================================================
:install_dlls
echo [INFO] Installing Ohook files...
echo.

:: Process each Office path
for %%p in (%OFFICE_PATHS%) do (
    set "CURRENT_PATH=%%~p"
    if not "!CURRENT_PATH!"=="" (
        echo [INFO] Processing: !CURRENT_PATH!

        :: Determine architecture
        echo !CURRENT_PATH! | findstr /i "x86" >nul
        if !errorlevel! equ 0 (
            set "DLL_SOURCE=%DLL32_PATH%"
            echo        Architecture: 32-bit
        ) else (
            echo !CURRENT_PATH! | findstr /i "Program Files (x86)" >nul
            if !errorlevel! equ 0 (
                set "DLL_SOURCE=%DLL32_PATH%"
                echo        Architecture: 32-bit
            ) else (
                set "DLL_SOURCE=%DLL64_PATH%"
                echo        Architecture: 64-bit
            )
        )

        :: Copy DLL
        copy /y "!DLL_SOURCE!" "!CURRENT_PATH!\sppc.dll" >nul 2>&1
        if !errorlevel! equ 0 (
            echo        [OK] sppc.dll installed
        ) else (
            echo        [ERROR] Failed to copy sppc.dll
            echo                Close all Office applications and try again
        )
        echo.
    )
)

:: ============================================================================
:: LICENSE INSTALLATION
:: ============================================================================
:install_licenses
echo [INFO] Installing Office licenses...
echo.

:: Find Office C2R license path
set "LICENSE_PATH="
for /f "tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v InstallationPath 2^>nul') do (
    set "LICENSE_PATH=%%b\root\Licenses16"
)

if not defined LICENSE_PATH (
    :: Try common paths
    if exist "%ProgramFiles%\Microsoft Office\root\Licenses16" (
        set "LICENSE_PATH=%ProgramFiles%\Microsoft Office\root\Licenses16"
    ) else if exist "%ProgramFiles(x86)%\Microsoft Office\root\Licenses16" (
        set "LICENSE_PATH=%ProgramFiles(x86)%\Microsoft Office\root\Licenses16"
    )
)

if defined LICENSE_PATH (
    if exist "!LICENSE_PATH!" (
        echo [INFO] License folder: !LICENSE_PATH!

        :: Install Grace licenses for all products
        for %%p in (
            "%ProgramFiles%\Microsoft Office\root\Office16"
            "%ProgramFiles(x86)%\Microsoft Office\root\Office16"
        ) do (
            if exist "%%~p\OSPP.VBS" (
                echo [INFO] Installing licenses via OSPP.VBS...

                :: Install volume licenses
                for /f "delims=" %%l in ('dir /b "!LICENSE_PATH!\*VL_*.xrm-ms" 2^>nul') do (
                    cscript //nologo "%%~p\OSPP.VBS" /inslic:"!LICENSE_PATH!\%%l" >nul 2>&1
                )

                echo [OK] Licenses installed
            )
        )
    )
)
echo.

:: ============================================================================
:: ACTIVATION
:: ============================================================================
:activate
echo [INFO] Activating Office...
echo.

:: Try to activate via OSPP
for %%p in (
    "%ProgramFiles%\Microsoft Office\root\Office16"
    "%ProgramFiles(x86)%\Microsoft Office\root\Office16"
    "%ProgramFiles%\Microsoft Office\Office16"
    "%ProgramFiles(x86)%\Microsoft Office\Office16"
) do (
    if exist "%%~p\OSPP.VBS" (
        echo [INFO] Using: %%~p\OSPP.VBS

        :: Set KMS host (local activation via Ohook)
        cscript //nologo "%%~p\OSPP.VBS" /sethst:127.0.0.1 >nul 2>&1

        :: Activate
        cscript //nologo "%%~p\OSPP.VBS" /act >nul 2>&1

        goto :verify
    )
)

:: ============================================================================
:: VERIFICATION
:: ============================================================================
:verify
echo.
echo [INFO] Verifying activation status...
echo.
echo ============================================

for %%p in (
    "%ProgramFiles%\Microsoft Office\root\Office16"
    "%ProgramFiles(x86)%\Microsoft Office\root\Office16"
    "%ProgramFiles%\Microsoft Office\Office16"
    "%ProgramFiles(x86)%\Microsoft Office\Office16"
) do (
    if exist "%%~p\OSPP.VBS" (
        cscript //nologo "%%~p\OSPP.VBS" /dstatus
        goto :cleanup
    )
)

:: ============================================================================
:: CLEANUP
:: ============================================================================
:cleanup
echo.
echo ============================================
echo.
echo [INFO] Cleaning up temporary files...
rd /s /q "%TEMP_DIR%" 2>nul
echo [OK] Cleanup complete
echo.

:: ============================================================================
:: END
:: ============================================================================
:end
echo ============================================
echo   Ohook Activation Complete!
echo ============================================
echo.
echo If activation failed, try:
echo   1. Close all Office applications
echo   2. Re-run this script as administrator
echo   3. Restart your computer and try again
echo.
echo Source: https://github.com/asdcorp/ohook
echo.
pause
exit /b 0
