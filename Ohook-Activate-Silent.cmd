@echo off
:: ============================================================================
:: Ohook Office Activation - Silent/Unattended Version
:: Based on ohook by asdcorp - https://github.com/asdcorp/ohook
:: ============================================================================
:: Usage: Ohook-Activate-Silent.cmd [/log]
::   /log  - Show messages (otherwise completely silent)
:: ============================================================================

setlocal EnableDelayedExpansion

:: Check for /log parameter
set "SILENT=1"
if /i "%~1"=="/log" set "SILENT=0"
if /i "%~1"=="-log" set "SILENT=0"

:: ============================================================================
:: CONFIGURATION
:: ============================================================================
set "OHOOK_VERSION=0.5"
set "DLL64_URL=https://github.com/asdcorp/ohook/releases/download/%OHOOK_VERSION%/sppc64.dll"
set "DLL32_URL=https://github.com/asdcorp/ohook/releases/download/%OHOOK_VERSION%/sppc32.dll"
set "TEMP_DIR=%TEMP%\ohook_%RANDOM%"
set "RESULT=0"

:: ============================================================================
:: ADMIN CHECK
:: ============================================================================
net session >nul 2>&1
if %errorlevel% neq 0 (
    if "%SILENT%"=="0" echo [ERROR] Administrator rights required
    exit /b 1
)

:: ============================================================================
:: MAIN
:: ============================================================================
if "%SILENT%"=="0" echo [INFO] Starting Ohook activation...

:: Create temp directory
mkdir "%TEMP_DIR%" 2>nul

:: Download DLLs
if "%SILENT%"=="0" (
    echo [INFO] Downloading sppc64.dll...
    powershell -Command ^
        "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;" ^
        "$ProgressPreference = 'Continue';" ^
        "try {" ^
        "    Invoke-WebRequest -Uri '%DLL64_URL%' -OutFile '%TEMP_DIR%\sppc64.dll' -UseBasicParsing;" ^
        "} catch {" ^
        "    Write-Host \"Error: $($_.Exception.Message)\";" ^
        "}"
    echo [INFO] Downloading sppc32.dll...
    powershell -Command ^
        "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;" ^
        "$ProgressPreference = 'Continue';" ^
        "try {" ^
        "    Invoke-WebRequest -Uri '%DLL32_URL%' -OutFile '%TEMP_DIR%\sppc32.dll' -UseBasicParsing;" ^
        "} catch {" ^
        "    Write-Host \"Error: $($_.Exception.Message)\";" ^
        "}"
) else (
    powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -Uri '%DLL64_URL%' -OutFile '%TEMP_DIR%\sppc64.dll' -UseBasicParsing" 2>nul
    powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -Uri '%DLL32_URL%' -OutFile '%TEMP_DIR%\sppc32.dll' -UseBasicParsing" 2>nul
)

if not exist "%TEMP_DIR%\sppc64.dll" (
    if "%SILENT%"=="0" echo [ERROR] Failed to download sppc64.dll
    set "RESULT=1"
    goto :cleanup
)
if not exist "%TEMP_DIR%\sppc32.dll" (
    if "%SILENT%"=="0" echo [ERROR] Failed to download sppc32.dll
    set "RESULT=1"
    goto :cleanup
)

:: Find and process Office installations
if "%SILENT%"=="0" echo [INFO] Detecting and activating Office...

:: Office C2R 64-bit
if exist "%ProgramFiles%\Microsoft Office\root\Office16" (
    copy /y "%TEMP_DIR%\sppc64.dll" "%ProgramFiles%\Microsoft Office\root\Office16\sppc.dll" >nul 2>&1
    if "%SILENT%"=="0" echo [OK] Office C2R 64-bit
)

:: Office C2R 32-bit
if exist "%ProgramFiles(x86)%\Microsoft Office\root\Office16" (
    copy /y "%TEMP_DIR%\sppc32.dll" "%ProgramFiles(x86)%\Microsoft Office\root\Office16\sppc.dll" >nul 2>&1
    if "%SILENT%"=="0" echo [OK] Office C2R 32-bit
)

:: Office MSI 64-bit
if exist "%ProgramFiles%\Microsoft Office\Office16" (
    copy /y "%TEMP_DIR%\sppc64.dll" "%ProgramFiles%\Microsoft Office\Office16\sppc.dll" >nul 2>&1
    if "%SILENT%"=="0" echo [OK] Office MSI 64-bit
)

:: Office MSI 32-bit
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16" (
    copy /y "%TEMP_DIR%\sppc32.dll" "%ProgramFiles(x86)%\Microsoft Office\Office16\sppc.dll" >nul 2>&1
    if "%SILENT%"=="0" echo [OK] Office MSI 32-bit
)

:: Office 2013 C2R
if exist "%ProgramFiles%\Microsoft Office\root\Office15" (
    copy /y "%TEMP_DIR%\sppc64.dll" "%ProgramFiles%\Microsoft Office\root\Office15\sppc.dll" >nul 2>&1
)
if exist "%ProgramFiles(x86)%\Microsoft Office\root\Office15" (
    copy /y "%TEMP_DIR%\sppc32.dll" "%ProgramFiles(x86)%\Microsoft Office\root\Office15\sppc.dll" >nul 2>&1
)

:: Process registry-based Office path
for /f "tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v InstallationPath 2^>nul') do (
    set "REG_PATH=%%b"
    if defined REG_PATH (
        :: Determine architecture from Platform registry
        for /f "tokens=2*" %%x in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v Platform 2^>nul') do (
            set "PLATFORM=%%y"
        )

        if /i "!PLATFORM!"=="x64" (
            if exist "!REG_PATH!\root\Office16" (
                copy /y "%TEMP_DIR%\sppc64.dll" "!REG_PATH!\root\Office16\sppc.dll" >nul 2>&1
            )
        ) else (
            if exist "!REG_PATH!\root\Office16" (
                copy /y "%TEMP_DIR%\sppc32.dll" "!REG_PATH!\root\Office16\sppc.dll" >nul 2>&1
            )
        )
    )
)

if "%SILENT%"=="0" echo [OK] Ohook activation completed

:cleanup
:: Cleanup
rd /s /q "%TEMP_DIR%" 2>nul

if "%SILENT%"=="0" (
    if "%RESULT%"=="0" (
        echo [OK] Success
    ) else (
        echo [ERROR] Failed
    )
)

exit /b %RESULT%
