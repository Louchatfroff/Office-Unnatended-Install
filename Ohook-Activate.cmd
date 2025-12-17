@echo off
:: ============================================================================
:: Ohook Office Activation Script - Standalone Version
:: Based on ohook by asdcorp - https://github.com/asdcorp/ohook
:: ============================================================================
:: Ce script active Microsoft Office en utilisant la methode Ohook
:: Il place un fichier sppc.dll personnalise dans le dossier Office
:: ============================================================================

setlocal EnableDelayedExpansion
title Ohook Office Activation

:: ============================================================================
:: CONFIGURATION - URLs des DLL (modifier si necessaire)
:: ============================================================================
set "OHOOK_VERSION=0.5"
set "OHOOK_URL_BASE=https://github.com/asdcorp/ohook/releases/download/%OHOOK_VERSION%"
set "DLL32_URL=%OHOOK_URL_BASE%/sppc32.dll"
set "DLL64_URL=%OHOOK_URL_BASE%/sppc64.dll"

:: SHA256 checksums pour verification
set "DLL32_HASH=09865ea5993215965e8f27a74b8a41d15fd0f60f5f404cb7a8b3c7757acdab02"
set "DLL64_HASH=393a1fa26deb3663854e41f2b687c188a9eacd87b23f17ea09422c4715cb5a9f"

:: Dossier temporaire
set "TEMP_DIR=%TEMP%\ohook_temp"
set "DLL32_PATH=%TEMP_DIR%\sppc32.dll"
set "DLL64_PATH=%TEMP_DIR%\sppc64.dll"

:: ============================================================================
:: VERIFICATION ADMIN
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
    echo [ERREUR] Ce script necessite les droits Administrateur.
    echo          Clic droit - Executer en tant qu'administrateur
    echo.
    pause
    exit /b 1
)
echo [OK] Droits administrateur confirmes
echo.

:: ============================================================================
:: DETECTION OFFICE
:: ============================================================================
:detect_office
echo [INFO] Detection des installations Office...
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
        echo [TROUVE] %%~a
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
        echo [TROUVE] %%~a
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
                echo [TROUVE] !C2R_PATH!\root\%%v
                set "OFFICE_FOUND=1"
                set "OFFICE_PATHS=!OFFICE_PATHS!!C2R_PATH!\root\%%v;"
            )
        )
    )
)

echo.
if "%OFFICE_FOUND%"=="0" (
    echo [ERREUR] Aucune installation Office detectee.
    echo          Installez Office avant d'executer ce script.
    echo.
    pause
    exit /b 1
)

:: ============================================================================
:: TELECHARGEMENT DES DLL
:: ============================================================================
:download_dlls
echo [INFO] Telechargement des fichiers Ohook...
echo.

if not exist "%TEMP_DIR%" mkdir "%TEMP_DIR%"

:: Download 64-bit DLL
echo [INFO] Telechargement sppc64.dll...
powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('%DLL64_URL%', '%DLL64_PATH%')}" 2>nul

if not exist "%DLL64_PATH%" (
    echo [ERREUR] Echec du telechargement de sppc64.dll
    echo          Verifiez votre connexion internet
    echo          URL: %DLL64_URL%
    pause
    exit /b 1
)
echo [OK] sppc64.dll telecharge

:: Download 32-bit DLL
echo [INFO] Telechargement sppc32.dll...
powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('%DLL32_URL%', '%DLL32_PATH%')}" 2>nul

if not exist "%DLL32_PATH%" (
    echo [ERREUR] Echec du telechargement de sppc32.dll
    echo          Verifiez votre connexion internet
    echo          URL: %DLL32_URL%
    pause
    exit /b 1
)
echo [OK] sppc32.dll telecharge
echo.

:: ============================================================================
:: VERIFICATION DES HASH (optionnel mais recommande)
:: ============================================================================
:verify_hash
echo [INFO] Verification de l'integrite des fichiers...

for /f "skip=1 tokens=* delims=" %%a in ('certutil -hashfile "%DLL64_PATH%" SHA256 2^>nul') do (
    set "COMPUTED_HASH=%%a"
    goto :check_hash64
)
:check_hash64
set "COMPUTED_HASH=%COMPUTED_HASH: =%"
if /i "%COMPUTED_HASH%"=="%DLL64_HASH%" (
    echo [OK] sppc64.dll - Hash valide
) else (
    echo [ATTENTION] sppc64.dll - Hash different (peut etre une nouvelle version^)
)

for /f "skip=1 tokens=* delims=" %%a in ('certutil -hashfile "%DLL32_PATH%" SHA256 2^>nul') do (
    set "COMPUTED_HASH=%%a"
    goto :check_hash32
)
:check_hash32
set "COMPUTED_HASH=%COMPUTED_HASH: =%"
if /i "%COMPUTED_HASH%"=="%DLL32_HASH%" (
    echo [OK] sppc32.dll - Hash valide
) else (
    echo [ATTENTION] sppc32.dll - Hash different (peut etre une nouvelle version^)
)
echo.

:: ============================================================================
:: INSTALLATION DES DLL
:: ============================================================================
:install_dlls
echo [INFO] Installation des fichiers Ohook...
echo.

:: Process each Office path
for %%p in (%OFFICE_PATHS%) do (
    set "CURRENT_PATH=%%~p"
    if not "!CURRENT_PATH!"=="" (
        echo [INFO] Traitement: !CURRENT_PATH!

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
            echo        [OK] sppc.dll installe
        ) else (
            echo        [ERREUR] Impossible de copier sppc.dll
            echo                 Fermez toutes les applications Office et reessayez
        )
        echo.
    )
)

:: ============================================================================
:: INSTALLATION DES LICENCES
:: ============================================================================
:install_licenses
echo [INFO] Installation des licences Office...
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
        echo [INFO] Dossier licences: !LICENSE_PATH!

        :: Install Grace licenses for all products
        for %%p in (
            "%ProgramFiles%\Microsoft Office\root\Office16"
            "%ProgramFiles(x86)%\Microsoft Office\root\Office16"
        ) do (
            if exist "%%~p\OSPP.VBS" (
                echo [INFO] Installation des licences via OSPP.VBS...

                :: Install volume licenses
                for /f "delims=" %%l in ('dir /b "!LICENSE_PATH!\*VL_*.xrm-ms" 2^>nul') do (
                    cscript //nologo "%%~p\OSPP.VBS" /inslic:"!LICENSE_PATH!\%%l" >nul 2>&1
                )

                echo [OK] Licences installees
            )
        )
    )
)
echo.

:: ============================================================================
:: ACTIVATION
:: ============================================================================
:activate
echo [INFO] Activation d'Office...
echo.

:: Try to activate via OSPP
for %%p in (
    "%ProgramFiles%\Microsoft Office\root\Office16"
    "%ProgramFiles(x86)%\Microsoft Office\root\Office16"
    "%ProgramFiles%\Microsoft Office\Office16"
    "%ProgramFiles(x86)%\Microsoft Office\Office16"
) do (
    if exist "%%~p\OSPP.VBS" (
        echo [INFO] Utilisation de: %%~p\OSPP.VBS

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
echo [INFO] Verification du statut d'activation...
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
:: NETTOYAGE
:: ============================================================================
:cleanup
echo.
echo ============================================
echo.
echo [INFO] Nettoyage des fichiers temporaires...
rd /s /q "%TEMP_DIR%" 2>nul
echo [OK] Nettoyage termine
echo.

:: ============================================================================
:: FIN
:: ============================================================================
:end
echo ============================================
echo   Activation Ohook terminee!
echo ============================================
echo.
echo Si l'activation a echoue, essayez:
echo   1. Fermez toutes les applications Office
echo   2. Relancez ce script en administrateur
echo   3. Redemarrez l'ordinateur et reessayez
echo.
echo Source: https://github.com/asdcorp/ohook
echo.
pause
exit /b 0
