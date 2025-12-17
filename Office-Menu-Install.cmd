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
    echo [ERROR] Lancez ce script en tant qu'administrateur.
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
echo      [1] Office 365 ProPlus (Recommande)
echo      [2] Office 2021 LTSC Professional Plus
echo      [3] Office 2019 Professional Plus
echo      [4] Installation avec fichier config personnalise
echo.
echo      [5] Activer Office uniquement (deja installe)
echo      [6] Verifier le statut d'activation
echo.
echo      [0] Quitter
echo.
echo  ============================================================
echo.

set /p choice="  Votre choix [0-6]: "

if "%choice%"=="1" goto install_365
if "%choice%"=="2" goto install_2021
if "%choice%"=="3" goto install_2019
if "%choice%"=="4" goto install_custom
if "%choice%"=="5" goto activate_only
if "%choice%"=="6" goto check_status
if "%choice%"=="0" goto end

echo  [ERROR] Choix invalide
timeout /t 2 >nul
goto menu

:: ============================================================================
:: OFFICE 365 INSTALLATION
:: ============================================================================
:install_365
cls
echo.
echo [INFO] Installation de Office 365 ProPlus...
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
echo [INFO] Installation de Office 2021 LTSC Professional Plus...
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
echo [INFO] Installation de Office 2019 Professional Plus...
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
echo [INFO] Installation avec fichier de configuration personnalise
echo.
set /p CONFIG_PATH="  Chemin du fichier XML: "

if not exist "%CONFIG_PATH%" (
    echo [ERROR] Fichier non trouve: %CONFIG_PATH%
    pause
    goto menu
)

set "USE_CUSTOM=1"
goto do_install

:: ============================================================================
:: COMMON INSTALLATION LOGIC
:: ============================================================================
:do_install
set "WORK_DIR=%~dp0"
set "ODT_DIR=%WORK_DIR%ODT"
set "ODT_EXE=%ODT_DIR%\setup.exe"
set "CONFIG_FILE=%ODT_DIR%\config.xml"

:: Create directory
if not exist "%ODT_DIR%" mkdir "%ODT_DIR%"

:: Download ODT
echo [INFO] Telechargement de Office Deployment Tool...
if not exist "%ODT_EXE%" (
    powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('https://officecdn.microsoft.com/pr/wsus/setup.exe', '%ODT_EXE%')}"
)

if not exist "%ODT_EXE%" (
    echo [ERROR] Echec du telechargement de ODT
    pause
    goto menu
)
echo [OK] ODT telecharge
echo.

:: Create or use config
if defined USE_CUSTOM (
    copy "%CONFIG_PATH%" "%CONFIG_FILE%" >nul
) else (
    call :create_config
)

:: Download Office
echo [INFO] Telechargement des fichiers Office...
echo        Cela peut prendre plusieurs minutes...
echo.
cd /d "%ODT_DIR%"
"%ODT_EXE%" /download "%CONFIG_FILE%"
echo.
echo [OK] Telechargement termine
echo.

:: Install Office
echo [INFO] Installation de Microsoft Office...
echo        Veuillez patienter 5-15 minutes...
echo.
"%ODT_EXE%" /configure "%CONFIG_FILE%"
echo.
echo [OK] Installation terminee
echo.

:: Wait
timeout /t 20 /nobreak >nul

:: Activate
call :run_ohook
echo.
echo [OK] Activation terminee
echo.

pause
goto menu

:: ============================================================================
:: CREATE CONFIG FILE
:: ============================================================================
:create_config
echo [INFO] Creation du fichier de configuration...

set "PIDKEY_LINE="
if defined PIDKEY set "PIDKEY_LINE= PIDKEY=\"%PIDKEY%\""

(
echo ^<Configuration^>
echo   ^<Add OfficeClientEdition="64" Channel="%CHANNEL%"^>
echo     ^<Product ID="%PRODUCT%"%PIDKEY_LINE%^>
echo       ^<Language ID="fr-fr" /^>
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

echo [OK] Configuration creee
goto :eof

:: ============================================================================
:: ACTIVATE ONLY
:: ============================================================================
:activate_only
cls
echo.
echo [INFO] Activation de Office avec Ohook...
echo.

call :run_ohook

echo.
echo [OK] Activation terminee
echo.
pause
goto menu

:: ============================================================================
:: RUN OHOOK FROM WEB
:: ============================================================================
:run_ohook
echo [INFO] Telechargement du script Ohook...

set "OHOOK_SCRIPT_URL=METTRE_URL_DU_SCRIPT_OHOOK_ICI"

powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('%OHOOK_SCRIPT_URL%', '%TEMP%\Ohook-Activate.cmd')" 2>nul

if exist "%TEMP%\Ohook-Activate.cmd" (
    echo [OK] Script telecharge
    call "%TEMP%\Ohook-Activate.cmd"
    del /f /q "%TEMP%\Ohook-Activate.cmd" 2>nul
) else (
    echo [ERROR] Echec du telechargement du script Ohook
    echo         URL: %OHOOK_SCRIPT_URL%
)
goto :eof

:: ============================================================================
:: CHECK STATUS
:: ============================================================================
:check_status
cls
echo.
echo [INFO] Verification du statut d'activation Office...
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
    echo [WARNING] Office ne semble pas etre installe ou OSPP.VBS introuvable
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
echo  Au revoir!
echo.
exit /b 0
