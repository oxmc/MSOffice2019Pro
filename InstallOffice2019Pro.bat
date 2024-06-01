@echo off
setlocal

:: Check if the script is running with administrative privileges
>nul 2>&1 "%SystemRoot%\system32\cacls.exe" "%SystemRoot%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo Requesting administrative privileges...
    goto NewUACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "cmd.exe", "/c ""%~0"" %*", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /b

:NewUACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"="
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    pushd "%CD%"
    CD /D "%~dp0"

:: Set URLs and local paths
set "ODT_URL=https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17531-20046.exe"
set "CONFIG_URL=https://cdn.oxmc.me/uploads"
set "DEFAULT_CONF_PATH=DownloadedConfig.xml"
set "EXTRACTDIR=Office2019Extracted"
set "SETUP_APP=Office365SetupTool.exe"

:: Parse command line arguments
set "CONF_MODE=default-office-config.xml"
set "CONF_PATH=%DEFAULT_CONF_PATH%"

:parse_args
if "%~1"=="" goto args_done
if "%~1"=="/conf" (
    shift
    set "CONF_MODE=%~1"
    shift
    goto parse_args
)
if "%~1"=="/conf-path" (
    shift
    set "CONF_PATH=%~1"
    shift
    goto parse_args
)
shift
goto parse_args
:args_done

:: Check conf mode
if "%CONF_MODE%"=="local" (
    if not defined CONF_PATH (
        echo /conf-path must be specified if /conf is set to local.
        timeout /t 3 /nobreak >nul
        exit /b 1
    )
) else if "%CONF_MODE%"=="url" (
    if not defined CONF_PATH (
        echo /conf-path must be specified if /conf is set to url.
        timeout /t 3 /nobreak >nul
        exit /b 1
    ) else (
        echo Downloading configuration file from %CONF_PATH%...
        curl -L -o "%DEFAULT_CONF_PATH%" "%CONF_PATH%"
        if errorlevel 1 (
            echo Failed to download configuration file.
            timeout /t 3 /nobreak >nul
            exit /b 1
        )
    )
) else (
    echo Downloading default configuration file from cdn.oxmc.me...
    curl -L -o "%DEFAULT_CONF_PATH%" "%CONFIG_URL%/default-office-config.xml"
    if errorlevel 1 (
        echo Failed to download configuration file.
        timeout /t 3 /nobreak >nul
        exit /b 1
    )
    set "CONF_PATH=%DEFAULT_CONF_PATH%"
)

echo Configuration mode: %CONF_MODE%
echo Configuration path: %CONF_PATH%

:: Download ODT
echo Downloading %ODT_URL% to Office365SetupTool.exe...
curl -L -o "%SETUP_APP%" "%ODT_URL%"
if errorlevel 1 (
    echo Failed to download %SETUP_APP%.
    timeout /t 3 /nobreak >nul
    exit /b 1
)

:: Check if DownloadedConfig.xml exists
if exist "DownloadedConfig.xml" (
    set "CONF_PATH=DownloadedConfig.xml"
    echo Using downloaded configuration file
) else (
    :: Download the CONF file
    echo Downloading %CONF_URL% to DownloadedConfig.xml...
    curl -L -o "DownloadedConfig.xml" "%CONF_URL%"
    if errorlevel 1 (
        echo Failed to download Config file.
        timeout /t 3 /nobreak >nul
        exit /b 1
    )
    set "CONF_PATH=DownloadedConfig.xml"
)

:: Wait for the files to finish downloading
:wait_for_tool
if not exist "%SETUP_APP%" goto wait_for_tool

:: Create extract directory
if not exist "%EXTRACTDIR%" mkdir "%EXTRACTDIR%"

:: Run the Setup executable
echo Running %SETUP_APP%
start "" "%SETUP_APP%" "/extract:%EXTRACTDIR%" "/passive"
if errorlevel 1 (
    echo Failed to run Setup.
    timeout /t 3 /nobreak >nul
    exit /b 1
)

:wait_for_setup
if not exist "%EXTRACTDIR%\setup.exe" goto wait_for_setup

:: Run the OfficeSetup executable
echo Running OfficeSetup with configuration path %CONF_PATH%...
start "" "%EXTRACTDIR%\setup.exe" "/configure" "%CONF_PATH%"
if errorlevel 1 (
    echo Failed to run OfficeSetup.
    timeout /t 3 /nobreak >nul
    exit /b 1
)

echo Done.
endlocal
exit /b 0