:: Download and run the system debloat scripts directly from GitHub

@echo off

:: Checking for Administrator elevation.

        openfiles>nul 2>&1

        if %errorlevel% EQU 0 goto :Download

        echo.
        echo.    You are not running as Administrator.
        echo.    This script cannot do it's job without elevation.
        echo.
        echo.    You need run this tool as Administrator.
        echo.

    exit


:: Download Required Files from https://github.com/Exploitacious/Windows_Toolz/tree/main/Windows%2BServer/System_Debloat
:Download

    PowerShell -executionpolicy bypass -Command "(New-Object Net.WebClient).DownloadFile('https://raw.githubusercontent.com/Exploitacious/Windows_Toolz/main/System_Debloat/SYSTEM-Debloat-MAIN.ps1', 'SYSTEM-Debloat-MAIN.ps1')"

    PowerShell -executionpolicy bypass -Command "(New-Object Net.WebClient).DownloadFile('https://raw.githubusercontent.com/Exploitacious/Windows_Toolz/main/System_Debloat/DebloatScript-HKCU.ps1', 'DebloatScript-HKCU.ps1')"
    
    PowerShell -executionpolicy bypass -Command "(New-Object Net.WebClient).DownloadFile('https://raw.githubusercontent.com/Exploitacious/Windows_Toolz/main/System_Debloat/FirstLogon.bat', 'FirstLogon.bat')"



:: Start Running the SYSTEM DEBLOAT scripts
:RunScript

    SET ThisScriptsDirectory=%~dp0
    SET PowerShellScriptPath=%ThisScriptsDirectory%SYSTEM-Debloat-MAIN.ps1
    PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& '%PowerShellScriptPath%'";