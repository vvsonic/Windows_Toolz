SET ThisScriptsDirectory=%~dp0
SET PowerShellScriptPath=%ThisScriptsDirectory%DebloatScript-HKCU.ps1

:: Run the main PowerShell script (for debloating, settings, etc.)
powershell.exe -ExecutionPolicy Bypass -File "C:\Windows\FirstUserLogon\DebloatScript-HKCU.ps1"

:: Run the Cmd-HKCU.cmd script after the PowerShell script finishes
call "C:\Windows\FirstUserLogon\Cmd-HKCU.cmd"

:: Pause to ensure the scripts finish executing before deleting registry key
timeout /t 5 /nobreak >nul

:: Delete the registry key to ensure the script doesn't run on subsequent logins
REG DELETE "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v FirstUserLogon /f

:: Exit the script
exit
