# Verify/Elevate Admin Session.
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { 
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit 
}

Write-Host " "
Write-Host " "
Write-Host "  ## ##   ####     ### ###    ##     ###  ##  ##  ###  ### ##   "
Write-Host " ##   ##   ##       ##  ##     ##      ## ##  ##   ##   ##  ##  "
Write-Host " ##        ##       ##       ## ##    # ## #  ##   ##   ##  ##  "
Write-Host " ##        ##       ## ##    ##  ##   ## ##   ##   ##   ##  ##  "
Write-Host " ##        ##       ##       ## ###   ##  ##  ##   ##   ## ##   "
Write-Host " ##   ##   ##  ##   ##  ##   ##  ##   ##  ##  ##   ##   ##      "
Write-Host "  ## ##   ### ###  ### ###  ###  ##  ###  ##   ## ##   ####    "
Write-Host " "
Write-Host "  Created by Alex Ivantsov "
Write-Host "  @Exploitacious "

Write-Host

Write-Host "Launching PS Modules & Windows Updates"
Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"C:\Temp\Cleanup\PSandWindowsUpdates.ps1`"" -Verb RunAs

# Countdown timer before launching the next script
$i = 180 #Seconds
do {
    Write-Host $i
    Sleep 1
    $i--
} while ($i -gt 0)

Write-Host "Launching De-Bloat Processes..."
Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"C:\Temp\Cleanup\UninstallBloat.ps1`"" -Verb RunAs

# Countdown timer before launching the next script
$i = 5 #Seconds
do {
    Write-Host $i
    Sleep 1
    $i--
} while ($i -gt 0)

Write-Host "Launching Windows tweaks and settings..."
Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"C:\Temp\Cleanup\PS-HKLM.ps1`"" -Verb RunAs

# Ensure proper exit when done
Write-Host "All tasks completed!"
Write-Host "Press Enter to exit the script."
Read-Host -Prompt "Finished! Press Enter to exit"

# Explicitly exit after completion to ensure the script terminates
exit
