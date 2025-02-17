# Start logging
$logFile = "C:\Temp\PSWindowsUpdate.log"
Start-Transcript -Path $logFile -Append

# Powershell and Windows Updates
Write-Host -ForegroundColor Green "Powershell modules and Windows Updates" | Tee-Object -FilePath $logFile -Append
Start-Sleep 3

# Verify/Elevate Admin Session. Comment out if not needed.
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { 
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs 
    exit
}

# Define and use TLS1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 

# Register PSGallery PSprovider, set it as Trusted source, Verify NuGet
Register-PSRepository -Default -ErrorAction SilentlyContinue
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction SilentlyContinue
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope AllUsers -Force -Confirm:$false -ErrorAction SilentlyContinue
Install-Module -Name PowerShellGet -MinimumVersion 2.2.4 -Scope AllUsers -Force -Confirm:$false -ErrorAction SilentlyContinue


# Update or install necessary modules
$modules = Get-InstalledModule | Select-Object -ExpandProperty "Name"
$Modules += @("PSWindowsUpdate")

Foreach ($Module In $Modules) {
    $currentVersion = $null
    if ($null -ne (Get-InstalledModule -Name $Module -ErrorAction SilentlyContinue)) {
        $currentVersion = (Get-InstalledModule -Name $module -AllVersions).Version
    }
    $CurrentModule = Find-Module -Name $module

    if ($null -eq $currentVersion) {
        Write-Host "$($CurrentModule.Name) - Installing $Module from PowerShellGallery. Version: $($CurrentModule.Version)" | Tee-Object -FilePath $logFile -Append
        try {
            Install-Module -Name $module -Force
        } catch {
            Write-Host -ForegroundColor Red "Failed to install $Module. Details: $_" | Tee-Object -FilePath $logFile -Append
        }
    } elseif ($CurrentModule.Version -eq $currentVersion) {
        Write-Host -ForegroundColor Green "$($CurrentModule.Name) is up to date. Version: $currentVersion" | Tee-Object -FilePath $logFile -Append
    } else {
        Write-Host "$($CurrentModule.Name) - Updating from version $currentVersion to $($CurrentModule.Version)" | Tee-Object -FilePath $logFile -Append
        try {
            Update-Module -Name $module -Force
            Write-Host -ForegroundColor Green "$Module Successfully Updated" | Tee-Object -FilePath $logFile -Append
        } catch {
            Write-Host -ForegroundColor Red "Failed to update $Module. Details: $_" | Tee-Object -FilePath $logFile -Append
        }
    }
}

Write-Host "`nRunning Windows Updates..." | Tee-Object -FilePath $logFile -Append
Import-Module PSWindowsUpdate -Force

# Check if Microsoft Update Service is available
Write-Host "Checking Microsoft Update Service Registration..." | Tee-Object -FilePath $logFile -Append
$MicrosoftUpdateServiceId = "7971f918-a847-4430-9279-4a52d1efe18d"
If ((Get-WUServiceManager -ServiceID $MicrosoftUpdateServiceId).ServiceID -eq $MicrosoftUpdateServiceId) {
    Write-Host "Confirmed: Microsoft Update Service is registered." | Tee-Object -FilePath $logFile -Append
} Else {
    Add-WUServiceManager -ServiceID $MicrosoftUpdateServiceId -Confirm:$true | Tee-Object -FilePath $logFile -Append
}

# Verify registration
If (!((Get-WUServiceManager -ServiceID $MicrosoftUpdateServiceId).ServiceID -eq $MicrosoftUpdateServiceId)) {
    Write-Error "ERROR: Microsoft Update Service is still not registered after 2 attempts. Try running a WU repair script." | Tee-Object -FilePath $logFile -Append
    Stop-Transcript
    exit
}

# Run Windows Updates (without auto-reboot)
Write-Host "Running Updates..." | Tee-Object -FilePath $logFile -Append
$updateResults = Install-WindowsUpdate -AcceptAll -ForceInstall -IgnoreReboot -Verbose | Tee-Object -FilePath $logFile -Append

Write-Host "`nUpdates Complete" | Tee-Object -FilePath $logFile -Append

# Stop logging
Stop-Transcript
