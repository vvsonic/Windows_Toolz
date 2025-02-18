# Start logging
$logFile = "C:\Temp\PSWindowsUpdate.log"
Write-Host "Powershell modules and Windows Updates" | Out-File -FilePath $logFile -Append
Start-Sleep 3

# Check if running as Administrator, if not, re-launch as Admin
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Not running as Administrator. Relaunching as Administrator..." | Out-File -FilePath $logFile -Append
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}
# Ensure TLS 1.2 for secure connections
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 

# Directly install the NuGet provider without checking if it's installed
Write-Host "Ensuring NuGet provider is installed..."
Install-PackageProvider -Name NuGet -Confirm:$false -Scope AllUsers -Force

# Verify if NuGet is now installed
$nugetInstalled = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
if ($null -eq $nugetInstalled) {
    Write-Host "NuGet provider is still not installed. Exiting script." -ForegroundColor Red
    exit
}

# Trust PSGallery to avoid prompts
if ((Get-PSRepository -Name 'PSGallery').InstallationPolicy -ne 'Trusted') {
    Write-Host "Trusting PSGallery..."
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction Stop
}

# Ensure PSWindowsUpdate module is installed
if (-not (Get-Module -Name PSWindowsUpdate -ListAvailable -ErrorAction SilentlyContinue)) {
    Write-Host "Installing PSWindowsUpdate module..."
    Install-Module -Name PSWindowsUpdate -Force -Confirm:$false -Scope AllUsers
}

# Import the module (if not already imported)
if (-not (Get-Module -Name PSWindowsUpdate)) {
    Write-Host "Importing PSWindowsUpdate module..."
    Import-Module PSWindowsUpdate -Force -ErrorAction Stop
}

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
        Write-Host "$($CurrentModule.Name) - Installing $Module from PowerShellGallery. Version: $($CurrentModule.Version)"
        try {
            Install-Module -Name $module -Force -Confirm:$false
        } catch {
            Write-Host -ForegroundColor Red "Failed to install $Module. Details: $_"
        }
    } elseif ($CurrentModule.Version -eq $currentVersion) {
        Write-Host -ForegroundColor Green "$($CurrentModule.Name) is up to date. Version: $currentVersion"
    } else {
        Write-Host "$($CurrentModule.Name) - Updating from version $currentVersion to $($CurrentModule.Version)"
        try {
            Update-Module -Name $module -Force -Confirm:$false
            Write-Host -ForegroundColor Green "$Module Successfully Updated"
        } catch {
            Write-Host -ForegroundColor Red "Failed to update $Module. Details: $_"
        }
    }
}

# Run Windows Updates
Write-Host "`nRunning Updates..." | Out-File -FilePath $logFile -Append

# Check Microsoft Update Service Registration
Write-Host "Checking Microsoft Update Service Registration..." | Out-File -FilePath $logFile -Append
$MicrosoftUpdateServiceId = "7971f918-a847-4430-9279-4a52d1efe18d"
If ((Get-WUServiceManager -ServiceID $MicrosoftUpdateServiceId).ServiceID -eq $MicrosoftUpdateServiceId) {
    Write-Host "Confirmed: Microsoft Update Service is registered." | Out-File -FilePath $logFile -Append
} Else {
    Add-WUServiceManager -ServiceID $MicrosoftUpdateServiceId -Confirm:$false | Out-File -FilePath $logFile -Append
}

# Verify registration
If (!((Get-WUServiceManager -ServiceID $MicrosoftUpdateServiceId).ServiceID -eq $MicrosoftUpdateServiceId)) {
    Write-Error "ERROR: Microsoft Update Service is still not registered after 2 attempts. Try running a WU repair script." | Out-File -FilePath $logFile -Append
    exit
}

# Run Windows Updates (without auto-reboot)
Write-Host "Running Updates..." | Out-File -FilePath $logFile -Append
$updateResults = Install-WindowsUpdate -AcceptAll -ForceInstall -IgnoreReboot -Verbose | Out-File -FilePath $logFile -Append

Write-Host "`nUpdates Complete" | Out-File -FilePath $logFile -Append
