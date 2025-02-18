# Powershell and Windows Updates
Write-Host -ForegroundColor Green "Powershell modules and Windows Updates" | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
Start-Sleep 3 | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append

# This script provides information about the module version (current and the latest available on PowerShell Gallery) and update to the latest version
# If you have a module with two or more versions, the script deletes them and reinstalls only the latest.

# Verify/Elevate Admin Session. Comment out if not needed.
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { 
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit 
}

# Define and use TLS1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 

## Register PSGallery PSprovider, set it as Trusted source, Verify NuGet
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope AllUsers -Force -ErrorAction SilentlyContinue | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
Register-PSRepository -Default -ErrorAction SilentlyContinue | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction SilentlyContinue | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
Install-Module PowerShellGet -MinimumVersion 2.2.4 -Scope AllUsers -Force -ErrorAction SilentlyContinue | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append

# All Modules (Update Only)
$modules = Get-InstalledModule | Select-Object -ExpandProperty "Name"

## New Apps and Modules To Install / Update
$Modules += @(
    "PSWindowsUpdate"
)

# Run through all modules, update and install new apps
Foreach ($Module In $Modules) {
    $currentVersion = $null
    if ($null -ne (Get-InstalledModule -Name $Module -ErrorAction SilentlyContinue)) {
        $currentVersion = (Get-InstalledModule -Name $module -AllVersions).Version
    }

    $CurrentModule = Find-Module -Name $module

    if ($null -eq $currentVersion) {
        Write-Host "$($CurrentModule.Name) - Installing $Module from PowerShellGallery. Version: $($CurrentModule.Version). Release date: $($CurrentModule.PublishedDate)" | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
        try {
            Install-Module -Name $module -Force | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
        }
        catch {
            Write-Host -ForegroundColor Red "Something went wrong when installing $Module. Please uninstall and try re-installing this module. (Remove-Module, Install-Module) Details:" | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
            Write-Error "$_.Exception.Message" | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
        }
    }
    elseif ($CurrentModule.Version -eq $currentVersion) {
        Write-Host -ForegroundColor Green "$($CurrentModule.Name) is installed and ready. Version: ($currentVersion. Release date: $($CurrentModule.PublishedDate))" | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
    }
    elseif ($currentVersion.count -gt 1) {
        Write-Host "$module is installed in $($currentVersion.count) versions (versions: $($currentVersion -join ' | '))" | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
        Write-Host "Uninstalling previous $module versions and will attempt to update." | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
        try {
            Get-InstalledModule -Name $module -AllVersions | Where-Object { $_.Version -ne $CurrentModule.Version } | Uninstall-Module -Force | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
        }
        catch {
            Write-Host -ForegroundColor Red "Something went wrong with Uninstalling $Module previous versions. Please Completely uninstall and re-install this module. (Remove-Module) Details:" | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
            Write-Error "$_.Exception.Message" | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
        }

        Write-Host "$($CurrentModule.Name) - Installing version from PowerShellGallery $($CurrentModule.Version). Release date: $($CurrentModule.PublishedDate)" | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append

        try {
            Install-Module -Name $module -Force | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
            Write-Host "$Module Successfully Installed" | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
        }
        catch {
            Write-Host -ForegroundColor Red "Something went wrong with installing $Module. Details:" | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
            Write-Error "$_.Exception.Message" | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
        }
    }
    else {       
        Write-Host "$($CurrentModule.Name) - Updating from PowerShellGallery from version $currentVersion to $($CurrentModule.Version). Release date: $($CurrentModule.PublishedDate)" | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
        try {
            Update-Module -Name $module -Force | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
            Write-Host -ForegroundColor Green "$Module Successfully Updated" | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
        }
        catch {
            Write-Host -ForegroundColor Red "Something went wrong with updating $Module. Details:" | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
            Write-Error "$_.Exception.Message" | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
        }
    }
}

Write-Host | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
Write-Host "Running Windows Updates..." | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
Write-Host | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append

Import-Module PSWindowsUpdate -Force | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append

# Check if the Microsoft Update service is available.
Write-Host "Checking Microsoft Update Service Registration..." | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
$MicrosoftUpdateServiceId = "7971f918-a847-4430-9279-4a52d1efe18d"
If ((Get-WUServiceManager -ServiceID $MicrosoftUpdateServiceId).ServiceID -eq $MicrosoftUpdateServiceId) {
    Write-Host "Confirmed Microsoft Update Service is registered." | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
}
Else {
    Add-WUServiceManager -ServiceID $MicrosoftUpdateServiceId -Confirm:$false | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
}

# Now, check again to ensure it is available. If not -- fail the script, otherwise proceed to updating Windows:
If (!((Get-WUServiceManager -ServiceID $MicrosoftUpdateServiceId).ServiceID -eq $MicrosoftUpdateServiceId)) {
    Write-Error "ERROR: Microsoft Update Service still not registered after 2 attempts. Try running WU repair script." | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
}
else {
    # Initiate download and install of all pending Windows updates
    Write-Host "Running Updates...." | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
    Install-WindowsUpdate -AcceptAll -ForceInstall -IgnoreReboot -Verbose | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
    Write-Host | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
    Write-Host "Updates Complete" | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
}
# Wait for 60 seconds
Write-Host "Waiting for 60 seconds before exiting..." | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
Start-Sleep -Seconds 60

# Exit the script
Write-Host "Exiting script." | Tee-Object -FilePath "C:\Temp\PSWindowsUpdate.log" -Append
Exit

