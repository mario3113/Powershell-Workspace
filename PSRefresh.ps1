#requires -version 5.1
#requires -RunAsAdministrator

#PSRefresh.ps1

<#
Update key PowerShell components on a new Windows 10/11 installation.

This script is not intended for server operating systems. The script
should be run in an interactive console session and not in a remoting session.

You can modify this script to use a different package manager like Chocolatey.

If you use the Offline option, make sure your file names match the script.

This script is offered AS-IS and without warranty. Use at your own risk.
#>

#TODO: Add SupportsShouldProcess code
#TODO: Add proper error handling

[CmdletBinding()]
Param(
    [Parameter(Position = 0,Mandatory,HelpMessage = 'The path to a configuration data file')]
    [ValidateScript({ Test-Path -Path $_})]
    [ValidatePattern('\.psd1$')]
    [string]$ConfigurationData,
    [Parameter(HelpMessage = 'Specify a location with previously installed Appx packages')]
    [ValidateScript({ Test-Path -Path $_ })]
    [string]$Offline
)

#this script should be run in the console, not the ISE or VSCode
if ($Host.name -ne 'ConsoleHost') {
    Write-Warning 'This script should be run in the PowerShell console, not the ISE or VSCode'
    return
}

#region Setup
Try {
    $data = Import-PowerShellDataFile -Path $ConfigurationData -ErrorAction Stop
}
Catch {
    Write-Warning "Failed to import $ConfigurationData"
    Return
}

#define a list of winget package IDs
#$wingetPackages = @('Microsoft.VisualStudioCode', 'Git.Git', 'GitHub.cli', 'Microsoft.WindowsTerminal')
$wingetPackages = $data.wingetPackages
$PSModules = $data.PSModules
$Scope = $data.Scope
$vscExtensions = $data.vscExtensions

#install winget apps and additional PowerShell modules via background jobs
$jobs = @()

$installParams = @{
    Scope        = $Scope
    Repository   = 'PSGallery'
    Force        = $true
    AllowClobber = $true
    Name         = $null
}

$progParams = @{
    Activity = $MyInvocation.MyCommand
    Status   = 'Initializing'
    CurrentOperation = 'Bootstrapping a NuGet provider update'
    PercentComplete = 1
}

#set TLS just in case
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Progress @progParams

#Bootstrap Nuget provider update  to avoid interactive prompts
[void](Install-PackageProvider -Name Nuget -ForceBootstrap -Force)

#endregion

#region Update PowerShellGet
$progParams.Status = "Module updates"
$progParams.CurrentOperation = "Updating PowerShellGet"
$progParams.PercentComplete = 10
Write-Progress @progParams

$get = Get-Module PowerShellGet -ListAvailable | Select-Object -First 1
if ($get.Version.major -eq 1) {
    #Write-Host 'Installing the latest version of PowerShellGet' -ForegroundColor Yellow
    $installParams.Name = 'PowerShellGet'
    Install-Module @installParams
}
else {
    #Write-Host 'Updating PowerShellGet' -ForegroundColor Yellow
    Update-Module -Name PowerShellGet -Force
}

#reload PowerShellGet
Remove-Module PowerShellGet
Import-Module PowerShellGet

#endregion

#region Install Microsoft.PowerShell.PSResourceGet
$progParams.Status = "Module updates"
$progParams.CurrentOperation = "Microsoft.PowerShell.PSResourceGet"
$progParams.PercentComplete = 20
Write-Progress @progParams
#Write-Host 'Installing Microsoft.PowerShell.PSResourceGet' -ForegroundColor Yellow
$installParams.Name = 'Microsoft.PowerShell.PSResourceGet'
Install-Module @installParams

Import-Module Microsoft.PowerShell.PSResourceGet

#endregion

#region Install updated Modules

$progParams.Status = "Module updates"
$progParams.CurrentOperation = "PSReadLine"
$progParams.PercentComplete = 25
Write-Progress @progParams

#Write-Host 'Installing PSReadLine' -ForegroundColor Yellow
Install-PSResource -Name PSReadLine -Scope $Scope -Repository PSGallery -TrustRepository

$progParams.Status = "Module updates"
$progParams.CurrentOperation = "Pester - You may see a warning."
$progParams.PercentComplete = 30
Write-Progress @progParams

#Write-Host 'Installing Pester. You might see a warning.' -ForegroundColor Yellow
Install-PSResource -Name Pester -Scope $Scope -Repository PSGallery -TrustRepository

#endregion

#region install winget dependencies

$progParams.Status = "Installing Winget"
$progParams.CurrentOperation = "Processing dependencies"
$progParams.PercentComplete = 40
Write-Progress @progParams
#Write-Host 'Adding Nuget.org as a package source' -ForegroundColor Yellow
[void](Register-PackageSource -Name Nuget.org -ProviderName NuGet -Force -ForceBootstrap -Location 'https://nuget.org/api/v2')

#Write-Host 'Installing winget dependencies' -ForegroundColor Yellow

if ($Offline) {
    Add-AppxPackage "$Offline\microsoft.ui.xaml.2.8.appx"
    Add-AppxPackage "$Offline\VCLibs.appx"
}
else {
    [void](Install-Package -Name microsoft.ui.xaml -Source nuget.org -Force)
    Copy-Item -Path (Get-Package microsoft.ui.xaml).source -Destination $env:TEMP\microsoft.ui.xaml.zip
    Expand-Archive $env:temp\microsoft.ui.xaml.zip -DestinationPath "$env:temp\ui" -Force
    Add-AppxPackage $env:temp\ui\tools\appx\x64\release\Microsoft.UI.Xaml.2.8.appx
    $uri = 'https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx'
    Invoke-WebRequest -Uri $uri -OutFile $env:temp\VCLibs.appx
    Add-AppxPackage $env:temp\VCLibs.appx
}

#endregion

#region Install winget

#Write-Host 'Installing winget' -ForegroundColor Yellow
$progParams.Status = "Installing Winget"
$progParams.CurrentOperation = "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
$progParams.PercentComplete = 50
Write-Progress @progParams
if ($Offline) {
    Add-AppxPackage "$Offline\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
}
else {
    #Winget is a 246MB download
    $uri = 'https://api.github.com/repos/microsoft/winget-cli/releases'
    $get = Invoke-RestMethod -Uri $uri -Method Get -ErrorAction stop
    $current = $get[0].assets | Where-Object name -Match 'msixbundle'

    #Write-Host "Downloading $($current.name)" -ForegroundColor Yellow
    $out = Join-Path -Path $env:temp -child $current.name
    Try {
        Invoke-WebRequest -Uri $current.browser_download_url -OutFile $out -ErrorAction Stop
        Add-AppxPackage -Path $out
    }
    Catch {
        $_
    }
}

#endregion

#region Install winget packages
$progParams.Status = "Installing packages via winget"
$pct = 50

foreach ($package in $wingetPackages) {
    $progParams.CurrentOperation = $package
    $progParams.PercentComplete = $pct+=2
    Write-Progress @progParams
    $jobs+= Start-Job -Name $package -ScriptBlock {
        #This script does not take scope into account for Winget installations.
        #You might want to change that.
        Param($package)
        winget install --id $package --silent --accept-package-agreements --accept-source-agreements --source winget
    } -ArgumentList $package

    #Write-Host "Installing $package" -ForegroundColor Yellow
    #winget install --id $package --silent --accept-package-agreements --accept-source-agreements --source winget
} #foreach package

#endregion

#region install additional PowerShell modules

if ($PSModules) {
    $progParams.Status = "Installing additional PowerShell modules"
    foreach ($Mod in $PSModules) {
        $progParams.CurrentOperation = $Mod
        $progParams.PercentComplete = $pct+=2
        Write-Progress @progParams
        $jobs+= Start-Job -Name $Mod -ScriptBlock {
            Param($ModuleName,$scope)
            Import-Module Microsoft.PowerShell.PSResourceGet
            Install-PSResource -Name $ModuleName -Scope $Scope -Repository PSGallery -AcceptLicense -TrustRepository -Quiet
        } -ArgumentList $Mod,$Scope
    } #foreach module
}

#endregion

#region install VSCode extensions

if ($vscExtensions) {

    $progParams.Status = "Installing VSCode extensions"
    $progParams.PercentComplete = $pct+=2
    foreach ($Extension in $vscExtensions) {
        $progParams.CurrentOperation = $Extension
        Write-Progress @progParams
        $jobs+= Start-Job -Name $Extension -ScriptBlock {
        Param($Name)
        &"$HOME\AppData\Local\Programs\Microsoft VS Code\bin\code.cmd" --install-extension $Name --force
        } -ArgumentList $Extension
    }
}

#endregion

#region Update help

$progParams.Status = "Updating Help. Some errors are to be expected."
$progParams.CurrentOperation = "Update-Help -Force"
$progParams.PercentComplete = $pct+=5
Write-Progress @progParams
#Write-Host 'Updating PowerShell help. Some errors are to be expected.' -ForegroundColor Yellow
Update-Help -Force

#endregion

#region Wait for end
$progParams.Status = "Waiting for $($jobs.count) background jobs to complete"
$progParams.CurrentOperation = "Wait-Job"
$progParams.PercentComplete = $pct+=2
Write-Progress @progParams
$jobs | Wait-Job | Select-Object Name,State

Write-Progress -Activity $progParams.Activity -Completed -Status "All tasks completed." -PercentComplete 100

$msg = @'

Refresh is complete. You might also want to install the following modules using Install-PSResource:

    Microsoft.Winget.Client
    Microsoft.PowerShell.SecretStore
    Microsoft.PowerShell.SecretManagement
    Platyps

And the following packages via Winget:

    Microsoft.PowerToys
    Microsoft.PowerShell

You will need to configure VSCode with your preferred extensions and settings or configure
it to synch your saved settings.

Please restart your PowerShell session.

'@

Write-Host $msg -ForegroundColor Green

#endregion

#EOF
