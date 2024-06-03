param(                         ## if no parameters used then don't prompt to install missing modules, just install them
    [switch]$prompt = $false   ## if -prompt used then prompt to install missing modules
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Update all the relevant Microsoft Online PowerShell modules

Source - https://github.com/directorcia/Office365/blob/master/o365-update.ps1

## Prerequisites = 1
## 1. Run PowerShell environment as an administrator

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

function test-package($packagename) {
    try {
        $found = Get-PackageProvider -Name $packagename -erroraction SilentlyContinue    
    }
    catch {
        $found = $false
    }
    if ($found) {          ## If module exists then update
        #get version of the module (selects the first if there are more versions installed)
        $version = (Get-PackageProvider -name $packagename) | Sort-Object Version -Descending  | Select-Object Version -First 1
        #get version of the module in psgallery
        $psgalleryversion = Find-PackageProvider -Name $packagename | Sort-Object Version -Descending | Select-Object Version -First 1
        #convert to string for comparison
        $stringver = $version | Select-Object @{n='Version'; e={$_.Version -as [string]}}
        $a = $stringver | Select-Object version -ExpandProperty version
        #convert to string for comparison
        $onlinever = $psgalleryversion | Select-Object @{n='Version'; e={$_.Version -as [string]}}
        $b = $onlinever | Select-Object Version -ExpandProperty Version
        #version compare
        if ([version]"$a" -ge [version]"$b") {
            Write-Host -foregroundcolor $processmessagecolor "    Local package $a greater or equal to Gallery package $b"
            write-host -foregroundcolor $processmessagecolor "    No update required`n"
        }
        else {
            Write-Host -foregroundcolor $warningmessagecolor "    Local package $a lower version than Gallery package $b"
            write-host -foregroundcolor $warningmessagecolor "    Will be updated"
            update-packageprovider -name $packagename -force -confirm:$false
            Write-Host
        }
    }
    else {                                                      ## If module doesn't exist then prompt to update
        write-host -foregroundcolor $warningmessagecolor -nonewline "    [Warning]"$pacakgename" package not found.`n"
        if ($prompt) {
            do {
                $result = Read-host -prompt "Install this package (Y/N)?"
            } until (-not [string]::isnullorempty($result))
            if ($result -eq 'Y' -or $result -eq 'y') {
                write-host -foregroundcolor $processmessagecolor "Installing package",$packagename"`n"
                Install-PackageProvider -Name $packagename -Force -confirm:$false
            }
        } else {
            write-host -foregroundcolor $processmessagecolor "Installing package",$packagename"`n"
            Install-PackageProvider -Name $packagename -Force -confirm:$false
        }
    }
}


Function test-install($modulename) {
    try {
        $found = Get-InstalledModule -Name $modulename -erroraction SilentlyContinue    
    }
    catch {
        $found = $false
    }
    if ($found) {          ## If module exists then update
        #get version of the module (selects the first if there are more versions installed)
        $version = (Get-InstalledModule -name $modulename) | Sort-Object Version -Descending  | Select-Object Version -First 1
        #get version of the module in psgallery
        $psgalleryversion = Find-Module -Name $modulename | Sort-Object Version -Descending | Select-Object Version -First 1
        #convert to string for comparison
        $stringver = $version | Select-Object @{n='ModuleVersion'; e={$_.Version -as [string]}}
        $a = $stringver | Select-Object Moduleversion -ExpandProperty Moduleversion
        #convert to string for comparison
        $onlinever = $psgalleryversion | Select-Object @{n='OnlineVersion'; e={$_.Version -as [string]}}
        $b = $onlinever | Select-Object OnlineVersion -ExpandProperty OnlineVersion
        #version compare
        if ([version]"$a" -ge [version]"$b") {
            Write-Host -foregroundcolor $processmessagecolor "    Local module $a greater or equal to Gallery module $b"
            write-host -foregroundcolor $processmessagecolor "    No update required`n"
        }
        else {
            Write-Host -foregroundcolor $warningmessagecolor "    Local module $a lower version than Gallery module $b"
            write-host -foregroundcolor $warningmessagecolor "    Will be updated"
            update-module -name $modulename -force -confirm:$false
            Write-Host
        }
    }
    else {                                                      ## If module doesn't exist then prompt to update
        write-host -foregroundcolor $warningmessagecolor -nonewline "    [Warning]"$modulename" module not found.`n"
        if ($prompt) {
            do {
                $result = Read-host -prompt "    Install this module (Y/N)?"
            } until (-not [string]::isnullorempty($result))
            if ($result -eq 'Y' -or $result -eq 'y') {
                write-host -foregroundcolor $processmessagecolor "    Installing module",$modulename"`n"
                install-Module -Name $modulename -Force -confirm:$false -allowclobber
            }
        } else {
            write-host -foregroundcolor $processmessagecolor "    Installing module",$modulename"`n"
            install-Module -Name $modulename -Force -confirm:$false -allowclobber
        }
    }
}

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Start Script`n"
write-host -ForegroundColor $processmessagecolor "Prompt to install missing modules =",$prompt"`n"

$ps = $PSVersionTable.PSVersion
Write-host -foregroundcolor $processmessagecolor "Detected supported PowerShell version: $($ps.Major).$($ps.Minor)`n"
if ($ps.Major -lt 7) {
    $modulecount = 16       
} else {
    $modulecount = 15       # NUGET is not supported in PowerShell 7
}

$counter = 0

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    if ($ps.Major -lt 7) {
        ++$counter
        write-host -foregroundcolor $processmessagecolor "($($counter) of $($modulecount)) Update NuGet provider"
        test-package -packagename NuGet
    }
    ++$counter
    write-host -foregroundcolor $processmessagecolor "($($counter) of $($modulecount)) Update Azure AD module"
    test-install -modulename AzureAD
    ++$counter
    write-host -foregroundcolor $processmessagecolor "($($counter) of $($modulecount)) Update Azure Information Protection module"
    $aadrmcheck = get-module -listavailable -name aadrm
    if ($aadrmcheck) {
        write-host -foregroundcolor $warningmessagecolor "    [Warning] Older module Azure AD Rights management module (AADRM) is still installed"
        write-host -foregroundcolor $processmessagecolor "    Uninstalling AADRM module as support ended July 15, 2020 "
        uninstall-module aadrm -allversions -force -confirm:$false
        write-host -foregroundcolor $processmessagecolor "    Now Azure Information Protection module will now be installed"
    }
    test-install -modulename AIPService
    ++$counter
    write-host -foregroundcolor $processmessagecolor "($($counter) of $($modulecount)) Update Teams Module"
    test-install -modulename MicrosoftTeams
    ++$counter
    write-host -foregroundcolor $processmessagecolor "($($counter) of $($modulecount)) Update SharePoint Online module"
    test-install -modulename Microsoft.Online.SharePoint.PowerShell
    ++$counter
    write-host -foregroundcolor $processmessagecolor "($($counter) of $($modulecount)) Update Microsoft Online module"
    test-install -modulename MSOnline
    ++$counter
    write-host -foregroundcolor $processmessagecolor "($($counter) of $($modulecount)) Update PowerShellGet module"
    test-install -modulename PowershellGet
    ++$counter
    write-host -foregroundcolor $processmessagecolor "($($counter) of $($modulecount)) Update Exchange Online module"
    test-install -modulename ExchangeOnlineManagement
    ++$counter
    write-host -foregroundcolor $processmessagecolor "($($counter) of $($modulecount)) Update Azure module"
    test-install -modulename Az 
    ++$counter
    write-host -foregroundcolor $processmessagecolor "($($counter) of $($modulecount)) Update SharePoint PnP module"
    $pnpcheck = get-module -listavailable -name SharePointPnPPowerShellOnline
    if ($pnpcheck) {
        write-host -foregroundcolor $warningmessagecolor "    [Warning] Older SharePoint PnP module is still installed"
        write-host -foregroundcolor $processmessagecolor "    Uninstalling older SharePoint PnP module"
        uninstall-module SharePointPnPPowerShellOnline -allversions -force -confirm:$false
        write-host -foregroundcolor $processmessagecolor "    New SharePoint PnP module will now be installed"
    }
    test-install -modulename PnP.PowerShell
    ++$counter
    write-host -foregroundcolor $processmessagecolor "($($counter) of $($modulecount)) Update Microsoft Graph module"
    test-install -modulename Microsoft.Graph 
    ++$counter
    write-host -foregroundcolor $processmessagecolor "($($counter) of $($modulecount)) Update Windows Autopilot Module"
    ## will also update dependent AzureAD and Microsoft.Graph.Intune modules
    test-install -modulename WindowsAutoPilotIntune
    ++$counter
    write-host -foregroundcolor $processmessagecolor "($($counter) of $($modulecount)) Centralised Add-in Deployment"
    test-install -modulename O365CentralizedAddInDeployment
    ++$counter
    write-host -foregroundcolor $processmessagecolor "($($counter) of $($modulecount)) PowerApps"
    test-install -modulename Microsoft.PowerApps.PowerShell
    ++$counter
    write-host -foregroundcolor $processmessagecolor "($($counter) of $($modulecount)) PowerApps Administration module"
    test-install -modulename Microsoft.PowerApps.Administration.PowerShell
    ++$counter
    write-host -foregroundcolor $processmessagecolor "($($counter) of $($modulecount)) Microsoft 365 Commerce module"
    test-install -modulename MSCommerce
}
Else {
    write-host -foregroundcolor $errormessagecolor "*** ERROR *** - Please re-run PowerShell environment as Administrator`n"
}
write-host -foregroundcolor $systemmessagecolor "`nScript completed"