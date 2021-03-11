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

Function test-install($modulename) {
    try {
        $found = Get-InstalledModule -Name $modulename -erroraction Stop    
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
        write-host -foregroundcolor $warningmessagecolor -nonewline "    [Warning]"$modulename" module not found. "
        if ($prompt) {
            do {
                $result = Read-host -prompt "Install this module (Y/N)?"
            } until (-not [string]::isnullorempty($result))
            if ($result -eq 'Y' -or $result -eq 'y') {
                write-host -foregroundcolor $processmessagecolor "Installing module",$modulename
                install-Module -Name $modulename -Force -confirm:$false
            }
        } else {
            write-host -foregroundcolor $processmessagecolor "Installing module",$modulename
            install-Module -Name $modulename -Force
        }
    }
}

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Start Script`n"
write-host -ForegroundColor $processmessagecolor "Prompt to install missing modules =",$prompt"`n"

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    write-host -foregroundcolor $processmessagecolor "Update Azure AD module"
    test-install -modulename AzureAD
    write-host -foregroundcolor $processmessagecolor "Update Azure Information Protection module"
    test-install -modulename AIPService
    write-host -foregroundcolor $processmessagecolor "Update Teams Module"
    test-install -modulename MicrosoftTeams
    write-host -foregroundcolor $processmessagecolor "Update SharePoint Online module"
    test-install -modulename Microsoft.Online.SharePoint.PowerShell
    write-host -foregroundcolor $processmessagecolor "Update Microsoft Online module"
    test-install -modulename MSOnline
    write-host -foregroundcolor $processmessagecolor "Update PowerShellget module"
    test-install -modulename PowershellGet
    write-host -foregroundcolor $processmessagecolor "Update Exchange Online V2 module"
    test-install -modulename ExchangeOnlineManagement
    write-host -foregroundcolor $processmessagecolor "Update Azure module"
    test-install -modulename Az 
    write-host -foregroundcolor $processmessagecolor "Update SharePoint PnP module"
    test-install -modulename SharePointPnPPowerShellOnline
    write-host -foregroundcolor $processmessagecolor "Update Microsoft Graph module"
    test-install -modulename Microsoft.Graph 
    write-host -foregroundcolor $processmessagecolor "Update Windows Autopilot Module"
    ## will also update dependent AzureAD and Microsoft.Graph.Intune modules
    test-install -modulename WindowsAutoPilotIntune
    write-host -foregroundcolor $processmessagecolor "Centralised Add-in Deployment"
    test-install -modulename O365CentralizedAddInDeployment
}
Else {
    write-host -foregroundcolor $errormessagecolor "*** ERROR *** - Please re-run PowerShell environment as Administrator`n"
}
write-host -foregroundcolor $systemmessagecolor "Script completed`n"