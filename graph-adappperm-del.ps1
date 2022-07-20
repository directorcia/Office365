param(                        
    [switch]$debug = $false     ## if -debug parameter don't prompt for input
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Delete permissions from an Azure AD enterprise application
Source - https://github.com/directorcia/Office365/blob/master/graph-adappperm-del.ps1

Prerequisites = 1
1. Azure AD Module loaded

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

Clear-Host
if ($debug) {
    Start-transcript "..\graph-adappperm-del.txt" | Out-Null                                   ## Log file created in current directory that is overwritten on each run
}
write-host -foregroundcolor $systemmessagecolor "Script started`n"
write-host -foregroundcolor cyan -backgroundcolor DarkBlue ">>>>>> Copyright www.ciaops.com <<<<<<`n"
write-host "--- Script to delete app permissions from an Azure AD application in a tenant ---"

write-host -foregroundcolor $processmessagecolor "`nCheck for Azure AD PowerShell module"
if (get-module -listavailable -name AzureAD) {
    ## Has the AzureAD PowerShell module been loaded?
    write-host -foregroundcolor $processmessagecolor "Azure AD PowerShell Module found"
}
else {
    write-host -foregroundcolor $warningmessagecolor -backgroundcolor $errormessagecolor "Azure AD PowerShell Module not installed. Please install and re-run script`n"
    write-host "You can install the Azure AD Powershell module by:`n"
    write-host "    1. Launching an elevated Powershell console then,"
    write-host "    2. Running the command,'install-module AzureAD'.`n"
    Stop-Transcript | Out-Null                                                          ## Terminate transcription
    Pause                                                                               ## Pause to view error on screen
    exit 0                                                                              ## Terminate script 
}
$results = get-azureadserviceprincipal -All $true | sort-object displayname | Out-GridView -PassThru -title "Select Enterprise Application (Multiple selections permitted)"
foreach ($result in $results) {             # loop through all selected options
    write-host -foregroundcolor $processmessagecolor "Commencing",$result.displayname
    # Get Service Principal using objectId
    $sp = Get-AzureADServicePrincipal -ObjectId $results.ObjectId
    # Menu selection for USer or Admin consent types
    $consenttype = @()
    $consenttype += [PSCustomObject]@{
        Name = "Admin consent";
        type = "allprincipals"
    }
    $consenttype += [PSCustomObject]@{
        Name = "User consent";
        type = "principal"
    }
    $consentselects = $consenttype | Out-GridView -PassThru -title "Select Consent type (Multiple selections permitted)"

    foreach ($consentselect in $consentselects) {           # Loop through all selected options
        write-host -foregroundcolor $processmessagecolor "Commencing for",$consentselect.Name
        # Get all delegated permissions for the service principal
        $spOAuth2PermissionsGrants = Get-AzureADOAuth2PermissionGrant -All $true | Where-Object { $_.clientId -eq $sp.ObjectId }
        $info = $spOAuth2PermissionsGrants | Where-Object { $_.consenttype -eq $consentselect.type }
        if ($info) {            # if there are permissions set
            if ($consentselect.type -eq "principal") {  # user consent
                $usernames = @()
                foreach ($item in $info) {
                    $usernames += get-azureaduser -ObjectId $item.PrincipalId
                }
                $selectusers = $usernames | select-object Displayname, userprincipalname, objectid | sort-object Displayname | Out-GridView -PassThru -title "Select Consent type (Multiple selections permitted)"
                foreach ($selectuser in $selectusers) {       # Loop through all selected options
                    $infoscopes = $info | Where-Object { $_.principalid -eq $selectuser.ObjectId }
                    write-host -foregroundcolor $processmessagecolor "`n"$consentselect.name,"permissions for user",$selectuser.displayname
                    foreach ($infoscope in $infoscopes) {
                        write-host "`nResource ID =",$infoscope.resourceid
                        $assignments = $infoscope.scope -split " "
                        foreach ($assignment in $assignments) {
                            write-host "-",$assignment
                        }
                    }
                    write-host -foregroundcolor $processmessagecolor "`nSelect items to remove`n"
                    $removes = $infoscopes | Select-object scope, resourceid, objectid | Out-GridView -PassThru -title "Select permissions to delete (Multiple selections permitted)"
                    foreach ($remove in $removes) {
                        Remove-AzureADOAuth2PermissionGrant -ObjectId $remove.ObjectId
                        write-host -foregroundcolor $warningmessagecolor "Removed consent for",$remove.scope
                    }
                }
            } 
            elseif ($consentselect.type -eq "allprincipals") {      # Admin consent
                $infoscopes = $info | Where-Object { $_.principalid -eq $null}
                write-host -foregroundcolor $processmessagecolor $consentselect.name,"permissions"
                foreach ($infoscope in $infoscopes) {
                    write-host "`nResource ID =",$infoscope.resourceid
                    $assignments = $infoscope.scope -split " "
                    foreach ($assignment in $assignments) {
                        write-host "-",$assignment
                    }
                }
                write-host -foregroundcolor $processmessagecolor "`nSelect items to remove`n"
                $removes = $infoscopes | Select-object scope, resourceid, objectid | Out-GridView -PassThru -title "Select permissions to delete (Multiple selections permitted)"
                foreach ($remove in $removes) {
                    Remove-AzureADOAuth2PermissionGrant -ObjectId $remove.ObjectId
                    write-host -foregroundcolor $warningmessagecolor "Removed consent for",$remove.scope
                }
            }
        } else {
            write-host -foregroundcolor $warningmessagecolor "`nNo",$consentselect.name,"permissions found for" ,$results.displayname,"`n"
        }
    }
}

Write-Host -ForegroundColor $systemmessagecolor "`nScript Finished"

if ($debug) {
    Stop-transcript | Out-Null
}
