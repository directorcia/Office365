param(                        
    [switch]$debug = $false, ## if -debug parameter don't prompt for input
    [switch]$prompt = $false  ## if -prompt parameter used user prompted for input
)
<#CIAOPS

Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Read and report Identity authorization policy best practices using Graph requests
Source - 

Prerequisites = 1
1. Ensure the MS Graph module is installed

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

function entra-settings() {
    # https://learn.microsoft.com/en-us/graph/api/authorizationpolicy-get?view=graph-rest-1.0
    # Application Permissions = Policy.Read.All, Policy.ReadWrite.Authorization
    $policyUrl = "https://graph.microsoft.com/beta/policies/authorizationPolicy"
    write-host -ForegroundColor Gray -backgroundcolor blue "`nGet Entra Auth policy"
    write-host -ForegroundColor Gray -backgroundcolor blue "---------------------"
    $request = Invoke-MgGraphRequest -Uri $policyUrl -Method GET
    if ($request.value.blockMsolPowerShell -notmatch $bpsettings.blockMsolPowerShell) {
        write-host -foregroundcolor $errormessagecolor "- Block MSOL PowerShell =", $request.value.blockMsolPowerShell
    }
    else {
        write-host -foregroundcolor $processmessagecolor "- Block MSOL PowerShell =", $request.value.blockMsolPowerShell
    }
    if ($request.value.allowinvitesfrom -notmatch $bpsettings.allowinvitesfrom) {
        write-host -foregroundcolor $errormessagecolor "- Allow invites from =", $request.value.allowinvitesfrom
    }
    else {
        write-host -foregroundcolor $processmessagecolor "- Allow invites from =", $request.value.allowinvitesfrom
    }
    # https://learn.microsoft.com/en-us/entra/identity/users/directory-self-service-signup
    if ($request.value.allowedToSignUpEmailBasedSubscriptions -notmatch $bpsettings.allowedToSignUpEmailBasedSubscriptions) {
        write-host -foregroundcolor $errormessagecolor "- Allowed to sign up to email based subscriptions =", $request.value.allowedToSignUpEmailBasedSubscriptions
    }
    else {
        write-host -foregroundcolor $processmessagecolor "- Allowed to sign up to email based subscription =", $request.value.allowedToSignUpEmailBasedSubscriptions
    }
    # https://learn.microsoft.com/en-us/entra/identity/users/directory-self-service-signup
    if ($request.value.allowEmailVerifiedUsersToJoinOrganization -notmatch $bpsettings.allowEmailVerifiedUsersToJoinOrganization) {
        write-host -foregroundcolor $errormessagecolor "- Allow email verified users to join organization =", $request.value.allowEmailVerifiedUsersToJoinOrganization
    }
    else {
        write-host -foregroundcolor $processmessagecolor "- Allow email verified users to join organization =", $request.value.allowEmailVerifiedUsersToJoinOrganization
    }
    if ($request.value.allowedToUseSSPR -notmatch $bpsettings.allowedToUseSSPR) {
        write-host -foregroundcolor $errormessagecolor "- Allow Self Service Password Reset (SSPR) =", $request.value.allowedToUseSSPR
    }
    else {
        write-host -foregroundcolor $processmessagecolor "- Allow Self Service Password Reset (SSPR) =", $request.value.allowedToUseSSPR
    }
    if ($request.value.allowUserConsentForRiskyApps -notmatch $bpsettings.allowUserConsentForRiskyApps) {
        write-host -foregroundcolor $errormessagecolor "- Allow user consent for risky apps =", $request.value.allowUserConsentForRiskyApps
    }
    else {
        write-host -foregroundcolor $processmessagecolor "- Allow user consent for risk apps =", $request.value.allowUserConsentForRiskyApps
    }
    # This setting corresponds to the Restrict non-admin users from creating tenants setting in the User settings menu in the Microsoft Entra admin center.
    if ($request.value.defaultuserrolepermissions.allowedToCreateTenants -notmatch $bpsettings.defaultuserrolepermissions.allowedToCreateTenants) {
        write-host -foregroundcolor $errormessagecolor "- Users can create Entra tenants =", $request.value.defaultuserrolepermissions.allowedToCreateTenants 
    }
    else {
        write-host -foregroundcolor $processmessagecolor "- Users can create Entra tenants =", $request.value.defaultuserrolepermissions.allowedToCreateTenants 
    }
    # This setting corresponds to the Users can register applications setting in the User settings menu in the Microsoft Entra admin center.
    if ($request.value.defaultuserrolepermissions.allowedToCreateApps -notmatch $bpsettings.defaultuserrolepermissions.allowedToCreateApps) {
        write-host -foregroundcolor $errormessagecolor "- Users can create apps =", $request.value.defaultuserrolepermissions.allowedToCreateApps 
    }
    else {
        write-host -foregroundcolor $processmessagecolor "- Users can create apps =", $request.value.defaultuserrolepermissions.allowedToCreateApps
    }
    # This setting corresponds to the following menus in the Microsoft Entra admin center:
    #    The Users can create security groups in Microsoft Entra admin centers, API or PowerShell setting in the Group settings menu.
    #    Users can create security groups setting in the User settings menu.
    if ($request.value.defaultuserrolepermissions.allowedToCreateSecurityGroups -notmatch $bpsettings.defaultuserrolepermissions.allowedToCreateSecurityGroups) {
        write-host -foregroundcolor $errormessagecolor "- Users can create security groups =", $request.value.defaultuserrolepermissions.allowedToCreateSecurityGroups 
    }
    else {
        write-host -foregroundcolor $processmessagecolor "- Users can security group =", $request.value.defaultuserrolepermissions.allowedToCreateSecurityGroups
    }
    if ($request.value.defaultuserrolepermissions.allowedToReadOtherUsers -notmatch $bpsettings.defaultuserrolepermissions.allowedToReadOtherUsers) {
        write-host -foregroundcolor $processmessagecolor "- Users can read other users =", $request.value.defaultuserrolepermissions.allowedToReadOtherUsers
    }
    else {
        write-host -foregroundcolor $errormessagecolor "- Users can read other users =", $request.value.defaultuserrolepermissions.allowedToReadOtherUsers
    }
    if ($request.value.defaultuserrolepermissions.allowedToReadBitlockerKeysForOwnedDevice -notmatch $bpsettings.defaultuserrolepermissions.allowedToReadBitlockerKeysForOwnedDevice) {
        write-host -foregroundcolor $errormessagecolor "- Users can read Bitlocker Keys for their own devices =", $request.value.defaultuserrolepermissions.allowedToReadBitlockerKeysForOwnedDevice
    }
    else {
        write-host -foregroundcolor $processmessagecolor "- Users can read Bitlocker Keys for their own devices =", $request.value.defaultuserrolepermissions.allowedToReadBitlockerKeysForOwnedDevice
    }
    # https://learn.microsoft.com/en-us/graph/api/resources/defaultuserrolepermissions?view=graph-rest-1.0
    # permissionGrantPoliciesAssigned  
}

function entra-auth-methods() {
    $policyUrl = "https://graph.microsoft.com/beta/policies/authenticationMethodsPolicy"
    write-host -ForegroundColor Gray -backgroundcolor blue "`nGet Entra authentication methods policy"
    write-host -ForegroundColor Gray -backgroundcolor blue "---------------------------------------"
    $request = Invoke-MgGraphRequest -Uri $policyUrl -Method GET
    write-host "- Identity authentication methods"
    $i = 0
    foreach ($method in $request.authenticationMethodConfigurations.id) {
        write-host "    ", $method, "=", $request.authenticationMethodConfigurations.state[$i++]
    }
}

if ($debug) {
    # create a log file of process if option enabled
    write-host "Script activity logged at .\graph-bp-get.txt"
    start-transcript ".\graph-bp-get.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}

Clear-Host
write-host -foregroundcolor $systemmessagecolor "Graph Identity Authorization Best Practice get script - Started`n"

write-host -foregroundcolor $processmessagecolor "Connect to MS Graph"
connect-mggraph -scopes "Policy.ReadWrite.Authorization", "Policy.Read.All", "Policy.ReadWrite.AuthenticationMethod" | Out-Null
$graphcontext = Get-MgContext
write-host -foregroundcolor $processmessagecolor "Connected account =", $graphcontext.Account
if ($prompt) {
    do {
        $response = read-host -Prompt "`nIs this correct? [Y/N]"
    } until (-not [string]::isnullorempty($response))
    if ($response -ne "Y" -and $response -ne "y") {
        Disconnect-MgGraph | Out-Null
        write-host -foregroundcolor $warningmessagecolor "[001] Disconnected from current Graph environment. Re-run script to login to desired environment"
        exit 1
    }
}

write-host -ForegroundColor $processmessagecolor "Get Identity Authorization Best Practice policy from CIAOPS Best Practice repo"
$asrbpurl = "https://raw.githubusercontent.com/directorcia/bp/main/EntraID/authorization.json"
try {
    $query = invoke-webrequest -method GET -ContentType "application/json" -uri $asrbpurl -UseBasicParsing
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[003]", $_.Exception.Message
}

$bpsettings = $query.content | ConvertFrom-Json

entra-settings
entra-auth-methods

write-host -foregroundcolor $processmessagecolor "Disconnect any existing Graph sessions"

write-host -foregroundcolor $systemmessagecolor "`nGraph Best Practice get script - Finished"
if ($debug) {
    Stop-Transcript | Out-Null      
}
