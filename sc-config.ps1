<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Source - https://github.com/directorcia/Office365/blob/master/sc-config.ps1

Description - Configuration for Switch Connect enablement and Teams calling

Prerequisites = 2
1. Ensure SharePoint online PowerShell module installed or updated
2. Connect to Skype for Business
- For MFA = https://github.com/directorcia/Office365/blob/master/o365-connect-mfa-s4b.ps1
- For non-MFA = https://github.com/directorcia/Office365/blob/master/o365-connect-s4b.ps1

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
## Change these prior to running script 
$scidentity = "user@domain.com.au"              ## User to be enabled full email address
$scphone = "+61XXXXXXXXX"                       ## User phone number
$scfqdn = "sbcXXX.teams.switchconnect.com.au"   ## Full Switch connect provided domain

Clear-Host

write-host -foregroundcolor cyan "Script started`n"

New-CSOnlinePSTNGateway -FQDN $scfqdn -SipSignallingPort 5061 -MaxConcurrentSessions 10 -ForwardPai $false -ForwardCallHistory $true -Enabled $true

Read-Host -Prompt "Press [Enter] to continue"

Set-CsOnlinePstnUsage -Identity Global -Usage @{Add = "Australia"}

Read-Host -Prompt "Press [Enter] to continue"

New-CsOnlineVoiceRoute -Identity "AU-Emergency" -NumberPattern "^+000$" -OnlinePstnGatewayList $scfqdn -Priority 1 -OnlinePstnUsages "Australia"

Read-Host -Prompt "Press [Enter] to continue"

New-CsOnlineVoiceRoute -Identity "AU-Service" -NumberPattern "^\+61(1\d{2,8})$" -OnlinePstnGatewayList $scfqdn -Priority 1 -OnlinePstnUsages "Australia"

Read-Host -Prompt "Press [Enter] to continue"

New-CsOnlineVoiceRoute -Identity "AU-National" -NumberPattern "^\+61\d{9}$" -OnlinePstnGatewayList $scfqdn -Priority 1 -OnlinePstnUsages "Australia"

Read-Host -Prompt "Press [Enter] to continue"

New-CsOnlineVoiceRoute -Identity "AU-International" -NumberPattern "^\+(?!(61190))([1-9]\d{9,})$" -OnlinePstnGatewayList $scfqdn -Priority 1 -OnlinePstnUsages "Australia"

Read-Host -Prompt "Press [Enter] to continue"

New-CsOnlineVoiceRoutingPolicy "Australia" -OnlinePstnUsages "Australia"

Read-Host -Prompt "Press [Enter] to continue"

Set-CsUser -Identity $scidentity -EnterpriseVoiceEnabled $true -HostedVoiceMail $true -OnPremLineURI tel:$scphone

Read-Host -Prompt "Press [Enter] to continue"

Grant-CsOnlineVoiceRoutingPolicy -Identity $scidentity -PolicyName Australia

Read-Host -Prompt "Press [Enter] to continue"

Get-CsOnlineUser -Identity $scidentity | Format-List -Property FirstName, LastName, EnterpriseVoiceEnabled, HostedVoiceMail, LineURI, UsageLocation, UserPrincipalName, WindowsEmailAddress, SipAddress, OnPremLineURI, OnlineVoiceRoutingPolicy, TeamsCallingPolicy, dialplan, TeamsInteropPolicy

Read-Host -Prompt "Press [Enter] to continue"

write-host -foregroundcolor cyan "`nScript completed`n"