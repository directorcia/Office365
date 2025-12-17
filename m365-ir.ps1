#Verify Audit Log Status
# Connect to Exchange Online PowerShell
Connect-ExchangeOnline -UserPrincipalName admin@contoso.com
  
# Verify unified audit log is enabled
Get-AdminAuditLogConfig | Format-List UnifiedAuditLogIngestionEnabled

# Expected output: UnifiedAuditLogIngestionEnabled : True

# ---------------------------------------------
# Enable MAilbox Auditing (On by Default)
# Verify mailbox auditing is enabled organization-wide
Get-OrganizationConfig | Format-List AuditDisabled
    
# Expected output: AuditDisabled : False (meaning auditing IS enabled)

#---------------------------------------------
# Install Microsoft Graph PowerShell SDK
# Install Microsoft Graph PowerShell module
Install-Module -Name Microsoft.Graph -Scope CurrentUser

# Install specific modules for user management
Install-Module Microsoft.Graph.Authentication
Install-Module Microsoft.Graph.Users.Actions
     
# Verify installation
Get-Module -ListAvailable Microsoft.Graph*

# ---------------------------------------------
# Install Exchange Online Management module
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser

# Connect to Exchange Online
Connect-ExchangeOnline -UserPrincipalName admin@contoso.com

# Verify connection
Get-OrganizationConfig | Select-Object Name

# ---------------------------------------------
# Install Microsoft Entra IR PowerShell module
Install-Module -Name AzureADIncidentResponse

# ---------------------------------------------
# Identify the Suspicious Email
# Connect to Exchange Online
Connect-ExchangeOnline

# Search for messages from suspicious sender (last 7 days)
$suspiciousMessages = Get-MessageTrace -SenderAddress phishing@malicious.com `
    -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date)

$suspiciousMessages | Format-Table Received, SenderAddress, RecipientAddress, Subject, Status, Size

# Get detailed trace for a specific message (use actual MessageTraceId from above)
if ($suspiciousMessages) {
    $firstMessage = $suspiciousMessages[0]
    Get-MessageTraceDetail -MessageTraceId $firstMessage.MessageTraceId `
        -RecipientAddress $firstMessage.RecipientAddress
}

# Search for messages with specific subject line
# Note: Subject filtering must be done client-side as Get-MessageTrace doesn't support it
Get-MessageTrace -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) -PageSize 5000 |
    Where-Object {$_.Subject -like "*urgent invoice*"} |
    Format-Table Received, SenderAddress, RecipientAddress, Subject

# ---------------------------------------------
# Search Unified Audit Log
# Connect to Exchange Online PowerShell
Connect-ExchangeOnline

# Option 1: Track message delivery using Message Trace (more reliable for email tracking)
$messageId = "<message-id>@domain.com"
$traceResults = Get-MessageTrace -MessageId $messageId `
    -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date)

$traceResults | Format-Table Received, SenderAddress, RecipientAddress, Subject, Status, MessageId

# Option 2: Search Unified Audit Log for Send events (if you need audit data)
$auditResults = Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) `
    -Operations "Send" `
    -ResultSize 5000

# Parse and display results with error handling
$auditResults | ForEach-Object {
    $auditData = $_.AuditData | ConvertFrom-Json
    [PSCustomObject]@{
        Timestamp = $_.CreationDate
        User = $_.UserIds
        Sender = if ($auditData.PSObject.Properties.Name -contains 'From') { $auditData.From } else { "N/A" }
        Recipients = if ($auditData.PSObject.Properties.Name -contains 'To') { $auditData.To -join "; " } else { "N/A" }
        Subject = if ($auditData.PSObject.Properties.Name -contains 'Subject') { $auditData.Subject } else { "N/A" }
        ClientIP = if ($auditData.PSObject.Properties.Name -contains 'ClientIP') { $auditData.ClientIP } else { "N/A" }
    }
} | Format-Table -AutoSize

# Track phishing link clicks using Safe Links logs (requires Defender for Office 365)
Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) `
    -Operations "SafeLinksBlocked","SafeLinksClicked" `
    -ResultSize 5000 |
    Select-Object CreationDate, UserIds, Operations, ResultStatus |
    Format-Table -AutoSize

# ---------------------------------------------
# Check for User Interaction with Phishing Email
# Search for file downloads from suspicious emails
Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) `
    -Operations "FileDownloaded" `
    -UserIds "user@contoso.com" `
    -ResultSize 5000
     
# Search for URL clicks (ClickTracking events)
Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) `
    -RecordType "ThreatIntelligenceUrl" `
    -ResultSize 5000 |
    ForEach-Object {
    $data = $_.AuditData | ConvertFrom-Json
    [PSCustomObject]@{
        Time = $_.CreationDate
        User = $_.UserIds
        URL = $data.Url
        Action = $data.EventType
    }
} | Format-Table -AutoSize

# ---------------------------------------------
# Review Sign in LOgs
# Install and import Microsoft Graph PowerShell
Install-Module Microsoft.Graph.Reports -Scope CurrentUser
Import-Module Microsoft.Graph.Reports

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "AuditLog.Read.All", "Directory.Read.All"
    
# Get sign-in logs for specific user (last 7 days)
$userId = "user@contoso.com"
$startDate = (Get-Date).AddDays(-7).ToString("yyyy-MM-dd")
    
$signIns = Get-MgAuditLogSignIn -Filter "userPrincipalName eq '$userId' and createdDateTime ge $startDate" -All
    
# Display sign-in summary
$signIns | Select-Object CreatedDateTime, UserPrincipalName, AppDisplayName, IPAddress, 
    Location, DeviceDetail, Status | Format-Table -AutoSize
    
# Identify suspicious sign-ins (multiple countries)
$signIns | Group-Object -Property {$_.Location.CountryOrRegion} | 
Select-Object Name, Count | Sort-Object Count -Descending
     
# Find failed sign-ins
$signIns | Where-Object {$_.Status.ErrorCode -ne 0} | 
        Select-Object CreatedDateTime, UserPrincipalName, IPAddress, 
        @{Name='FailureReason';Expression={$_.Status.FailureReason}} | 
        Format-Table -AutoSize

# ---------------------------------------------
# Check for Suspicious Inbox Rules
# Connect to Exchange Online
Connect-ExchangeOnline
    
# Get all inbox rules for a user
Get-InboxRule -Mailbox user@contoso.com | 
    Format-Table Name, Enabled, Priority, ForwardTo, ForwardAsAttachmentTo, 
    RedirectTo, DeleteMessage, MoveToFolder -AutoSize
     
# Find rules that forward externally
Get-InboxRule -Mailbox user@contoso.com | 
    Where-Object {$_.ForwardTo -ne $null -or $_.ForwardAsAttachmentTo -ne $null -or $_.RedirectTo -ne $null} |
    Select-Object Name, Enabled, ForwardTo, ForwardAsAttachmentTo, RedirectTo
     
# Get all users with suspicious forwarding rules (org-wide check)
Get-Mailbox -ResultSize Unlimited | ForEach-Object {
    $rules = Get-InboxRule -Mailbox $_.PrimarySmtpAddress | 
    Where-Object {$_.ForwardTo -ne $null -or $_.ForwardAsAttachmentTo -ne $null}
    if ($rules) {
        [PSCustomObject]@{
            User = $_.PrimarySmtpAddress
            RuleName = $rules.Name -join "; "
            ForwardTo = $rules.ForwardTo -join "; "
        }
    }
}

# ---------------------------------------------
# Check if SMTP forwarding is configured on the mailbox
Get-Mailbox -Identity user@contoso.com | 
    Format-List ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward
    
# Check for external forwarding across all mailboxes
Get-Mailbox -ResultSize Unlimited | 
    Where-Object {$_.ForwardingSmtpAddress -ne $null} |
    Select-Object DisplayName, PrimarySmtpAddress, ForwardingSmtpAddress, 
    DeliverToMailboxAndForward | 
    Format-Table -AutoSize

# ---------------------------------------------
# Search mailbox audit log for specific user (last 90 days)
Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date) `
    -UserIds "user@contoso.com" `
    -Operations "New-InboxRule","Set-InboxRule","UpdateInboxRules" `
    -ResultSize 5000 |
    Select-Object CreationDate, UserIds, Operations, AuditData |
    Format-Table -AutoSize
    
# Search for mailbox permissions changes
Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date) `
    -UserIds "user@contoso.com" `
    -Operations "Add-MailboxPermission","Remove-MailboxPermission" `
    -ResultSize 5000
    
# Search for delegate additions
Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date) `
    -UserIds "user@contoso.com" `
    -Operations "Add-MailboxFolderPermission" `
    -ResultSize 5000

# ---------------------------------------------
# Search for sent emails from compromised account
Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) `
    -UserIds "user@contoso.com" `
    -Operations "Send" `
    -ResultSize 5000 |
     ForEach-Object {
        $data = $_.AuditData | ConvertFrom-Json
            [PSCustomObject]@{
                Timestamp = $_.CreationDate
               From = $data.From
               Recipients = $data.To -join "; "
               Subject = $data.Subject
               ClientIP = $data.ClientIP
             }
        } | Format-Table -AutoSize
     
# Use Content Search to find and review sent items
# (Requires Security & Compliance PowerShell)
Connect-IPPSSession
     
# Create content search for sent items
$searchName = "CompromisedAccount-SentItems-" + (Get-Date -Format "yyyyMMdd")
    New-ComplianceSearch -Name $searchName `
    -ExchangeLocation "user@contoso.com" `
    -ContentMatchQuery "kind:email AND sent:>=$(Get-Date -Format 'yyyy-MM-dd')"
    
# Start the search
Start-ComplianceSearch -Identity $searchName
     
# Check search status
Get-ComplianceSearch -Identity $searchName | Format-List Status,Items,Size
     
# Preview results (first 100 items)
Get-ComplianceSearchAction -Identity "$searchName_Preview" | 
Select-Object -ExpandProperty Results

# ---------------------------------------------
# Search for Sensitive Data Exposure
# Connect to Security & Compliance PowerShell
Connect-IPPSSession
     
# Search for specific sensitive information (e.g., credit card numbers)
$searchName = "DataBreach-Investigation-" + (Get-Date -Format "yyyyMMdd")
New-ComplianceSearch -Name $searchName `
    -ExchangeLocation All `
    -SharePointLocation All `
    -OneDriveLocation All `
    -ContentMatchQuery "SensitiveType:'Credit Card Number'"
    
# Start search
Start-ComplianceSearch -Identity $searchName
    
# Monitor progress
Get-ComplianceSearch -Identity $searchName | 
    Format-List Name,Status,Items,Size,SuccessResults
     
# Get DLP incidents for review
Get-DlpIncident -StartDate (Get-Date).AddDays(-30) |
    Select-Object PolicyName, IncidentTime, PolicyMatchCount, Severity |
    Format-Table -AutoSize

# ---------------------------------------------
# Export comprehensive audit log for incident investigation
$startDate = (Get-Date).AddDays(-30)
$endDate = Get-Date
$outputFile = "AuditLog_Incident_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
# Create array to store all results
$allResults = @()
    
# Search in batches (max 5000 per query)
$sessionId = [Guid]::NewGuid().ToString()
$resultSize = 5000
$iterations = 0
$moreData = $true
    
while ($moreData) {
    $results = Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate `
        -SessionId $sessionId `
        -SessionCommand ReturnLargeSet `
        -ResultSize $resultSize
    
    if ($results) {
        $allResults += $results
        $iterations++
        Write-Host "Retrieved $($results.Count) records (Total: $($allResults.Count))"
    } else {
        $moreData = $false
    }
}

# Export to CSV
$allResults | Select-Object CreationDate, UserIds, Operations, AuditData | 
    Export-Csv -Path $outputFile -NoTypeInformation

Write-Host "Exported $($allResults.Count) audit records to $outputFile"

# ---------------------------------------------
# Disable Account
# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.ReadWrite.All"
    
# Store user details
$userUPN = "compromised.user@contoso.com"
$user = Get-MgUser -Search "UserPrincipalName:'$userUPN'" -ConsistencyLevel Eventual
    
# Disable the account
Update-MgUser -UserId $user.Id -AccountEnabled $false

# Verify account is disabled
Get-MgUser -UserId $user.Id | Select-Object DisplayName, UserPrincipalName, AccountEnabled

# ---------------------------------------------
# Reset Pssword
# Generate strong random password
$newPassword = -join ((48..57) + (65..90) + (97..122) + (33..47) | 
    Get-Random -Count 16 | ForEach-Object {[char]$_})
     
# Reset password
   $passwordProfile = @{
   Password = $newPassword
   ForceChangePasswordNextSignIn = $true
}
    
Update-MgUser -UserId $user.Id -PasswordProfile $passwordProfile

Write-Host "Password reset for $userUPN"
Write-Host "New temporary password: $newPassword"
Write-Host "DO NOT send this password via email!"

# ---------------------------------------------

# Revoke All Active Sessions
# Set execution policy (if needed)
 Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
    
# Install required modules
Install-Module Microsoft.Graph.Users.Actions -Scope CurrentUser
 
# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.RevokeSessions.All"
     
# Revoke all sessions for compromised user
$userUPN = "compromised.user@contoso.com"
Revoke-MgUserSignInSession -UserId $userUPN
  
# Confirm revocation
Write-Host "All active sessions revoked for $userUPN"

# ---------------------------------------------

# Using Content Search and Purge
# Connect to Security & Compliance PowerShell
Connect-IPPSSession
    
# Create content search for malicious email
$searchName = "PhishingEmail-Removal-" + (Get-Date -Format "yyyyMMdd_HHmmss")
$messageId = "<malicious-message-id>@domain.com"
     
New-ComplianceSearch -Name $searchName `
    -ExchangeLocation All `
    -ContentMatchQuery "Subject:'Urgent Invoice' AND from:phishing@malicious.com"
    
# Start the search
Start-ComplianceSearch -Identity $searchName
    
# Wait for search to complete
do {
    Start-Sleep -Seconds 5
    $status = Get-ComplianceSearch -Identity $searchName
    Write-Host "Search status: $($status.Status) - Items found: $($status.Items)"
} while ($status.Status -ne "Completed")
     
# Review results before deletion
Get-ComplianceSearch -Identity $searchName | 
    Format-List Name,Items,Size,SuccessResults
    
# Purge the emails (soft delete - recoverable)
New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType SoftDelete
    
# For hard delete (PERMANENT - use with caution):
# New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType HardDelete

# Monitor purge progress
Get-ComplianceSearchAction -Identity "$searchName_Purge" | 
    Format-List Name,Status,Results

# ---------------------------------------------
# Remove Suspicious Inbox Rules
# Connect to Exchange Online
Connect-ExchangeOnline
    
# List all inbox rules for the user
$userUPN = "compromised.user@contoso.com"
$rules = Get-InboxRule -Mailbox $userUPN
   
# Display rules
$rules | Format-Table Name, Enabled, ForwardTo, RedirectTo, DeleteMessage
    
# Remove specific suspicious rule (if it exists)
$suspiciousRuleName = "Auto Forward External"
Remove-InboxRule -Mailbox $userUPN -Identity $suspiciousRuleName -Confirm:$false -ErrorAction SilentlyContinue
     
# Remove ALL inbox rules (if comprehensive cleanup needed)
Get-InboxRule -Mailbox $userUPN | ForEach-Object {
    Remove-InboxRule -Mailbox $userUPN -Identity $_.Name -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removed rule: $($_.Name)"
}
    
Write-Host "All inbox rules removed for $userUPN"

# ---------------------------------------------
# Remove SMTP Forwarding
# Check current forwarding configuration
Get-Mailbox -Identity $userUPN | 
    Format-List ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward
     
# Remove SMTP forwarding
Set-Mailbox -Identity $userUPN `
    -ForwardingSmtpAddress $null `
    -ForwardingAddress $null `
    -DeliverToMailboxAndForward $false
  
# Verify removal
Get-Mailbox -Identity $userUPN | 
    Format-List ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward

# ---------------------------------------------
# review and Remove MFA Methods
# Get user's authentication methods
Connect-MgGraph -Scopes "UserAuthenticationMethod.ReadWrite.All"
    
$userId = (Get-MgUser -Filter "userPrincipalName eq '$userUPN'").Id
    
# List all authentication methods
Get-MgUserAuthenticationMethod -UserId $userId | 
    Format-Table Id, AdditionalProperties
    
# Remove all authentication methods
Get-MgUserAuthenticationMethod -UserId $userId | ForEach-Object {
    Remove-MgUserAuthenticationMethod -UserId $userId -AuthenticationMethodId $_.Id -ErrorAction SilentlyContinue
    Write-Host "Removed authentication method: $($_.Id)"
}

# ---------------------------------------------
# Review Application Consents
# Get OAuth grants for user
Connect-MgGraph -Scopes "Directory.Read.All"
   
$grants = Get-MgUserOAuth2PermissionGrant -UserId $userId
    
# Display grants
$grants | ForEach-Object {
    $app = Get-MgServicePrincipal -ServicePrincipalId $_.ClientId
    [PSCustomObject]@{
        Application = $app.DisplayName
        ClientId = $_.ClientId
        Scope = $_.Scope
        ConsentType = $_.ConsentType
        StartTime = $_.StartTime
    }
} | Format-Table -AutoSize
    
# Revoke all OAuth permission grants
$grants | ForEach-Object {
    Remove-MgUserOAuth2PermissionGrant -OAuth2PermissionGrantId $_.Id -ErrorAction SilentlyContinue
    Write-Host "Revoked grant for: $($_.ClientId)"
}

# ---------------------------------------------
# Check for Unauthorized Delegates and Permissions
# Check mailbox delegates
$mailboxPerms = Get-MailboxPermission -Identity $userUPN | 
    Where-Object {$_.User -notlike "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false}

$mailboxPerms | Format-Table User, AccessRights, IsInherited
   
# Remove unauthorized delegates
$mailboxPerms | ForEach-Object {
    Remove-MailboxPermission -Identity $userUPN `
        -User $_.User `
        -AccessRights $_.AccessRights `
        -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removed mailbox permission for: $($_.User)"
}

# Check for send-as permissions
$recipientPerms = Get-RecipientPermission -Identity $userUPN | 
    Where-Object {$_.Trustee -ne "NT AUTHORITY\SELF"}

$recipientPerms | Format-Table Trustee, AccessRights

# Remove unauthorized recipient permissions
$recipientPerms | ForEach-Object {
    Remove-RecipientPermission -Identity $userUPN `
        -Trustee $_.Trustee `
        -AccessRights $_.AccessRights `
        -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removed recipient permission for: $($_.Trustee)"
}

# Check calendar permissions
Get-MailboxFolderPermission -Identity "$userUPN:\Calendar" |
    Format-Table User, AccessRights

# ---------------------------------------------
# Block Malicious Sender Domains
# Add sender to blocked senders list
Set-HostedContentFilterPolicy -Identity "Default" `
    -BlockedSenders @{Add="phishing@malicious.com"}
    
# Add entire domain to block list
Set-HostedContentFilterPolicy -Identity "Default" `
    -BlockedSenderDomains @{Add="malicious.com"}
    
# Create mail flow rule to block sender
New-TransportRule -Name "Block Phishing Domain - malicious.com" `
    -FromAddressContainsWords "malicious.com" `
    -DeleteMessage $true `
    -Comments "Created during incident response - $(Get-Date)"
    
# Verify rule was created
Get-TransportRule "Block Phishing Domain - malicious.com" | Format-List Name,State,FromAddressContainsWords

# ---------------------------------------------
# Check user's directory role memberships
$roleAssignments = Get-MgUserMemberOf -UserId $userId

$directoryRoles = $roleAssignments | Where-Object {
    $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.directoryRole'
} | ForEach-Object {
    [PSCustomObject]@{
        RoleId = $_.Id
        RoleName = $_.AdditionalProperties.displayName
    }
}

$directoryRoles | Format-Table -AutoSize

# Remove all discovered admin role assignments
$directoryRoles | ForEach-Object {
    Remove-MgDirectoryRoleMemberByRef -DirectoryRoleId $_.RoleId -DirectoryObjectId $userId -ErrorAction SilentlyContinue
    Write-Host "Removed role assignment: $($_.RoleName)"
}

# ---------------------------------------------
# Enable per-user MFA (legacy) for compromised user
# Install MSOnline module if needed
Install-Module -Name MSOnline -Scope CurrentUser -ErrorAction SilentlyContinue

# Connect to MSOnline service
Connect-MsolService

# Apply MFA requirement to compromised user
$mfa = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
$mfa.RelyingParty = "*"
$mfa.State = "Enforced"

Set-MsolUser -UserPrincipalName $userUPN -StrongAuthenticationRequirements $mfa -ErrorAction SilentlyContinue
Write-Host "Enforced MFA for: $userUPN"

# Verify MFA status
Get-MsolUser -UserPrincipalName $userUPN -ErrorAction SilentlyContinue | 
    Select-Object DisplayName, UserPrincipalName, 
    @{Name='MFAStatus';Expression={if ($_.StrongAuthenticationRequirements.State) { $_.StrongAuthenticationRequirements.State } else { "Not enforced" }}}

# ---------------------------------------------
# Generate strong password
Add-Type -AssemblyName 'System.Web'
$newPassword = [System.Web.Security.Membership]::GeneratePassword(20,5)
    
# Reset password and require change on next sign-in
$passwordProfile = @{
    Password = $newPassword
    ForceChangePasswordNextSignIn = $true
}
    
Connect-MgGraph -Scopes "User.ReadWrite.All" -ErrorAction SilentlyContinue
Update-MgUser -UserId $user.Id -PasswordProfile $passwordProfile -ErrorAction SilentlyContinue
     
# Securely communicate password to user (not via email)
Write-Host "New temporary password for $userUPN: $newPassword"
Write-Host "User must change password on next sign-in"
Write-Host "Communicate password via secure channel (phone, in-person, SMS)"

# ---------------------------------------------
# Re-enable the account
Update-MgUser -UserId $user.Id -AccountEnabled $true -ErrorAction SilentlyContinue
Write-Host "Account re-enabled for: $userUPN"

# Verify account is enabled
Get-MgUser -UserId $user.Id | 
    Select-Object DisplayName, UserPrincipalName, AccountEnabled

# ---------------------------------------------
# Monitor sign-ins for the recovered account (initial post-recovery check)
# Note: For continuous 24-hour monitoring, schedule this script to run hourly via Task Scheduler or a scheduled job
$monitorCheckStart = (Get-Date).AddHours(-1)

# Check sign-ins from the last hour
$recentSignIns = Get-MgAuditLogSignIn -Filter "userPrincipalName eq '$userUPN' and createdDateTime ge $($monitorCheckStart.ToString('yyyy-MM-ddTHH:mm:ssZ'))" -Top 50 -ErrorAction SilentlyContinue

if ($recentSignIns) {
    Write-Host "[$(Get-Date)] Post-recovery sign-in activity detected for $userUPN"
    $recentSignIns | Select-Object CreatedDateTime, IPAddress, @{Name='Country';Expression={$_.Location.CountryOrRegion}}, Status | 
        Format-Table -AutoSize
    
    # Define allowed countries (customize based on organization)
    $allowedCountries = @("United States", "Canada")
    
    # Alert on unusual locations
    $unusualLocations = $recentSignIns | 
        Where-Object {$_.Location.CountryOrRegion -notin $allowedCountries}
    
    if ($unusualLocations) {
        Write-Warning "ALERT: Unusual sign-in location detected for $userUPN!"
        $unusualLocations | Format-Table CreatedDateTime, IPAddress, 
            @{Name='Country';Expression={$_.Location.CountryOrRegion}}
    }
} else {
    Write-Host "No sign-in activity detected for $userUPN in the last hour"
}

# To enable continuous 24-hour monitoring, use the following scheduled task approach:
# Register-ScheduledJob -Name "M365IR-Monitor-$userUPN" -ScriptBlock { <# paste monitoring code here #> } -Trigger (New-JobTrigger -RepeatIndefinitely -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 1))

# ---------------------------------------------
# Check if user is blocked from sending email
$blockedStatus = Get-BlockedSenderAddress | Where-Object {$_.SenderAddress -eq $userUPN}

if ($blockedStatus) {
    Write-Host "User is blocked from sending, removing from blocked senders list..."
    Remove-BlockedSenderAddress -SenderAddress $userUPN -ErrorAction SilentlyContinue
    Write-Host "Removal submitted for: $userUPN"
} else {
    Write-Host "User is not on the blocked senders list"
}

# Verify removal
$verifyBlocked = Get-BlockedSenderAddress | Where-Object {$_.SenderAddress -eq $userUPN}
if ($verifyBlocked) {
    Write-Warning "User still appears in blocked senders list"
} else {
    Write-Host "User successfully removed from blocked senders list"
}

# ---------------------------------------------
# Send test email from recovered account (using Microsoft Graph)
# Define test email recipient (customize as needed)
$testRecipient = "admin@contoso.com"

try {
    # Create email message body
    $emailMessage = @{
        message = @{
            subject = "Email Flow Test - Account Recovery - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            body = @{
                contentType = "HTML"
                content = "This is a test email sent after account recovery to verify email flow functionality."
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $testRecipient
                    }
                }
            )
        }
        saveToSentItems = "true"
    }
    
    # Send email on behalf of the recovered user
    Send-MgUserMail -UserId $user.Id -BodyParameter $emailMessage -ErrorAction SilentlyContinue
    Write-Host "Test email sent from $userUPN to $testRecipient"
} catch {
    Write-Warning "Failed to send test email: $_"
}

# Check message trace for sent email (may take a few minutes to appear)
Write-Host "Waiting 30 seconds for message trace to reflect sent email..."
Start-Sleep -Seconds 30

$sentMessages = Get-MessageTrace -SenderAddress $userUPN -StartDate (Get-Date).AddHours(-1) -EndDate (Get-Date) -ErrorAction SilentlyContinue

if ($sentMessages) {
    Write-Host "Message trace results:"
    $sentMessages | Format-Table Received, SenderAddress, RecipientAddress, Subject, Status, Size
} else {
    Write-Host "No messages found in trace yet (may take several minutes to appear)"
}

# ---------------------------------------------
# Review and Restore Mailbox Settings
Write-Host "=== Post-Recovery Mailbox Verification ===" -ForegroundColor Cyan

# Verify forwarding is disabled
$forwardingConfig = Get-Mailbox -Identity $userUPN -ErrorAction SilentlyContinue
if ($forwardingConfig) {
    if ($forwardingConfig.ForwardingSmtpAddress -or $forwardingConfig.ForwardingAddress) {
        Write-Warning "WARNING: Email forwarding is still configured!"
        $forwardingConfig | Format-List ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward
    } else {
        Write-Host "[✓] Email forwarding is disabled" -ForegroundColor Green
    }
} else {
    Write-Warning "Could not retrieve mailbox configuration"
}

# Verify no inbox rules exist
$inboxRules = Get-InboxRule -Mailbox $userUPN -ErrorAction SilentlyContinue
if ($inboxRules) {
    Write-Warning "WARNING: Inbox rules still exist - review the following:"
    $inboxRules | Format-Table Name, Enabled, Priority
} else {
    Write-Host "[✓] No inbox rules configured" -ForegroundColor Green
}

# Check mailbox permissions are correct (should only show SELF permissions)
$unauthorizedPerms = Get-MailboxPermission -Identity $userUPN -ErrorAction SilentlyContinue | 
    Where-Object {$_.User -notlike "NT AUTHORITY\SELF" -and $_.User -notlike "SYSTEM"}
    
if ($unauthorizedPerms) {
    Write-Warning "WARNING: Unauthorized mailbox permissions detected - review the following:"
    $unauthorizedPerms | Format-Table User, AccessRights
} else {
    Write-Host "[✓] No unauthorized mailbox permissions" -ForegroundColor Green
}

# Verify mailbox size and quotas
$mailboxStats = Get-MailboxStatistics -Identity $userUPN -ErrorAction SilentlyContinue
if ($mailboxStats) {
    Write-Host "`nMailbox Statistics:"
    $mailboxStats | Format-List DisplayName, ItemCount, TotalItemSize, LastLogonTime
} else {
    Write-Warning "Could not retrieve mailbox statistics"
}

Write-Host "`n=== Verification Complete ===" -ForegroundColor Cyan

# ---------------------------------------------
# Comprehensive security check script
function Test-AccountSecurity {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
        
    $results = @()
    Write-Host "Running security checks for $UserPrincipalName..." -ForegroundColor Yellow
    
    # Check MFA (using Microsoft Graph instead of deprecated MSOnline)
    try {
        $mfaUser = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'" -ErrorAction Stop
        if ($mfaUser) {
            # Check if user has any authentication methods registered
            $authMethods = @(Get-MgUserAuthenticationMethod -UserId $mfaUser.Id -ErrorAction SilentlyContinue)
            $hasMfa = $authMethods | Where-Object { 
                $_.AdditionalProperties['@odata.type'] -in @(
                    '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod',
                    '#microsoft.graph.phoneAuthenticationMethod',
                    '#microsoft.graph.fido2AuthenticationMethod',
                    '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod'
                )
            }
            $mfaStatus = if($hasMfa) { "Enabled" } else { "Not Configured" }
            $results += [PSCustomObject]@{
                Check = "MFA Enabled"
                Status = if($mfaStatus -eq "Enabled") {"✅ Pass"} else {"❌ Fail"}
                Detail = $mfaStatus
            }
        } else {
            throw "User not found"
        }
    } catch {
        Write-Warning "Could not check MFA status: $_"
        $results += [PSCustomObject]@{
            Check = "MFA Enabled"
            Status = "⚠️ Error"
            Detail = "Unable to retrieve MFA status"
        }
    }
    
    # Check forwarding rules
    try {
        $rules = @(Get-InboxRule -Mailbox $UserPrincipalName -ErrorAction Stop)
        $ruleCount = @($rules).Count
        $results += [PSCustomObject]@{
            Check = "No Forwarding Rules"
            Status = if($ruleCount -eq 0) {"✅ Pass"} else {"❌ Fail"}
            Detail = "$ruleCount rules found"
        }
    } catch {
        Write-Warning "Could not check inbox rules: $_"
        $results += [PSCustomObject]@{
            Check = "No Forwarding Rules"
            Status = "⚠️ Error"
            Detail = "Unable to retrieve inbox rules"
        }
    }
   
    # Check SMTP forwarding
    try {
        $mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
        $hasForwarding = $mailbox.ForwardingSmtpAddress -or $mailbox.ForwardingAddress
        $forwardingDetail = if($hasForwarding) { $mailbox.ForwardingSmtpAddress } else { "None" }
        $results += [PSCustomObject]@{
            Check = "No SMTP Forwarding"
            Status = if(-not $hasForwarding) {"✅ Pass"} else {"❌ Fail"}
            Detail = $forwardingDetail
        }
    } catch {
        Write-Warning "Could not check SMTP forwarding: $_"
        $results += [PSCustomObject]@{
            Check = "No SMTP Forwarding"
            Status = "⚠️ Error"
            Detail = "Unable to retrieve forwarding config"
        }
    }
         
    # Check delegates
    try {
        $delegates = @(Get-MailboxPermission -Identity $UserPrincipalName -ErrorAction Stop | 
            Where-Object {$_.User -notlike "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false})
        $delegateCount = @($delegates).Count
        $results += [PSCustomObject]@{
            Check = "No Unauthorized Delegates"
            Status = if($delegateCount -eq 0) {"✅ Pass"} else {"⚠️ Review"}
            Detail = "$delegateCount delegates found"
        }
    } catch {
        Write-Warning "Could not check mailbox permissions: $_"
        $results += [PSCustomObject]@{
            Check = "No Unauthorized Delegates"
            Status = "⚠️ Error"
            Detail = "Unable to retrieve permissions"
        }
    }
    
    Write-Host "`n=== Security Check Results ===" -ForegroundColor Cyan
    return $results | Format-Table -AutoSize
}
    
# Run security check
Test-AccountSecurity -UserPrincipalName $userUPN

# ---------------------------------------------
# Implement Enhanced MOnitoring
# Create alert policy for recovered account
Connect-IPPSSession

$alertNotify = "security@contoso.com"

try {
    New-ProtectionAlert -Name "Recovered Account Monitoring - $userUPN" `
        -Category "ThreatManagement" `
        -NotifyUser $alertNotify `
        -ThreatType @("Malware","Phish","Spam") `
        -AggregationType "SimpleAggregation" `
        -AlertFor @($userUPN)
    Write-Host "Protection alert created for $userUPN and notifications sent to $alertNotify"
} catch {
    Write-Warning "Failed to create protection alert: $_"
}

# Monitor for specific activities
$activities = @(
    "New-InboxRule",
    "Set-InboxRule",
    "Set-Mailbox",
    "Add-MailboxPermission",
    "Set-TransportRule"
)
     
# Daily monitoring script (run via scheduled task)
$monitorStart = (Get-Date).AddHours(-24)
foreach ($activity in $activities) {
    $events = Search-UnifiedAuditLog -StartDate $monitorStart -EndDate (Get-Date) `
        -UserIds $userUPN `
        -Operations $activity `
        -ResultSize 5000 `
        -ErrorAction SilentlyContinue
        
    if ($events) {
        Write-Warning "Activity detected: $activity"
        $events | Select-Object CreationDate, Operation, ResultStatus | 
            Format-Table -AutoSize
            
        # Send alert email (requires existing Connect-MgGraph session with Mail.Send scope)
        $alertSender = "alerts@contoso.com"
        $alertSubject = "Alert: Activity on Recovered Account $userUPN"
        $alertBody = "Suspicious activity detected: $activity on $(Get-Date)"
        $alertMessage = @{
            message = @{
                subject = $alertSubject
                body = @{
                    contentType = "Text"
                    content = $alertBody
                }
                toRecipients = @(
                    @{ emailAddress = @{ address = $alertNotify } }
                )
            }
            saveToSentItems = "false"
        }

        try {
            Send-MgUserMail -UserId $alertSender -BodyParameter $alertMessage -ErrorAction Stop
            Write-Host "Alert email sent to $alertNotify for activity $activity"
        } catch {
            Write-Warning "Failed to send alert email: $_"
        }
    }
}

# ----------------------------------------------
# Generate incident timeline from audit logs
$incidentStart = Get-Date "2024-11-25 08:00:00"
$incidentEnd = Get-Date "2024-11-26 12:00:00"
$affectedUser = "user@contoso.com"
    
# Collect key events
$timeline = @()
    
# Get sign-in events
$signIns = Search-UnifiedAuditLog -StartDate $incidentStart -EndDate $incidentEnd `
    -UserIds $affectedUser `
    -Operations "UserLoggedIn" `
    -ResultSize 5000
    
foreach ($event in $signIns) {
    $data = $event.AuditData | ConvertFrom-Json
    $timeline += [PSCustomObject]@{
        Timestamp = $event.CreationDate
        EventType = "Sign-In"
        Description = "User logged in from $($data.ClientIP) - $($data.City), $($data.CountryCode)"
        Source = "Entra ID"
        Evidence = $event.ResultIndex
    }
}
     
# Get mailbox rule changes
$ruleChanges = Search-UnifiedAuditLog -StartDate $incidentStart -EndDate $incidentEnd `
    -UserIds $affectedUser `
    -Operations "New-InboxRule","Set-InboxRule","UpdateInboxRules" `
    -ResultSize 5000
    
foreach ($event in $ruleChanges) {
    $timeline += [PSCustomObject]@{
        Timestamp = $event.CreationDate
        EventType = "Mailbox Rule"
        Description = "Inbox rule modified: $($event.Operation)"
        Source = "Exchange Online"
        Evidence = $event.ResultIndex
    }
}
     
# Get email sending activity
$sentEmails = Search-UnifiedAuditLog -StartDate $incidentStart -EndDate $incidentEnd `
    -UserIds $affectedUser `
    -Operations "Send" `
    -ResultSize 5000
    
foreach ($event in $sentEmails) {
    $data = $event.AuditData | ConvertFrom-Json
    $timeline += [PSCustomObject]@{
    Timestamp = $event.CreationDate
        EventType = "Email Sent"
        Description = "Email sent to $($data.To -join ', ') - Subject: $($data.Subject)"
        Source = "Exchange Online"
        Evidence = $event.ResultIndex
    }
}

# Sort timeline chronologically
$timeline = $timeline | Sort-Object Timestamp

# Export to CSV for reporting
$timeline | Export-Csv "Incident_Timeline_$(Get-Date -Format 'yyyyMMdd').csv" -NoTypeInformation

# Display formatted timeline
$timeline | Format-Table Timestamp, EventType, Description -AutoSize

# ---------------------------------------------
# Connect to Exchange Online
Connect-ExchangeOnline -UserPrincipalName admin@contoso.com
  
# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.ReadWrite.All","AuditLog.Read.All"
  
# Connect to Security & Compliance
Connect-IPPSSession -UserPrincipalName admin@contoso.com

# Connect to MSOnline (legacy MFA)
Connect-MsolService

# ---------------------------------------------
# Quick Investigation Commands
# Check audit log status
Get-AdminAuditLogConfig | Format-List UnifiedAuditLogIngestionEnabled
   
# Search audit logs (last 7 days) for the recovered user
Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) -UserIds $userUPN -ResultSize 5000
   
# Check recent sign-ins for the recovered user (last 7 days)
$signInFilter = "userPrincipalName eq '$userUPN' and createdDateTime ge $(Get-Date).AddDays(-7).ToString('yyyy-MM-ddTHH:mm:ssZ')"
Get-MgAuditLogSignIn -Filter $signInFilter -Top 50

# ---------------------------------------------
# Microsoft Documentation
https://learn.microsoft.com/en-us/security/operations/incident-response-playbooks
https://learn.microsoft.com/en-us/security/operations/incident-response-planning
https://learn.microsoft.com/en-us/purview/audit-search
https://learn.microsoft.com/en-us/security/operations/incident-response-playbook-phishing
https://learn.microsoft.com/en-us/defender-office-365/threat-explorer-real-time-detections-about
https://learn.microsoft.com/en-us/defender-office-365/responding-to-a-compromised-email-account
https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/search-unifiedauditlog?view=exchange-ps
