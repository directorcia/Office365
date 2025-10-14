<#
.SYNOPSIS
    Ultra-deep diagnostic to find why a mailbox name/alias is blocked
    
.DESCRIPTION
    This script performs an exhaustive search to find ANY occurrence of a name/alias
    in Exchange Online and Azure AD that could be blocking mailbox creation.
    
    It checks for conflicts in:
    - Display names, Names, Aliases (mailNickname)
    - Soft-deleted mailboxes
    - Azure AD user accounts
    - Distribution groups and Microsoft 365 Groups
    - Mail contacts, mail users, remote mailboxes
    - System mailboxes and public folder mailboxes
    
.PARAMETER SearchName
    Optional. The name/alias to search for. If not provided, you'll be prompted to enter it.
    
.PARAMETER DomainName
    Optional. Your organization's domain name (e.g., 'contoso.com'). If not provided, you'll be prompted.
    Used for generating suggested email addresses and mailbox creation commands.
    
.EXAMPLE
    .\find-name-conflict.ps1
    Interactive mode - prompts for both name and domain
    
.EXAMPLE
    .\find-name-conflict.ps1 -SearchName "sales" -DomainName "contoso.com"
    
.EXAMPLE
    .\find-name-conflict.ps1 -SearchName "sales"
    Prompts only for domain name
    
.NOTES
    Author: Created for CIAOPS
    Date: 2025-10-14
    Version: 2.0
    
.LINK
    https://github.com/directorcia/Office365/wiki/Find-Name-Conflict-%E2%80%90-Shared-Mailbox-Diagnostic-Tool
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$SearchName,
    
    [Parameter(Mandatory=$false)]
    [string]$DomainName
)

Write-Host "`n╔════════════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║                    SHARED MAILBOX CONFLICT DIAGNOSTIC                         ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

# Prompt for search name if not provided
if ([string]::IsNullOrWhiteSpace($SearchName)) {
    Write-Host "`n"
    Write-Host "This script will help you troubleshoot why a mailbox name cannot be created." -ForegroundColor Yellow
    Write-Host "It searches for conflicts across Exchange Online and Azure AD.`n" -ForegroundColor Yellow
    
    Write-Host "Common error: " -ForegroundColor White -NoNewline
    Write-Host """The name is already being used. Please try another name.""`n" -ForegroundColor Red
    
    do {
        Write-Host "Enter the name/alias you're trying to create " -ForegroundColor Cyan -NoNewline
        Write-Host "(e.g., 'treasurer', 'sales', 'info'): " -ForegroundColor White -NoNewline
        $SearchName = Read-Host
        
        if ([string]::IsNullOrWhiteSpace($SearchName)) {
            Write-Host "  ✗ Name cannot be empty. Please try again.`n" -ForegroundColor Red
        }
    } while ([string]::IsNullOrWhiteSpace($SearchName))
    
    Write-Host "`n"
}

# Prompt for domain name if not provided
if ([string]::IsNullOrWhiteSpace($DomainName)) {
    Write-Host "Enter your domain name " -ForegroundColor Cyan -NoNewline
    Write-Host "(e.g., 'contoso.com', 'company.org'): " -ForegroundColor White -NoNewline
    $DomainName = Read-Host
    
    # If still empty, try to get from Exchange Online
    if ([string]::IsNullOrWhiteSpace($DomainName)) {
        try {
            $defaultDomain = Get-AcceptedDomain -ErrorAction SilentlyContinue | Where-Object { $_.Default -eq $true } | Select-Object -ExpandProperty DomainName
            if ($defaultDomain) {
                $DomainName = $defaultDomain
                Write-Host "  Using default domain from Exchange Online: $DomainName" -ForegroundColor Green
            } else {
                $DomainName = "yourdomain.com"
                Write-Host "  ⚠ Could not determine domain. Using placeholder: $DomainName" -ForegroundColor Yellow
            }
        } catch {
            $DomainName = "yourdomain.com"
            Write-Host "  ⚠ Could not determine domain. Using placeholder: $DomainName" -ForegroundColor Yellow
        }
    }
    
    Write-Host "`n"
}

Write-Host "`nSearching for: '$SearchName' (case-insensitive)" -ForegroundColor Yellow
Write-Host "Domain: $DomainName" -ForegroundColor Yellow
Write-Host "This may take several minutes...`n" -ForegroundColor Gray

$findings = @()
$checkNumber = 0

function Write-Check {
    param([string]$Message)
    $script:checkNumber++
    Write-Host "`n[$script:checkNumber] " -ForegroundColor Cyan -NoNewline
    Write-Host $Message -ForegroundColor White
    Write-Host ("─" * 80) -ForegroundColor DarkGray
}

function Write-Finding {
    param(
        [string]$Type,
        [string]$Details,
        [object]$Object
    )
    Write-Host "  ► FOUND: " -ForegroundColor Red -NoNewline
    Write-Host "$Type - $Details" -ForegroundColor Yellow
    
    $script:findings += [PSCustomObject]@{
        Type = $Type
        Details = $Details
        Object = $Object
    }
}

# Verify connection
Write-Check "Verifying Exchange Online Connection"
try {
    $orgConfig = Get-OrganizationConfig -ErrorAction Stop
    Write-Host "  ✓ Connected to: $($orgConfig.Name)" -ForegroundColor Green
} catch {
    Write-Host "  ✗ Not connected to Exchange Online" -ForegroundColor Red
    Write-Host "  Attempting to connect..." -ForegroundColor Yellow
    
    try {
        # Check if ExchangeOnlineManagement module is installed
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Write-Host "  ✗ ExchangeOnlineManagement module not found!" -ForegroundColor Red
            Write-Host "  Please install it first: Install-Module -Name ExchangeOnlineManagement -Force" -ForegroundColor Yellow
            exit 1
        }
        
        # Import the module if not already loaded
        if (-not (Get-Module -Name ExchangeOnlineManagement)) {
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
        }
        
        # Connect to Exchange Online
        Write-Host "  Connecting to Exchange Online (you may be prompted to sign in)..." -ForegroundColor Cyan
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        
        # Verify connection
        $orgConfig = Get-OrganizationConfig -ErrorAction Stop
        Write-Host "  ✓ Successfully connected to: $($orgConfig.Name)" -ForegroundColor Green
        
    } catch {
        Write-Host "  ✗ Failed to connect to Exchange Online!" -ForegroundColor Red
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "`n  Please ensure you have:" -ForegroundColor Yellow
        Write-Host "    1. Exchange Online Management module installed" -ForegroundColor Yellow
        Write-Host "    2. Appropriate admin permissions" -ForegroundColor Yellow
        Write-Host "    3. Network connectivity to Exchange Online" -ForegroundColor Yellow
        Write-Host "`n  Manual connection command: Connect-ExchangeOnline" -ForegroundColor Cyan
        exit 1
    }
}

# 1. ALL Recipients - exhaustive search
Write-Check "Scanning ALL Recipients (this may take a while...)"
try {
    $allRecipients = Get-Recipient -ResultSize Unlimited
    Write-Host "  Total recipients to scan: $($allRecipients.Count)" -ForegroundColor Gray
    
    # Search with exact and partial matches (but flag exact matches)
    $matches = $allRecipients | Where-Object {
        # Exact match (these WILL block creation)
        $_.DisplayName -eq $SearchName -or
        $_.Name -eq $SearchName -or
        $_.Alias -eq $SearchName -or
        $_.PrimarySmtpAddress -eq "$SearchName@*" -or
        # Partial match (word boundary aware - for reference only)
        ($_.DisplayName -match "\b$SearchName\b") -or
        ($_.Name -match "\b$SearchName\b") -or
        ($_.Alias -match "\b$SearchName\b")
    }
    
    if ($matches) {
        foreach ($match in $matches) {
            # Determine if exact match (critical) or partial match (informational)
            $isExactMatch = $false
            $exactFields = @()
            
            if ($match.DisplayName -eq $SearchName) { $isExactMatch = $true; $exactFields += "DisplayName" }
            if ($match.Name -eq $SearchName) { $isExactMatch = $true; $exactFields += "Name" }
            if ($match.Alias -eq $SearchName) { $isExactMatch = $true; $exactFields += "Alias" }
            
            $matchType = if ($isExactMatch) { "EXACT MATCH - BLOCKS CREATION" } else { "Partial/Word match - Reference only" }
            $color = if ($isExactMatch) { "Red" } else { "Yellow" }
            
            Write-Finding -Type "$($match.RecipientTypeDetails) [$matchType]" -Details "DisplayName: '$($match.DisplayName)', Alias: '$($match.Alias)', Email: $($match.PrimarySmtpAddress)" -Object $match
            
            if ($isExactMatch) {
                Write-Host "     🚫 EXACT MATCH on: $($exactFields -join ', ')" -ForegroundColor Red
            } else {
                Write-Host "     ℹ Partial match - may not block creation" -ForegroundColor Gray
            }
            
            # Show all email addresses
            if ($match.EmailAddresses) {
                Write-Host "     Email Addresses: $($match.EmailAddresses -join '; ')" -ForegroundColor DarkYellow
            }
        }
    } else {
        Write-Host "  ✓ No recipients found" -ForegroundColor Green
    }
} catch {
    Write-Host "  ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
}

# 2. Soft-deleted mailboxes - CRITICAL CHECK
Write-Check "Checking Soft-Deleted Mailboxes (CRITICAL)"
try {
    $softDeleted = Get-Mailbox -SoftDeletedMailbox -ResultSize Unlimited | Where-Object {
        # Exact matches only (these are critical)
        $_.DisplayName -eq $SearchName -or
        $_.Name -eq $SearchName -or
        $_.Alias -eq $SearchName
    }
    
    if ($softDeleted) {
        Write-Host "  ⚠ LIKELY CAUSE OF ISSUE! Soft-deleted mailboxes found:" -ForegroundColor Red -BackgroundColor Yellow
        foreach ($mb in $softDeleted) {
            Write-Finding -Type "SOFT-DELETED $($mb.RecipientTypeDetails)" -Details "DisplayName: '$($mb.DisplayName)', Alias: '$($mb.Alias)', Email: $($mb.PrimarySmtpAddress), Deleted: $($mb.WhenSoftDeleted)" -Object $mb
            
            Write-Host "`n  ╔═══════════════════════════════════════════════════════════════════════╗" -ForegroundColor Red
            Write-Host "  ║  SOLUTION: Permanently delete this mailbox:                          ║" -ForegroundColor Red
            Write-Host "  ╠═══════════════════════════════════════════════════════════════════════╣" -ForegroundColor Red
            Write-Host "  ║  Remove-Mailbox -Identity '$($mb.PrimarySmtpAddress)' ``              ║" -ForegroundColor Yellow
            Write-Host "  ║      -PermanentlyDelete -Confirm:`$false                              ║" -ForegroundColor Yellow
            Write-Host "  ╚═══════════════════════════════════════════════════════════════════════╝" -ForegroundColor Red
        }
    } else {
        Write-Host "  ✓ No soft-deleted mailboxes found" -ForegroundColor Green
    }
} catch {
    Write-Host "  ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
}

# 3. Inactive mailboxes
Write-Check "Checking Inactive Mailboxes"
try {
    $inactiveMailboxes = Get-Mailbox -InactiveMailboxOnly -ResultSize Unlimited | Where-Object {
        $_.DisplayName -like "*$SearchName*" -or
        $_.Name -like "*$SearchName*" -or
        $_.Alias -like "*$SearchName*"
    }
    
    if ($inactiveMailboxes) {
        foreach ($mb in $inactiveMailboxes) {
            Write-Finding -Type "INACTIVE $($mb.RecipientTypeDetails)" -Details "DisplayName: '$($mb.DisplayName)', Alias: '$($mb.Alias)', Email: $($mb.PrimarySmtpAddress)" -Object $mb
        }
    } else {
        Write-Host "  ✓ No inactive mailboxes found" -ForegroundColor Green
    }
} catch {
    Write-Host "  ⚠ Inactive mailbox check unavailable (requires appropriate license)" -ForegroundColor Yellow
}

# 4. EXOMailbox with all properties
Write-Check "Deep Mailbox Scan with All Properties"
try {
    $exoMailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties DisplayName,Name,Alias,PrimarySmtpAddress,EmailAddresses,LegacyExchangeDN,ExternalDirectoryObjectId | Where-Object {
        # Prioritize exact matches
        $_.DisplayName -eq $SearchName -or
        $_.Name -eq $SearchName -or
        $_.Alias -eq $SearchName -or
        # Also include word boundary matches for reference
        ($_.DisplayName -match "\b$SearchName\b") -or
        ($_.Name -match "\b$SearchName\b")
    }
    
    if ($exoMailboxes) {
        foreach ($mb in $exoMailboxes) {
            $isExactMatch = ($mb.DisplayName -eq $SearchName) -or ($mb.Name -eq $SearchName) -or ($mb.Alias -eq $SearchName)
            $matchType = if ($isExactMatch) { "EXACT" } else { "PARTIAL" }
            
            Write-Finding -Type "EXOMailbox-$($mb.RecipientTypeDetails) [$matchType]" -Details "DisplayName: '$($mb.DisplayName)', Alias: '$($mb.Alias)', LegacyDN: $($mb.LegacyExchangeDN)" -Object $mb
            
            if ($isExactMatch) {
                Write-Host "     🚫 This is an EXACT match - WILL BLOCK creation" -ForegroundColor Red
            } else {
                Write-Host "     ℹ Partial match - likely won't block creation" -ForegroundColor Gray
            }
        }
    } else {
        Write-Host "  ✓ No mailboxes found with EXO cmdlet" -ForegroundColor Green
    }
} catch {
    Write-Host "  ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
}

# 5. Distribution groups - all types
Write-Check "Checking ALL Distribution Group Types"
try {
    $allGroups = Get-DistributionGroup -ResultSize Unlimited | Where-Object {
        $_.DisplayName -eq $SearchName -or
        $_.Name -eq $SearchName -or
        $_.Alias -eq $SearchName -or
        # Also word boundary matches for reference
        ($_.DisplayName -match "\b$SearchName\b") -or
        ($_.Name -match "\b$SearchName\b")
    }
    
    if ($allGroups) {
        foreach ($grp in $allGroups) {
            $isExactMatch = ($grp.DisplayName -eq $SearchName) -or ($grp.Name -eq $SearchName) -or ($grp.Alias -eq $SearchName)
            $matchType = if ($isExactMatch) { "EXACT" } else { "PARTIAL" }
            
            Write-Finding -Type "$($grp.RecipientTypeDetails) [$matchType]" -Details "DisplayName: '$($grp.DisplayName)', Alias: '$($grp.Alias)', Email: $($grp.PrimarySmtpAddress)" -Object $grp
            
            if ($isExactMatch) {
                Write-Host "     🚫 EXACT match - WILL BLOCK creation" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "  ✓ No distribution groups found" -ForegroundColor Green
    }
} catch {
    Write-Host "  ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
}

# 6. Microsoft 365 Groups
Write-Check "Checking Microsoft 365 Groups (Unified Groups)"
try {
    $unifiedGroups = Get-UnifiedGroup -ResultSize Unlimited | Where-Object {
        $_.DisplayName -like "*$SearchName*" -or
        $_.Alias -like "*$SearchName*" -or
        $_.PrimarySmtpAddress -like "*$SearchName*"
    }
    
    if ($unifiedGroups) {
        foreach ($grp in $unifiedGroups) {
            Write-Finding -Type "UnifiedGroup" -Details "DisplayName: '$($grp.DisplayName)', Alias: '$($grp.Alias)', Email: $($grp.PrimarySmtpAddress)" -Object $grp
        }
    } else {
        Write-Host "  ✓ No M365 groups found" -ForegroundColor Green
    }
} catch {
    Write-Host "  ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
}

# 7. Dynamic Distribution Groups
Write-Check "Checking Dynamic Distribution Groups"
try {
    $dynamicGroups = Get-DynamicDistributionGroup -ResultSize Unlimited | Where-Object {
        $_.DisplayName -like "*$SearchName*" -or
        $_.Name -like "*$SearchName*" -or
        $_.Alias -like "*$SearchName*"
    }
    
    if ($dynamicGroups) {
        foreach ($grp in $dynamicGroups) {
            Write-Finding -Type "DynamicDistributionGroup" -Details "DisplayName: '$($grp.DisplayName)', Alias: '$($grp.Alias)', Email: $($grp.PrimarySmtpAddress)" -Object $grp
        }
    } else {
        Write-Host "  ✓ No dynamic distribution groups found" -ForegroundColor Green
    }
} catch {
    Write-Host "  ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
}

# 8. Mail contacts and mail users
Write-Check "Checking Mail Contacts"
try {
    $contacts = Get-MailContact -ResultSize Unlimited | Where-Object {
        $_.DisplayName -like "*$SearchName*" -or
        $_.Alias -like "*$SearchName*"
    }
    
    if ($contacts) {
        foreach ($contact in $contacts) {
            Write-Finding -Type "MailContact" -Details "DisplayName: '$($contact.DisplayName)', Alias: '$($contact.Alias)', External: $($contact.ExternalEmailAddress)" -Object $contact
        }
    } else {
        Write-Host "  ✓ No mail contacts found" -ForegroundColor Green
    }
} catch {
    Write-Host "  ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Check "Checking Mail Users"
try {
    $mailUsers = Get-MailUser -ResultSize Unlimited | Where-Object {
        $_.DisplayName -like "*$SearchName*" -or
        $_.Alias -like "*$SearchName*"
    }
    
    if ($mailUsers) {
        foreach ($user in $mailUsers) {
            Write-Finding -Type "MailUser" -Details "DisplayName: '$($user.DisplayName)', Alias: '$($user.Alias)', External: $($user.ExternalEmailAddress)" -Object $user
        }
    } else {
        Write-Host "  ✓ No mail users found" -ForegroundColor Green
    }
} catch {
    Write-Host "  ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
}

# 9. Remote mailboxes (hybrid)
Write-Check "Checking Remote Mailboxes (Hybrid Environments)"
try {
    $remoteMailboxes = Get-RemoteMailbox -ResultSize Unlimited | Where-Object {
        $_.DisplayName -like "*$SearchName*" -or
        $_.Alias -like "*$SearchName*"
    }
    
    if ($remoteMailboxes) {
        foreach ($mb in $remoteMailboxes) {
            Write-Finding -Type "RemoteMailbox" -Details "DisplayName: '$($mb.DisplayName)', Alias: '$($mb.Alias)', RemoteRoutingAddress: $($mb.RemoteRoutingAddress)" -Object $mb
        }
    } else {
        Write-Host "  ✓ No remote mailboxes found" -ForegroundColor Green
    }
} catch {
    Write-Host "  ⚠ Remote mailbox cmdlet not available (not a hybrid environment)" -ForegroundColor Gray
}

# 10. Arbitration and system mailboxes
Write-Check "Checking System and Arbitration Mailboxes"
try {
    # Try with RecipientTypeDetails filter instead
    $systemMailboxes = Get-Mailbox -RecipientTypeDetails ArbitrationMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue | Where-Object {
        $_.DisplayName -like "*$SearchName*" -or
        $_.Name -like "*$SearchName*" -or
        $_.Alias -like "*$SearchName*"
    }
    
    if ($systemMailboxes) {
        foreach ($mb in $systemMailboxes) {
            Write-Finding -Type "SystemMailbox-Arbitration" -Details "DisplayName: '$($mb.DisplayName)', Alias: '$($mb.Alias)'" -Object $mb
        }
    } else {
        Write-Host "  ✓ No system mailboxes found" -ForegroundColor Green
    }
} catch {
    Write-Host "  ⚠ System mailbox check unavailable in this Exchange Online version" -ForegroundColor Gray
}

# 11. Public folder mailboxes
Write-Check "Checking Public Folder Mailboxes"
try {
    # Use RecipientTypeDetails for compatibility
    $pfMailboxes = Get-Mailbox -RecipientTypeDetails PublicFolderMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue | Where-Object {
        $_.DisplayName -like "*$SearchName*" -or
        $_.Alias -like "*$SearchName*"
    }
    
    if ($pfMailboxes) {
        foreach ($mb in $pfMailboxes) {
            Write-Finding -Type "PublicFolderMailbox" -Details "DisplayName: '$($mb.DisplayName)', Alias: '$($mb.Alias)'" -Object $mb
        }
    } else {
        Write-Host "  ✓ No public folder mailboxes found" -ForegroundColor Green
    }
} catch {
    Write-Host "  ⚠ Public folder mailbox check unavailable" -ForegroundColor Gray
}

# 12. Check using Get-User (Azure AD sync)
Write-Check "Checking Azure AD Users via Get-User"
try {
    $adUsers = Get-User -ResultSize Unlimited | Where-Object {
        # Only exact matches or word boundary matches
        $_.DisplayName -eq $SearchName -or
        $_.Name -eq $SearchName -or
        $_.SamAccountName -eq $SearchName -or
        # Word boundary matches for reference
        ($_.DisplayName -match "\b$SearchName\b") -or
        ($_.Name -match "\b$SearchName\b")
    }
    
    if ($adUsers) {
        foreach ($user in $adUsers) {
            # Determine which fields match - EXACT matches only
            $matchingFields = @()
            $criticalMatches = @()
            
            # Check for exact matches only
            if ($user.DisplayName -eq $SearchName) { 
                $matchingFields += "DisplayName: '$($user.DisplayName)' [EXACT]"
                $criticalMatches += "DisplayName"
            } elseif ($user.DisplayName -match "\b$SearchName\b") { 
                $matchingFields += "DisplayName: '$($user.DisplayName)' [contains word]" 
            }
            
            if ($user.Name -eq $SearchName) { 
                $matchingFields += "Name: '$($user.Name)' [EXACT - BLOCKS CREATION]"
                $criticalMatches += "Name"
            } elseif ($user.Name -match "\b$SearchName\b") { 
                $matchingFields += "Name: '$($user.Name)' [contains word]" 
            }
            
            if ($user.SamAccountName -eq $SearchName) { 
                $matchingFields += "SamAccountName/mailNickname: '$($user.SamAccountName)' [EXACT - BLOCKS CREATION]"
                $criticalMatches += "SamAccountName/mailNickname"
            } elseif ($user.SamAccountName -match "\b$SearchName\b") { 
                $matchingFields += "SamAccountName/mailNickname: '$($user.SamAccountName)' [contains word]" 
            }
            
            # WindowsEmailAddress - only flag if exact match on local part
            if ($user.WindowsEmailAddress -like "$SearchName@*") {
                $matchingFields += "WindowsEmailAddress: '$($user.WindowsEmailAddress)' [EXACT]"
                $criticalMatches += "WindowsEmailAddress"
            } elseif ($user.WindowsEmailAddress -match "\b$SearchName\b") {
                $matchingFields += "WindowsEmailAddress: '$($user.WindowsEmailAddress)' [contains word]"
            }
            
            # Skip this user if no fields actually matched
            if ($matchingFields.Count -eq 0) {
                continue
            }
            
            $matchInfo = $matchingFields -join " | "
            
            # Determine if this is a critical finding or just reference
            $findingType = if ($criticalMatches.Count -gt 0) {
                "AzureAD-User [EXACT MATCH - BLOCKS CREATION]"
            } else {
                "AzureAD-User [Partial/Word match - Reference only]"
            }
            
            Write-Finding -Type $findingType -Details "UPN: $($user.UserPrincipalName), RecipientType: $($user.RecipientType)" -Object $user
            Write-Host "     Matching Fields: $matchInfo" -ForegroundColor $(if ($criticalMatches.Count -gt 0) { "Red" } else { "Gray" })
            
            if ($criticalMatches.Count -gt 0) {
                Write-Host "`n     ⚠ CRITICAL CONFLICT DETECTED!" -ForegroundColor Red
                Write-Host "     The following properties EXACTLY match '$SearchName' and will BLOCK mailbox creation:" -ForegroundColor Red
                foreach ($match in $criticalMatches) {
                    Write-Host "       🚫 $match" -ForegroundColor Red
                }
            } else {
                Write-Host "     ℹ This is a partial match - should NOT block creation of '$SearchName'" -ForegroundColor Gray
            }
            
            Write-Host "`n     User Details:" -ForegroundColor Gray
            Write-Host "       - DisplayName: '$($user.DisplayName)'" -ForegroundColor Gray
            Write-Host "       - Name: '$($user.Name)'" -ForegroundColor $(if ($user.Name -eq $SearchName) { "Red" } else { "Gray" })
            Write-Host "       - FirstName: '$($user.FirstName)'" -ForegroundColor Gray
            Write-Host "       - LastName: '$($user.LastName)'" -ForegroundColor Gray
            Write-Host "       - SamAccountName/mailNickname: '$($user.SamAccountName)'" -ForegroundColor $(if ($user.SamAccountName -eq $SearchName) { "Red" } else { "Gray" })
            Write-Host "       - WindowsEmailAddress: '$($user.WindowsEmailAddress)'" -ForegroundColor Gray
            Write-Host "       - UserPrincipalName: '$($user.UserPrincipalName)'" -ForegroundColor Gray
            Write-Host "       - RecipientType: $($user.RecipientType)" -ForegroundColor Gray
            Write-Host "       - RecipientTypeDetails: $($user.RecipientTypeDetails)" -ForegroundColor Gray
            
            # Get additional mailbox info if user has a mailbox
            $associatedMailbox = $null
            if ($user.RecipientType -ne "User") {
                try {
                    $associatedMailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
                    if ($associatedMailbox) {
                        Write-Host "`n     Associated Mailbox:" -ForegroundColor Cyan
                        Write-Host "       - Alias: '$($associatedMailbox.Alias)'" -ForegroundColor $(if ($associatedMailbox.Alias -eq $SearchName) { "Red" } else { "Cyan" })
                        Write-Host "       - PrimarySmtpAddress: $($associatedMailbox.PrimarySmtpAddress)" -ForegroundColor Cyan
                        Write-Host "       - MailboxType: $($associatedMailbox.RecipientTypeDetails)" -ForegroundColor Cyan
                        
                        if ($associatedMailbox.Alias -eq $SearchName) {
                            Write-Host "       🚫 Mailbox Alias EXACTLY matches '$SearchName' - BLOCKS CREATION!" -ForegroundColor Red
                            $criticalMatches += "Mailbox.Alias"
                        }
                    }
                } catch {
                    # Silently continue if mailbox lookup fails
                }
            }
            
            # Generate fix commands
            if ($criticalMatches.Count -gt 0) {
                Write-Host "`n     ╔═══════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
                Write-Host "     ║                    AUTOMATED FIX COMMANDS                         ║" -ForegroundColor Green
                Write-Host "     ╚═══════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
                Write-Host "`n     Copy and run these PowerShell commands to fix the conflict:`n" -ForegroundColor Yellow
                
                # Suggest new name based on user details
                $suggestedName = if ($user.FirstName -and $user.LastName) {
                    "$($user.FirstName).$($user.LastName)".ToLower() -replace '\s',''
                } elseif ($user.DisplayName) {
                    $user.DisplayName -replace '\s','.' -replace '[^a-zA-Z0-9.]',''
                } else {
                    "user-$($user.UserPrincipalName.Split('@')[0])"
                }
                
                Write-Host "     # Fix the Name property" -ForegroundColor Cyan
                Write-Host "     Set-User -Identity '$($user.UserPrincipalName)' -Name '$($user.DisplayName)'" -ForegroundColor White
                
                if ($user.SamAccountName -eq $SearchName -or ($associatedMailbox -and $associatedMailbox.Alias -eq $SearchName)) {
                    Write-Host "`n     # Fix the mailNickname/Alias (if user has mailbox)" -ForegroundColor Cyan
                    Write-Host "     Set-Mailbox -Identity '$($user.UserPrincipalName)' -Alias '$suggestedName'" -ForegroundColor White
                }
                
                $userDomain = $user.UserPrincipalName.Split('@')[1]
                Write-Host "`n     # Alternative: Change User Principal Name (updates multiple properties)" -ForegroundColor Cyan
                Write-Host "     Set-User -Identity '$($user.UserPrincipalName)' -UserPrincipalName '$suggestedName@$userDomain'" -ForegroundColor White
                
                Write-Host "`n     # Verify the changes" -ForegroundColor Cyan
                Write-Host "     Get-User -Identity '$($user.UserPrincipalName)' | Select-Object DisplayName, Name, SamAccountName, UserPrincipalName" -ForegroundColor White
                
                Write-Host "`n     ⏱ After running these commands, wait 5-10 minutes for sync, then try creating your shared mailbox.`n" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "  ✓ No Azure AD users found" -ForegroundColor Green
    }
} catch {
    Write-Host "  ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
}

# 13. Check exact match with Get-Recipient filter
Write-Check "Checking Exact Match with Recipient Filter"
try {
    $exactDisplayName = Get-Recipient -Filter "DisplayName -eq '$SearchName'" -ResultSize Unlimited
    $exactAlias = Get-Recipient -Filter "Alias -eq '$SearchName'" -ResultSize Unlimited
    $exactName = Get-Recipient -Filter "Name -eq '$SearchName'" -ResultSize Unlimited
    
    $allExact = @()
    $allExact += $exactDisplayName
    $allExact += $exactAlias
    $allExact += $exactName
    $allExact = $allExact | Select-Object -Unique
    
    if ($allExact) {
        Write-Host "  ⚠ EXACT MATCHES FOUND (most likely cause):" -ForegroundColor Red
        foreach ($item in $allExact) {
            # Determine which property exactly matches
            $exactMatchFields = @()
            if ($item.DisplayName -eq $SearchName) { $exactMatchFields += "DisplayName" }
            if ($item.Alias -eq $SearchName) { $exactMatchFields += "Alias" }
            if ($item.Name -eq $SearchName) { $exactMatchFields += "Name" }
            
            Write-Finding -Type "EXACT-$($item.RecipientTypeDetails)" -Details "DisplayName: '$($item.DisplayName)', Name: '$($item.Name)', Alias: '$($item.Alias)', Email: $($item.PrimarySmtpAddress)" -Object $item
            Write-Host "     ⚠ CONFLICT ON: $($exactMatchFields -join ', ')" -ForegroundColor Red -BackgroundColor Yellow
            Write-Host "     Identity: $($item.Identity)" -ForegroundColor Gray
            Write-Host "     DistinguishedName: $($item.DistinguishedName)" -ForegroundColor Gray
            Write-Host "     Guid: $($item.Guid)" -ForegroundColor Gray
            
            # Show all email addresses
            if ($item.EmailAddresses) {
                Write-Host "     All Email Addresses:" -ForegroundColor Gray
                $item.EmailAddresses | ForEach-Object {
                    Write-Host "       - $_" -ForegroundColor DarkGray
                }
            }
        }
    } else {
        Write-Host "  ✓ No exact matches found" -ForegroundColor Green
    }
} catch {
    Write-Host "  ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
}

# 13B. Cross-reference findings with mailbox name conflict
Write-Check "Cross-Referencing All Findings with '$SearchName'"
if ($findings.Count -gt 0) {
    Write-Host "  Analyzing which findings could block shared mailbox creation:" -ForegroundColor Cyan
    Write-Host "  (Only EXACT matches will block creation)`n" -ForegroundColor Gray
    
    $criticalFindings = 0
    $referenceFindings = 0
    
    foreach ($finding in $findings) {
        $obj = $finding.Object
        $conflictReasons = @()
        $hasExactMatch = $false
        
        # Check DisplayName conflict
        if ($obj.DisplayName -eq $SearchName) {
            $conflictReasons += "DisplayName='$($obj.DisplayName)' (EXACT match with requested name)"
            $hasExactMatch = $true
        } elseif ($obj.DisplayName -match "\b$SearchName\b") {
            $conflictReasons += "DisplayName='$($obj.DisplayName)' (contains word '$SearchName' - reference only)"
        }
        
        # Check Name conflict
        if ($obj.Name -eq $SearchName) {
            $conflictReasons += "Name='$($obj.Name)' (EXACT match - THIS WILL BLOCK CREATION)"
            $hasExactMatch = $true
        } elseif ($obj.Name -match "\b$SearchName\b") {
            $conflictReasons += "Name='$($obj.Name)' (contains word '$SearchName' - reference only)"
        }
        
        # Check Alias conflict
        if ($obj.Alias -eq $SearchName) {
            $conflictReasons += "Alias='$($obj.Alias)' (EXACT match - THIS WILL BLOCK CREATION)"
            $hasExactMatch = $true
        } elseif ($obj.Alias -match "\b$SearchName\b") {
            $conflictReasons += "Alias='$($obj.Alias)' (contains word '$SearchName' - reference only)"
        }
        
        # Check email addresses (only flag if starts with searchname@)
        if ($obj.PrimarySmtpAddress -like "$SearchName@*") {
            $conflictReasons += "PrimarySmtpAddress='$($obj.PrimarySmtpAddress)' (EXACT - WILL BLOCK)"
            $hasExactMatch = $true
        } elseif ($obj.PrimarySmtpAddress -like "*$SearchName*") {
            $conflictReasons += "PrimarySmtpAddress='$($obj.PrimarySmtpAddress)' (contains '$SearchName' - reference only)"
        }
        
        if ($conflictReasons.Count -gt 0) {
            if ($hasExactMatch) {
                $criticalFindings++
                Write-Host "`n  [$($finding.Type)] ⚠ CRITICAL" -ForegroundColor Red
            } else {
                $referenceFindings++
                Write-Host "`n  [$($finding.Type)] ℹ Reference" -ForegroundColor Yellow
            }
            
            Write-Host "    Object: $($finding.Details)" -ForegroundColor White
            Write-Host "    Conflict Reasons:" -ForegroundColor Cyan
            foreach ($reason in $conflictReasons) {
                $isBlocking = $reason -match "EXACT|THIS WILL BLOCK"
                if ($isBlocking) {
                    Write-Host "      🚫 $reason" -ForegroundColor Red
                } else {
                    Write-Host "      ℹ  $reason" -ForegroundColor Gray
                }
            }
        }
    }
    
    Write-Host "`n  ═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Summary: $criticalFindings CRITICAL conflict(s), $referenceFindings reference item(s)" -ForegroundColor Cyan
    if ($criticalFindings -eq 0) {
        Write-Host "  ✓ No EXACT matches found - name should be available!" -ForegroundColor Green
    }
} else {
    Write-Host "  No findings to analyze" -ForegroundColor Green
}

# 14. Test actual creation attempt
Write-Check "Attempting Simulated Creation (WhatIf)"
Write-Host "  Testing: New-Mailbox -Shared -Name '$SearchName' -Alias '$SearchName' -WhatIf" -ForegroundColor Gray
try {
    $whatIfOutput = New-Mailbox -Shared -Name $SearchName -Alias $SearchName -WhatIf 2>&1
    Write-Host "  Result: $whatIfOutput" -ForegroundColor Yellow
} catch {
    Write-Host "  ✗ Error during simulation: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.Message -like "*already*") {
        Write-Host "     This confirms the name conflict!" -ForegroundColor Red
    }
}

# SUMMARY
Write-Host "`n`n╔════════════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Magenta
Write-Host "║                              DIAGNOSTIC SUMMARY                                ║" -ForegroundColor Magenta
Write-Host "╚════════════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Magenta

# Count critical vs reference findings
$criticalFindings = $findings | Where-Object { $_.Type -like "*EXACT*" -or $_.Type -like "*CRITICAL*" -or $_.Type -like "*SOFT-DELETED*" }
$referenceFindings = $findings | Where-Object { $_.Type -like "*Partial*" -or $_.Type -like "*Reference*" }

Write-Host "`nTotal Findings: $($findings.Count)" -ForegroundColor $(if ($findings.Count -eq 0) { "Green" } elseif ($criticalFindings.Count -gt 0) { "Red" } else { "Yellow" })
if ($criticalFindings.Count -gt 0) {
    Write-Host "  - Critical conflicts (WILL BLOCK): $($criticalFindings.Count)" -ForegroundColor Red
}
if ($referenceFindings.Count -gt 0) {
    Write-Host "  - Reference items (won't block): $($referenceFindings.Count)" -ForegroundColor Gray
}

if ($findings.Count -eq 0) {
    Write-Host "`n✓ NO CONFLICTS FOUND!" -ForegroundColor Green
    Write-Host "`nThe name '$SearchName' appears to be available." -ForegroundColor Green
    Write-Host "`nYou should be able to create your shared mailbox with:" -ForegroundColor Cyan
    Write-Host "  New-Mailbox -Shared -Name '$SearchName' -Alias '$SearchName' -PrimarySmtpAddress '$SearchName@$DomainName'" -ForegroundColor White
    Write-Host "`nIf you still get an error, it may be due to:" -ForegroundColor Yellow
    Write-Host "  1. A timing/replication delay in Exchange Online (wait 5-10 minutes)" -ForegroundColor Gray
    Write-Host "  2. A reserved name (e.g., 'admin', 'postmaster', 'administrator')" -ForegroundColor Gray
    Write-Host "  3. A recent deletion that hasn't fully cleared (wait up to 30 days)" -ForegroundColor Gray
    Write-Host "  4. Network/permission issues" -ForegroundColor Gray
} elseif ($criticalFindings.Count -eq 0 -and $referenceFindings.Count -gt 0) {
    Write-Host "`n✓ NO CRITICAL CONFLICTS FOUND!" -ForegroundColor Green
    Write-Host "`nOnly partial/reference matches were found - these should NOT block creation." -ForegroundColor Yellow
    Write-Host "The name '$SearchName' should be available." -ForegroundColor Green
    Write-Host "`nYou should be able to create your shared mailbox with:" -ForegroundColor Cyan
    Write-Host "  New-Mailbox -Shared -Name '$SearchName' -Alias '$SearchName' -PrimarySmtpAddress '$SearchName@$DomainName'" -ForegroundColor White
    Write-Host "`nIf you still get an error, review the reference items above to ensure" -ForegroundColor Yellow
    Write-Host "they truly don't have an exact match on the Name or Alias properties." -ForegroundColor Yellow
} else {
    Write-Host "`n╔════════════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Red
    Write-Host "║                         CONFLICTS IDENTIFIED                                   ║" -ForegroundColor Red
    Write-Host "╚════════════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Red
    
    $findings | Group-Object -Property Type | ForEach-Object {
        Write-Host "`n[$($_.Name)] - $($_.Count) item(s):" -ForegroundColor Yellow
        $_.Group | ForEach-Object {
            Write-Host "  • $($_.Details)" -ForegroundColor Red
        }
    }
    
    Write-Host "`n╔════════════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║                            RESOLUTION STEPS                                    ║" -ForegroundColor Green
    Write-Host "╚════════════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    
    $stepNumber = 1
    
    # Check for Azure AD user conflicts (Name/mailNickname)
    $azureADUsers = $findings | Where-Object { $_.Type -like "*AzureAD*" -or $_.Type -like "*User*" }
    if ($azureADUsers) {
        Write-Host "`n$stepNumber. FIX AZURE AD USER CONFLICTS (Name/mailNickname property):" -ForegroundColor Yellow
        $stepNumber++
        Write-Host "   These are the MOST COMMON cause of 'name already in use' errors!`n" -ForegroundColor Red
        
        foreach ($item in $azureADUsers) {
            $user = $item.Object
            Write-Host "   User: $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor White
            
            # Generate suggested name
            $suggestedName = if ($user.FirstName -and $user.LastName) {
                "$($user.FirstName).$($user.LastName)".ToLower() -replace '\s',''
            } else {
                $user.DisplayName -replace '\s','.' -replace '[^a-zA-Z0-9.]','' | Select-Object -First 1
            }
            
            # Fix Name property
            if ($user.Name -eq $SearchName) {
                Write-Host "   # Fix Name property:" -ForegroundColor Cyan
                Write-Host "   Set-User -Identity '$($user.UserPrincipalName)' -Name '$($user.DisplayName)'" -ForegroundColor White
            }
            
            # Fix mailNickname/Alias
            if ($user.SamAccountName -eq $SearchName) {
                Write-Host "   # Fix mailNickname/Alias:" -ForegroundColor Cyan
                Write-Host "   Set-Mailbox -Identity '$($user.UserPrincipalName)' -Alias '$suggestedName' -ErrorAction SilentlyContinue" -ForegroundColor White
            }
            
            Write-Host ""
        }
    }
    
    # Check for soft-deleted items
    $softDeletedItems = $findings | Where-Object { $_.Type -like "*SOFT-DELETED*" }
    if ($softDeletedItems) {
        Write-Host "`n$stepNumber. REMOVE SOFT-DELETED ITEMS (CRITICAL):" -ForegroundColor Yellow
        $stepNumber++
        foreach ($item in $softDeletedItems) {
            $email = $item.Object.PrimarySmtpAddress
            Write-Host "   Remove-Mailbox -Identity '$email' -PermanentlyDelete -Confirm:`$false" -ForegroundColor Cyan
        }
    }
    
    # Check for other mailboxes
    $otherMailboxes = $findings | Where-Object { $_.Type -like "*Mailbox*" -and $_.Type -notlike "*SOFT-DELETED*" -and $_.Type -notlike "*AzureAD*" }
    if ($otherMailboxes) {
        Write-Host "`n$stepNumber. REMOVE OR RENAME EXISTING MAILBOXES:" -ForegroundColor Yellow
        $stepNumber++
        foreach ($item in $otherMailboxes) {
            $email = $item.Object.PrimarySmtpAddress
            Write-Host "   Remove-Mailbox -Identity '$email' -Confirm:`$false" -ForegroundColor Cyan
            Write-Host "   # Or rename:" -ForegroundColor Gray
            Write-Host "   Set-Mailbox -Identity '$email' -Alias 'new-alias-name'" -ForegroundColor Gray
        }
    }
    
    # Check for groups
    $groups = $findings | Where-Object { $_.Type -like "*Group*" }
    if ($groups) {
        Write-Host "`n$stepNumber. REMOVE OR RENAME GROUPS:" -ForegroundColor Yellow
        $stepNumber++
        foreach ($item in $groups) {
            if ($item.Type -eq "UnifiedGroup") {
                Write-Host "   Remove-UnifiedGroup -Identity '$($item.Object.PrimarySmtpAddress)' -Confirm:`$false" -ForegroundColor Cyan
            } else {
                Write-Host "   Remove-DistributionGroup -Identity '$($item.Object.PrimarySmtpAddress)' -Confirm:`$false" -ForegroundColor Cyan
            }
        }
    }
    
    # General wait instruction
    Write-Host "`n$stepNumber. WAIT FOR SYNCHRONIZATION:" -ForegroundColor Yellow
    $stepNumber++
    Write-Host "   After making ANY changes above, wait 5-10 minutes for Azure AD/Exchange sync." -ForegroundColor White
    Write-Host "   Then re-run this diagnostic to verify the conflict is resolved." -ForegroundColor White
    
    # Create shared mailbox command
    Write-Host "`n$stepNumber. CREATE YOUR SHARED MAILBOX:" -ForegroundColor Yellow
    Write-Host "   New-Mailbox -Shared -Name '$SearchName' -Alias '$SearchName' -PrimarySmtpAddress '$SearchName@$DomainName'" -ForegroundColor Cyan
}

# Export results
$exportPath = Join-Path $env:TEMP "NameConflict_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$findings | Export-Clixml -Path ($exportPath -replace '\.txt$', '.xml')
Write-Host "`n📄 Detailed results exported to: $exportPath.xml" -ForegroundColor Green
Write-Host "`n" ("═" * 80) -ForegroundColor Cyan
