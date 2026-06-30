<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Source = https://github.com/directorcia/Office365/blob/master/o365-exo-addins.ps1

Description - Check which add-ins are present on each mailbox in Exchange Online

Prerequisites = 1
1. Ensure connection to Exchange Online has already been completed

More scripts available by joining http://www.ciaopspatron.com

#>

param(
    [string]$MailboxFilter,
    [string]$ExportPath,
    [switch]$EnabledOnly,
    [switch]$DisabledOnly
)

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$warningmessagecolor = "yellow"
$errormessagecolor = "red"

Clear-Host

Write-Host -ForegroundColor $systemmessagecolor "Script started`n"

## Validate Exchange Online connection
try {
    $null = Get-OrganizationConfig -ErrorAction Stop
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "ERROR: Not connected to Exchange Online. Please run o365-connect-exo.ps1 first.`n"
    exit 1
}

## Retrieve mailboxes
try {
    Write-Host -ForegroundColor $processmessagecolor "Retrieving mailboxes..."
    
    if ($MailboxFilter) {
        $mailboxes = Get-Mailbox -Filter "displayname -like '*$MailboxFilter*'" -ResultSize Unlimited -ErrorAction Stop
    }
    else {
        $mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
    }
    
    if ($mailboxes.Count -eq 0) {
        Write-Host -ForegroundColor $warningmessagecolor "No mailboxes found.`n"
        exit 0
    }
    
    Write-Host -ForegroundColor $processmessagecolor "Found $($mailboxes.Count) mailbox(es)`n"
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "ERROR: Failed to retrieve mailboxes: $_`n"
    exit 1
}

## Process add-ins for each mailbox
$addinsCollection = @()
$processedCount = 0
$failedCount = 0

Write-Host -ForegroundColor $processmessagecolor "Checking add-ins for each mailbox...`n"

foreach ($mailbox in $mailboxes) {
    $processedCount++
    $percentComplete = [math]::Round(($processedCount / $mailboxes.Count) * 100, 0)
    Write-Progress -Activity "Processing mailboxes" -Status "$($mailbox.PrimarySmtpAddress)" -PercentComplete $percentComplete
    
    try {
        $apps = Get-App -Mailbox $mailbox.UserPrincipalName -ErrorAction Stop | 
                Select-Object DisplayName, ProviderName, Enabled, AppVersion
        
        if ($apps) {
            # Apply filters
            if ($EnabledOnly) {
                $apps = $apps | Where-Object { $_.Enabled -eq $true }
            }
            elseif ($DisabledOnly) {
                $apps = $apps | Where-Object { $_.Enabled -eq $false }
            }
            
            if ($apps) {
                foreach ($app in $apps) {
                    $addinsCollection += [PSCustomObject]@{
                        Mailbox      = $mailbox.PrimarySmtpAddress
                        DisplayName  = $app.DisplayName
                        ProviderName = $app.ProviderName
                        Enabled      = $app.Enabled
                        AppVersion   = $app.AppVersion
                    }
                }
            }
        }
    }
    catch {
        Write-Host -ForegroundColor $errormessagecolor "ERROR: Failed to get add-ins for $($mailbox.PrimarySmtpAddress): $_"
        $failedCount++
    }
}

Write-Progress -Activity "Processing mailboxes" -Completed

Write-Host ""

## Display results
if ($addinsCollection.Count -gt 0) {
    Write-Host -ForegroundColor $processmessagecolor "Add-ins Summary:"
    Write-Host "-" * 100
    $addinsCollection | Format-Table -AutoSize -Property Mailbox, DisplayName, ProviderName, Enabled, AppVersion
    
    ## Export to CSV if requested
    if ($ExportPath) {
        try {
            $addinsCollection | Export-Csv -Path $ExportPath -NoTypeInformation -Force
            Write-Host -ForegroundColor $processmessagecolor "Results exported to: $ExportPath`n"
        }
        catch {
            Write-Host -ForegroundColor $errormessagecolor "ERROR: Failed to export results: $_`n"
        }
    }
    
    ## Display statistics
    Write-Host -ForegroundColor $processmessagecolor "Statistics:"
    Write-Host "  Total add-ins found: $($addinsCollection.Count)"
    Write-Host "  Enabled add-ins: $($addinsCollection | Where-Object { $_.Enabled -eq $true } | Measure-Object | Select-Object -ExpandProperty Count)"
    Write-Host "  Disabled add-ins: $($addinsCollection | Where-Object { $_.Enabled -eq $false } | Measure-Object | Select-Object -ExpandProperty Count)`n"
}
else {
    Write-Host -ForegroundColor $warningmessagecolor "No add-ins found for the specified criteria.`n"
}

## Display processing summary
Write-Host -ForegroundColor $processmessagecolor "Processing Summary:"
Write-Host "  Mailboxes processed: $processedCount"
Write-Host "  Mailboxes failed: $failedCount`n"

Write-Host -ForegroundColor $systemmessagecolor "Script Completed`n"