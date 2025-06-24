<#
.SYNOPSIS
    Analyzes an email header file to provide a spam and security report, focusing on Exchange Online headers.
.DESCRIPTION
    This script reads an email header from a specified file, parses various key headers related to
    authentication (SPF, DKIM, DMARC), spam filtering (SCL, Forefront, Microsoft Antispam),
    ATP Safe Attachments/Links, and other general message properties.
    It then presents a formatted report with color-coded results to help identify potential issues.
.PARAMETER HeaderFilePath
    The mandatory path to the text file containing the raw email header.
.EXAMPLE
    .\EmailHeaderAnalyzer.ps1 -HeaderFilePath "C:\temp\email_header.txt"
    This command will analyze the header content from the 'email_header.txt' file.
.NOTES
    Author: CIAOPS
    Version: 1.1
    Last Modified: 2025-05-27
    Requires PowerShell 5.0 or higher for some features (e.g., Get-Content -Raw).
    Source https://github.com/directorcia/Office365/blob/master/email-header-report.ps1
    Documentation https://github.com/directorcia/Office365/wiki/Email-Header-Report-Tool
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$HeaderFilePath,
    
    [Parameter(Mandatory = $false)]
    [switch]$Verbose = $false,
    
    [Parameter(Mandatory = $false)]
    [switch]$ShowRawHeaders = $false
)

#region Helper Functions for Output Formatting

# Color functions for better readability
function Write-ColoredText {
    param([string]$Text, [string]$Color = "White")
    Write-Host $Text -ForegroundColor $Color
}

function Write-Header {
    param([string]$Title)
    Write-Host "`n$(New-Object System.String -ArgumentList @('=', 80))" -ForegroundColor Cyan
    Write-Host $Title.ToUpper().PadLeft(($Title.Length + 80) / 2) -ForegroundColor Yellow
    Write-Host "$(New-Object System.String -ArgumentList @('=', 80))" -ForegroundColor Cyan
}

function Write-SubHeader {
    param([string]$Title)
    Write-Host "`n$Title" -ForegroundColor Green
    Write-Host (New-Object System.String -ArgumentList @('-', $Title.Length)) -ForegroundColor Green
}

function Write-TestResult {
    param(
        [string]$Test,
        [string]$Result,
        [string]$Description = "",
        [string]$Reference = "",
        [string]$ForcedColor = "", # Allows overriding default color logic
        [string]$VerboseDetails = "", # Additional details for verbose mode
        [string]$RawHeaderValue = "" # Raw header value for display in verbose mode
    )
    
    $effectiveColor = "White" # Default color

    if ($ForcedColor) {
        $effectiveColor = $ForcedColor
    }
    else {
        # Determine color based on result keywords (case-insensitive)
        $upperResult = $Result.ToUpper()
        switch -regex ($upperResult) {
            "PASS|ALLOW|NONE|NEUTRAL" { $effectiveColor = "Green"; break }
            "FAIL|BLOCK|SPAM|MALWARE" { $effectiveColor = "Red"; break }
            "WARN|SOFT|SOFTFAIL" { $effectiveColor = "Yellow"; break } # Added SOFTFAIL
            # Add more keywords as needed
            default { $effectiveColor = "White" } # Default for non-keyword results
        }
    }
    
    Write-Host "  [$Test]" -NoNewline -ForegroundColor Cyan
    Write-Host " $Result" -ForegroundColor $effectiveColor
    if ($Description) { Write-Host "    Description: $Description" -ForegroundColor Gray }
    if ($Reference) { Write-Host "    Reference: $Reference" -ForegroundColor DarkGray }
    
    # Display additional verbose information if verbose mode is enabled
    if ($Global:VerboseOutput) {
        if ($VerboseDetails) { 
            Write-Host "    Verbose Details: $VerboseDetails" -ForegroundColor Magenta 
        }
        if ($RawHeaderValue -and $Global:ShowRawHeaders) {
            Write-Host "    Raw Header Value: " -NoNewline -ForegroundColor DarkGray
            Write-Host $RawHeaderValue -ForegroundColor Gray
        }
    }
}

#endregion Helper Functions

#region Core Logic Functions

# Check if file exists
if (-not (Test-Path $HeaderFilePath)) {
    Write-Error "File not found: $HeaderFilePath"
    exit 1
}

# Read the header file
try {
    $headerContent = Get-Content $HeaderFilePath -Raw
    Write-ColoredText "Successfully loaded email header from: $HeaderFilePath" "Green"
    Write-ColoredText "Header size: $([math]::Round($headerContent.Length / 1024, 2)) KB" "Gray"
    
    # Set up global verbose flags based on parameters
    $Global:VerboseOutput = $Verbose
    $Global:ShowRawHeaders = $ShowRawHeaders
    
    if ($Global:VerboseOutput) {
        Write-ColoredText "VERBOSE OUTPUT MODE ENABLED" "Magenta"
        if ($Global:ShowRawHeaders) {
            Write-ColoredText "RAW HEADERS DISPLAY ENABLED" "Magenta"
        }
    }
}
catch {
    Write-Error "Failed to read file: $_"
    exit 1
}

Write-Header "EXCHANGE ONLINE EMAIL HEADER SPAM ANALYSIS REPORT"

# Extract header fields using regex
function Get-HeaderValue {
    param([string]$HeaderName, [string]$Content)
    # Handle multi-line headers that are folded with whitespace.
    # Looks for HeaderName:, optional whitespace, captures the value (.+?),
    # until a newline followed by a non-whitespace character (next header),
    # or a double newline (end of headers/body start), or end of content.
    $pattern = "(?im)^$([regex]::Escape($HeaderName))\s*:\s*(.+?)(?=\r?\n\S|\r?\n\r?\n|\Z)"
    if ($Content -match $pattern) {
        # Clean up the value: unfold multi-lines and normalize multiple spaces to a single space.
        $value = $matches[1] -replace '\r?\n\s+', ' ' -replace '\s{2,}', ' '
        return $value.Trim()
    }
    return $null
}

# Get specific values from complex headers like Authentication-Results
function Get-AuthResultValue {
    param([string]$AuthResults, [string]$Protocol)
    # Example: $AuthResults = "spf=pass (sender IP is 1.2.3.4); dkim=fail reason='signature_did_not_verify'"
    # $Protocol = "spf" -> returns "pass"
    # $Protocol = "dkim" -> returns "fail"
    # Matches protocol=value, where value is one or more word characters.
    # It also handles cases where the value might be followed by a space and details in parentheses or other separators.
    if ($AuthResults -match "(?i)$([regex]::Escape($Protocol))=(\w+)") {
        return $matches[1]
    }
    return $null
}

# Helper function to extract email from "Display Name <email@example.com>" or "email@example.com"
function Get-EmailFromHeaderValue {
    param([string]$Value)
    if ($Value -match '<([^>]+)>') {
        # Extracts email from <email@address.com>
        return $matches[1].Trim()
    }
    elseif ($Value -match '\b([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})\b') {
        # Extracts email@address.com
        return $matches[0].Trim()
    }
    return $Value.Trim() # Fallback, return trimmed input if no specific format matched
}

#endregion Core Logic Functions

#region Analysis Functions

function Analyze-Authentication {
    Write-SubHeader "AUTHENTICATION ANALYSIS"
    
    $authResultsHeader = Get-HeaderValue "Authentication-Results" $headerContent
    if ($authResultsHeader) {
        Write-Host "  Authentication Results Header Found:" -ForegroundColor Cyan
        Write-Host "    $authResultsHeader" -ForegroundColor Gray
        Write-Host ""
        
        # SPF Analysis
        $spfResult = Get-AuthResultValue $authResultsHeader "spf"
        if ($spfResult) {
            $description = switch ($spfResult.ToLower()) {
                "pass" { "Sender IP authorized by SPF record." }
                "fail" { "Sender IP NOT authorized by SPF record. High risk." }
                "softfail" { "SPF record suggests IP not authorized, but not a definitive failure." }
                "neutral" { "SPF record makes no assertion about the IP." }
                "none" { "No SPF record found for the sender's domain." }
                "temperror" { "Temporary DNS error during SPF lookup." }
                "permerror" { "Permanent error in SPF record (e.g., syntax error)." }
                default { "Unknown SPF result: $spfResult" }
            }
            Write-TestResult "SPF" $spfResult.ToUpper() $description "RFC 7208 - https://tools.ietf.org/html/rfc7208"
        }

        # DKIM Analysis
        $dkimResult = Get-AuthResultValue $authResultsHeader "dkim"
        if ($dkimResult) {
            $description = switch ($dkimResult.ToLower()) {
                "pass" { "DKIM signature verified successfully." }
                "fail" { "DKIM signature verification FAILED. High risk if DMARC not aligned." }
                "policy" { "Message failed DKIM policy check (e.g., ADSP)." }
                "neutral" { "DKIM signature present but not validated (e.g., key unavailable)." }
                "temperror" { "Temporary error during DKIM validation." }
                "permerror" { "Permanent error in DKIM signature or policy (e.g., key syntax)." }
                default { "Unknown DKIM result: $dkimResult" }
            }
            Write-TestResult "DKIM" $dkimResult.ToUpper() $description "RFC 6376 - https://tools.ietf.org/html/rfc6376"
        }

        # DMARC Analysis
        $dmarcResult = Get-AuthResultValue $authResultsHeader "dmarc"
        if ($dmarcResult) {
            $description = switch ($dmarcResult.ToLower()) {
                "pass" { "Message passes DMARC alignment and policy." }
                "fail" { "Message FAILS DMARC policy (SPF or DKIM failed and not aligned). Action (quarantine/reject) may have been taken." }
                "bestguesspass" { "No DMARC policy found, but message appears legitimate based on heuristics." } # Less common
                "none" { "No DMARC record found for the sender's domain." }
                "temperror" { "Temporary error during DMARC evaluation." }
                "permerror" { "Permanent error in DMARC policy (e.g., syntax error)." }
                default { "Unknown DMARC result: $dmarcResult" }
            }
            Write-TestResult "DMARC" $dmarcResult.ToUpper() $description "RFC 7489 - https://tools.ietf.org/html/rfc7489"
        }

        # CompAuth (Composite Authentication) - Microsoft specific
        if ($authResultsHeader -match "compauth=(\w+)") {
            $compauth = $matches[1]
            $reason = ""
            if ($authResultsHeader -match "reason=([^;]+)") { $reason = "Reason: $($matches[1])" }
            
            $description = switch ($compauth.ToLower()) {
                "pass" { "Composite authentication passed. $reason" }
                "fail" { "Composite authentication FAILED. Indicates potential spoofing. $reason" }
                "softpass" { "Composite authentication soft pass. $reason" } # Usually due to temporary issues or specific configurations
                "none" { "No composite authentication result. $reason" }
                default { "Unknown composite authentication result: $compauth. $reason" }
            }
            Write-TestResult "Composite Auth (CompAuth)" $compauth.ToUpper() $description.Trim()
        }
    }
    else {
        Write-TestResult "Authentication-Results" "NOT FOUND" "No Authentication-Results header found. This is unusual for Exchange Online."
    }

    # Check for DKIM-Signature header details
    $dkimSig = Get-HeaderValue "DKIM-Signature" $headerContent
    if ($dkimSig) {
        Write-Host "`n  DKIM Signature Details:" -ForegroundColor Cyan
        if ($dkimSig -match "d=([^;]+)") {
            Write-TestResult "DKIM Domain (d=)" $matches[1] "Domain that signed the message."
        }
        if ($dkimSig -match "s=([^;]+)") {
            Write-TestResult "DKIM Selector (s=)" $matches[1] "DKIM key selector used for lookup."
        }
        if ($dkimSig -match "a=([^;]+)") {
            Write-TestResult "DKIM Algorithm (a=)" $matches[1] "Signature and hash algorithm used."
        }
        # Add more DKIM tag parsing if needed (e.g., i= for identity)
    }

    # Check Received-SPF header (often present from edge MTA)
    $receivedSPF = Get-HeaderValue "Received-SPF" $headerContent
    if ($receivedSPF) {
        Write-Host "`n  Received-SPF Details:" -ForegroundColor Cyan
        Write-Host "    $receivedSPF" -ForegroundColor Gray
        # Extract key info from Received-SPF, e.g., client-ip
        if ($receivedSPF -match "client-ip=([0-9a-fA-F.:]+)") {
            # Supports IPv4 and IPv6
            Write-TestResult "Client IP (from Received-SPF)" $matches[1] "Sending server IP address as seen by SPF checking MTA."
        }
        # The result (pass, fail, etc.) is usually the first word
        if ($receivedSPF -match "^(\w+)") {
            Write-TestResult "Received-SPF Result" $matches[1].ToUpper() "Result from an earlier SPF check."
        }
    }
}

function Analyze-SpamFiltering {
    Write-SubHeader "EXCHANGE ONLINE SPAM FILTERING ANALYSIS"
    
    # X-MS-Exchange-Organization-SCL (Spam Confidence Level)
    $sclHeaderValue = Get-HeaderValue "X-MS-Exchange-Organization-SCL" $headerContent
    if ($sclHeaderValue) {
        $sclIntValue = -2 # Default for parsing failure
        try {
            $sclIntValue = [int]$sclHeaderValue
        }
        catch {
            Write-Warning "Could not parse SCL value '$sclHeaderValue' as integer."
        }

        $sclDescription = switch ($sclIntValue) {
            -1 { "Trusted sender (bypasses most spam filtering)." }
            0 { "Not spam (message determined to be clean by EOP content filter)." }
            1 { "Not spam (message determined to be clean by EOP content filter)." } # SCL 0 and 1 are generally treated similarly
            { $_ -in 2..4 } { "Low spam probability (content filter determined some suspicion)." }
            5 { "Uncertain (likely spam, typically delivered to Junk Email folder)." }
            6 { "High spam probability (content filter determined high suspicion, typically delivered to Junk Email folder)." }
            { $_ -in 7..8 } { "Very high spam probability (content filter determined very high suspicion, typically delivered to Junk Email folder or quarantined)." }
            9 { "Definite spam (highest confidence, typically quarantined or rejected)." }
            default { "Unknown SCL value or parse error for '$sclHeaderValue'." }
        }
        
        $sclDisplay = if ($sclIntValue -eq -2) { $sclHeaderValue } else { $sclIntValue }
        $sclColor = "White" # Default
        if ($sclIntValue -eq -1) { $sclColor = "Green" }
        elseif ($sclIntValue -ge 0 -and $sclIntValue -le 1) { $sclColor = "Green" }
        elseif ($sclIntValue -ge 2 -and $sclIntValue -le 4) { $sclColor = "Yellow" }
        elseif ($sclIntValue -ge 5 -and $sclIntValue -le 9) { $sclColor = "Red" }
        
        Write-TestResult "SCL (Spam Confidence Level)" "$sclDisplay - $sclDescription" "Scale: -1 to 9 (higher = more spam-like)" "https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/spam-confidence-levels" $sclColor
    }
    else {
        Write-TestResult "SCL (Spam Confidence Level)" "NOT FOUND" "X-MS-Exchange-Organization-SCL header missing."
    }

    # X-Forefront-Antispam-Report Analysis
    $forefrontReport = Get-HeaderValue "X-Forefront-Antispam-Report" $headerContent
    if ($forefrontReport) {
        Write-Host "`n  Forefront Anti-Spam Report Analysis (X-Forefront-Antispam-Report):" -ForegroundColor Cyan
        
        $reportParts = $forefrontReport -split ';'
        foreach ($part in $reportParts) {
            if ($part -match '([^:]+):(.*)') {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                
                $partDescription = ""
                $partColor = "" # For specific coloring if needed

                switch ($key.ToUpper()) {
                    "CIP" { $partDescription = "Connecting IP address of the last external hop."; Write-TestResult "Forefront: Client IP (CIP)" $value $partDescription; break }
                    "CTRY" { $partDescription = "Source country/region code of the connecting IP."; Write-TestResult "Forefront: Country (CTRY)" $value $partDescription; break }
                    "LANG" { $partDescription = "Detected message language code (e.g., en)."; Write-TestResult "Forefront: Language (LANG)" $value $partDescription; break }
                    "SCL" { 
                        $partDescription = "Spam Confidence Level as determined by Forefront (can differ from final SCL)."
                        $ffSclIntValue = -2; try { $ffSclIntValue = [int]$value } catch {}
                        if ($ffSclIntValue -eq -1) { $partColor = "Green" }
                        elseif ($ffSclIntValue -ge 0 -and $ffSclIntValue -le 1) { $partColor = "Green" }
                        elseif ($ffSclIntValue -ge 2 -and $ffSclIntValue -le 4) { $partColor = "Yellow" }
                        elseif ($ffSclIntValue -ge 5 -and $ffSclIntValue -le 9) { $partColor = "Red" }
                        Write-TestResult "Forefront: SCL" $value $partDescription "" $partColor
                        break
                    }
                    "SRV" { 
                        $partDescription = switch ($value.ToUpper()) {
                            "BULK" { "Sender identified as a bulk mailer." }
                            "SPAM" { "Sender identified as a known spam source." }
                            default { "Service type: $value" }
                        }
                        Write-TestResult "Forefront: Service Type (SRV)" $value $partDescription
                        break
                    }
                    "IPV" { 
                        $partDescription = switch ($value.ToUpper()) {
                            "NLI" { "IP not listed in any reputation database (Neutral)." }
                            "CAL" { "IP on Microsoft's internal allow list." }
                            "CBL" { "IP on Microsoft's internal block list." }
                            default { "IP verdict: $value" }
                        }
                        Write-TestResult "Forefront: IP Verdict (IPV)" $value $partDescription
                        break
                    }
                    "SFV" { 
                        $partDescription = switch ($value.ToUpper()) {
                            "SPM" { "Marked as spam by content filter rules." }
                            "BLK" { "Blocked due to content filter rules (e.g., transport rule)." }
                            "NSPM" { "Not spam according to content filter." }
                            "SFE" { "Spam Filter Engine bypassed (e.g., trusted sender)." }
                            default { "Spam filtering verdict: $value" }
                        }
                        Write-TestResult "Forefront: Spam Filter Verdict (SFV)" $value $partDescription
                        break
                    }
                    "H" { $partDescription = "SMTP HELO/EHLO string from connecting server."; Write-TestResult "Forefront: HELO/EHLO (H)" $value $partDescription; break }
                    "PTR" { $partDescription = "Reverse DNS lookup result (PTR record) for connecting IP."; Write-TestResult "Forefront: PTR Record (PTR)" $value $partDescription; break }
                    "CAT" { 
                        $partDescription = switch ($value.ToUpper()) {
                            "BULK" { "Categorized as bulk mail." }
                            "SPAM" { "Categorized as spam." }
                            "PHSH" { "Categorized as phishing." }
                            "MALW" { "Categorized as malware." }
                            "NONE" { "No specific category assigned." }
                            default { "Category: $value" }
                        }
                        Write-TestResult "Forefront: Category (CAT)" $value $partDescription
                        break
                    }
                    "SFS" { 
                        $partDescription = "Spam Filter Rules triggered (IDs)."
                        Write-TestResult "Forefront: Spam Filter Rules (SFS)" $value $partDescription
                        if ($value -match '\((\d+)\)') {
                            # If specific rule IDs are listed
                            Write-Host "      Rule Details:" -ForegroundColor Gray
                            $rules = [regex]::Matches($value, '\((\d+)\)')
                            foreach ($rule in $rules) {
                                Write-Host "        Rule ID: $($rule.Groups[1].Value)" -ForegroundColor DarkGray
                            }
                        }
                        break
                    }
                    "DIR" { 
                        $partDescription = switch ($value.ToUpper()) {
                            "INB" { "Inbound message to the organization." }
                            "OUT" { "Outbound message from the organization." } # Less common in typical analysis scenarios
                            default { "Direction: $value" }
                        }
                        Write-TestResult "Forefront: Direction (DIR)" $value $partDescription
                        break
                    }
                    default { Write-TestResult "Forefront: $key" $value "Unknown Forefront report field." }
                }
            }
        }
    }

    # X-Microsoft-Antispam header
    $msAntispam = Get-HeaderValue "X-Microsoft-Antispam" $headerContent
    if ($msAntispam) {
        Write-Host "`n  Microsoft Anti-Spam Analysis (X-Microsoft-Antispam):" -ForegroundColor Cyan
        Write-Host "    $msAntispam" -ForegroundColor Gray # Display the raw header for reference
        
        # Parse BCL (Bulk Complaint Level)
        if ($msAntispam -match "BCL:(\d+)") {
            $bcl = $matches[1]
            $bclIntValue = [int]$bcl
            $bclDescription = "Bulk Complaint Level (0-9). Higher values indicate the sender generates more bulk mail complaints."
            $bclColor = "White"
            if ($bclIntValue -le 3) { $bclColor = "Green" } # Low BCL
            elseif ($bclIntValue -le 6) { $bclColor = "Yellow" } # Medium BCL
            else { $bclColor = "Red" } # High BCL
            Write-TestResult "BCL (Bulk Complaint Level)" $bcl $bclDescription "https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/bulk-complaint-level-values" $bclColor
        }

        # Parse PCL (Phishing Confidence Level) - if present
        if ($msAntispam -match "PCL:(\d+)") {
            $pcl = $matches[1]
            # PCL interpretation can be complex, often internal.
            Write-TestResult "PCL (Phishing Confidence Level)" $pcl "Phishing Confidence Level. Interpretation varies."
        }
        # Other X-Microsoft-Antispam values like 'OrigIP' might be present
    }

    # X-MS-Exchange-AtpMessageProperties (Defender for Office 365 / ATP)
    $atpProps = Get-HeaderValue "X-MS-Exchange-AtpMessageProperties" $headerContent
    if ($atpProps) {
        Write-Host "`n  ATP Message Properties (X-MS-Exchange-AtpMessageProperties):" -ForegroundColor Cyan
        $props = $atpProps -split '\|' # Properties are often pipe-separated
        foreach ($prop in $props) {
            $trimmedProp = $prop.Trim()
            switch ($trimmedProp) {
                "SA" { Write-TestResult "ATP: Safe Attachments" "PROCESSED" "Message processed by Safe Attachments." }
                "SL" { Write-TestResult "ATP: Safe Links" "PROCESSED" "Message processed by Safe Links." }
                "HVE" { Write-TestResult "ATP: High Volume Email" "DETECTED" "Message flagged as high volume email by ATP." }
                # Add more ATP properties if known
                default { if ($trimmedProp) { Write-TestResult "ATP Property" $trimmedProp } }
            }
        }
    }

    # X-Microsoft-Antispam-Mailbox-Delivery (Information about final delivery)
    $mailboxDelivery = Get-HeaderValue "X-Microsoft-Antispam-Mailbox-Delivery" $headerContent
    if ($mailboxDelivery) {
        Write-Host "`n  Mailbox Delivery Analysis (X-Microsoft-Antispam-Mailbox-Delivery):" -ForegroundColor Cyan
        Write-Host "    $mailboxDelivery" -ForegroundColor Gray
        
        if ($mailboxDelivery -match "dest:([^;]+)") {
            $dest = $matches[1].Trim()
            $destDescription = switch ($dest.ToUpper()) {
                "I" { "Delivered to Inbox." }
                "J" { "Delivered to Junk Email folder." }
                "D" { "Deleted (e.g., by transport rule or malware filter)." }
                "Q" { "Quarantined." }
                "S" { "Skipped (e.g. public folder, or other non-mailbox recipient)" }
                default { "Destination: $dest" }
            }
            $destColor = if ($dest -eq "J" -or $dest -eq "D" -or $dest -eq "Q") { "Red" } else { "Green" }
            Write-TestResult "Delivery Destination (dest)" $dest $destDescription "" $destColor
        }

        if ($mailboxDelivery -match "RF:([^;]+)") {
            # Routing Flag
            $rf = $matches[1].Trim()
            Write-TestResult "Routing Flag (RF)" $rf "Internal message routing information."
        }
        # Other fields like 'uchk', 'auth' might be present
    }
}

function Analyze-SafeAttachments {
    Write-SubHeader "SAFE ATTACHMENTS ANALYSIS (DEFENDER FOR OFFICE 365)"
    
    $safeAttachmentsResult = Get-HeaderValue "X-MS-Exchange-Organization-SafeAttachment-Result" $headerContent
    if ($safeAttachmentsResult) {
        $saDescription = switch ($safeAttachmentsResult.ToLower()) {
            "clean" { "No malicious content detected in attachments by Safe Attachments." }
            "block" { "Malicious content detected; attachment(s) blocked or removed." } # Common for malware
            "replace" { "Malicious attachment replaced with a placeholder." }
            "dynamicdelivery" { "Dynamic Delivery in progress or completed; user may receive placeholder first." }
            "timeout" { "Safe Attachments analysis timed out." }
            "error" { "Error occurred during Safe Attachments analysis." }
            default { "Unknown Safe Attachments result: $safeAttachmentsResult" }
        }
        Write-TestResult "Safe Attachments Result" $safeAttachmentsResult.ToUpper() $saDescription "https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/safe-attachments-overview"
    }
    else {
        Write-TestResult "Safe Attachments" "NOT PROCESSED / NO HEADER" "Safe Attachments may not be enabled, no attachments, or header not present."
    }
}

function Analyze-SafeLinks {
    Write-SubHeader "SAFE LINKS ANALYSIS (DEFENDER FOR OFFICE 365)"
    
    $safeLinksResult = Get-HeaderValue "X-MS-Exchange-Organization-SafeLinks-Result" $headerContent
    if ($safeLinksResult) {
        $slDescription = switch ($safeLinksResult.ToLower()) {
            "clean" { "No malicious URLs detected by Safe Links at time of delivery." }
            "block" { "Malicious URLs detected and rewriting/blocking applied." } # Common for known bad URLs
            "pending" { "Safe Links processing is pending for some URLs." }
            "error" { "Error occurred during Safe Links analysis." }
            "notscanned" { "URLs were not scanned (e.g., due to policy)." }
            default { "Unknown Safe Links result: $safeLinksResult" }
        }
        Write-TestResult "Safe Links Result" $safeLinksResult.ToUpper() $slDescription "https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/safe-links-overview"
    }
    else {
        Write-TestResult "Safe Links" "NOT PROCESSED / NO HEADER" "Safe Links may not be enabled, no scannable URLs, or header not present."
    }
}

function Analyze-TransportRules {
    Write-SubHeader "TRANSPORT RULES (MAIL FLOW RULES) ANALYSIS"
    
    # X-MS-Exchange-Organization-Processed-By-MBT (Mailbox Transport) indicates processing by transport rules
    $processedByMBT = Get-HeaderValue "X-MS-Exchange-Organization-Processed-By-MBT" $headerContent
    if ($processedByMBT) {
        Write-TestResult "Transport Rule Processing" "INDICATED" "Message likely processed by Exchange Transport Rules (Mail Flow Rules)." "https://learn.microsoft.com/en-us/exchange/security-and-compliance/mail-flow-rules/mail-flow-rules"
        Write-Host "    Details (X-MS-Exchange-Organization-Processed-By-MBT): $processedByMBT" -ForegroundColor Gray
        
        # Check for specific transport rule hit header (less common to be exposed directly, but sometimes custom headers are added)
        # Example: X-Transport-Rule-Name: "My Custom Rule"
        # This would require knowing the custom header name.
    }
    else {
        Write-TestResult "Transport Rule Processing" "NO INDICATION" "No clear indication of transport rule processing from standard headers."
    }
}

function Analyze-GeneralHeaders {
    Write-SubHeader "GENERAL MESSAGE ANALYSIS"
    
    # X-Originating-IP (Often added by the first Microsoft Exchange server that receives the message from an external source)
    $originatingIP = Get-HeaderValue "X-Originating-IP" $headerContent
    if ($originatingIP) {
        Write-TestResult "Originating IP (X-Originating-IP)" $originatingIP "Typically the source IP of the message as seen by the first Microsoft server."
    }

    $messageId = Get-HeaderValue "Message-ID" $headerContent
    if ($messageId) {
        Write-TestResult "Message ID" $messageId "Unique identifier for this message, usually set by the sending system."
    }

    $networkMessageId = Get-HeaderValue "X-MS-Exchange-Organization-Network-Message-Id" $headerContent
    if ($networkMessageId) {
        Write-TestResult "Network Message ID" $networkMessageId "Exchange internal message identifier, useful for tracking within Exchange."
    }

    $subject = Get-HeaderValue "Subject" $headerContent
    if ($subject) {
        Write-TestResult "Subject" $subject
        if ($subject -match '(?i)\[SPAM\]|\[BULK\]|\*\*\*SPAM\*\*\*|\[SUSPECTED SPAM\]') {
            # Case-insensitive match for common spam tags
            Write-TestResult "Subject Spam Tag" "SPAM" "Subject line contains common spam/bulk tags." # Result "SPAM" will be colored Red
        }
    }

    # Return-Path vs From analysis
    $returnPathHeaderVal = Get-HeaderValue "Return-Path" $headerContent
    $fromHeaderVal = Get-HeaderValue "From" $headerContent

    if ($returnPathHeaderVal -and $fromHeaderVal) {
        $returnPathEmail = Get-EmailFromHeaderValue $returnPathHeaderVal
        $fromEmail = Get-EmailFromHeaderValue $fromHeaderVal

        if ($returnPathEmail -ne "" -and $fromEmail -ne "") {
            $lcReturnPathEmail = $returnPathEmail.ToLower()
            $lcFromEmail = $fromEmail.ToLower()

            if ($lcReturnPathEmail -ne $lcFromEmail) {
                $returnPathDomain = ""
                $fromDomain = ""
                if ($lcReturnPathEmail -match "@(.+)") { $returnPathDomain = $matches[1] }
                if ($lcFromEmail -match "@(.+)") { $fromDomain = $matches[1] }

                if ($returnPathDomain -ne "" -and $fromDomain -ne "" -and $returnPathDomain -ne $fromDomain) {
                    Write-TestResult "Return-Path vs From" "MISMATCH (DOMAINS)" "Return-Path domain ($returnPathDomain) differs from From domain ($fromDomain). RP: $returnPathEmail, From: $fromEmail" "" "Yellow"
                }
                else {
                    Write-TestResult "Return-Path vs From" "MISMATCH (ADDRESSES, SAME DOMAIN)" "Return-Path ($returnPathEmail) differs from From address ($fromEmail) but domains appear to match." "" "Gray"
                }
            }
            else {
                Write-TestResult "Return-Path vs From" "MATCH" "Return-Path and From addresses appear to match: $returnPathEmail"
            }
        }
        else {
            Write-TestResult "Return-Path vs From" "INFO" "Could not fully extract email from Return-Path ('$returnPathEmail') or From ('$fromEmail') for comparison."
        }
    }
}

#endregion Analysis Functions

#region Summary and Execution

function Show-Summary {
    Write-SubHeader "ANALYSIS SUMMARY & RECOMMENDATIONS"
    
    # Extract key values for summary
    $scl = Get-HeaderValue "X-MS-Exchange-Organization-SCL" $headerContent
    $authResults = Get-HeaderValue "Authentication-Results" $headerContent
    $forefrontReport = Get-HeaderValue "X-Forefront-Antispam-Report" $headerContent
    $deliveryInfo = Get-HeaderValue "X-Microsoft-Antispam-Mailbox-Delivery" $headerContent
    
    $riskFactors = [System.Collections.Generic.List[string]]::new()
    $positiveFactors = [System.Collections.Generic.List[string]]::new()
    $neutralFactors = [System.Collections.Generic.List[string]]::new()

    # SCL Analysis
    $sclIntValue = -2
    if ($scl) { 
        try { $sclIntValue = [int]$scl } catch {}
        if ($sclIntValue -ge 9) { $riskFactors.Add("SCL=$scl (Definite Spam)") }
        elseif ($sclIntValue -ge 5) { $riskFactors.Add("SCL=$scl (Likely Spam)") }
        elseif ($sclIntValue -ge 2) { $neutralFactors.Add("SCL=$scl (Low Spam Probability / Caution)") }
        else { $positiveFactors.Add("SCL=$scl (Not Spam / Trusted)") }
    }
    else {
        $neutralFactors.Add("SCL Not Found")
    }
    
    # Authentication Analysis
    if ($authResults) {
        if ($authResults -match "(?i)spf=pass") { $positiveFactors.Add("SPF Pass") }
        elseif ($authResults -match "(?i)spf=fail") { $riskFactors.Add("SPF Fail") }
        elseif ($authResults -match "(?i)spf=softfail") { $neutralFactors.Add("SPF Softfail") }
        else { $neutralFactors.Add("SPF Neutral/None/Error") }
        
        if ($authResults -match "(?i)dkim=pass") { $positiveFactors.Add("DKIM Pass") }
        elseif ($authResults -match "(?i)dkim=fail") { $riskFactors.Add("DKIM Fail") }
        else { $neutralFactors.Add("DKIM None/Neutral/Error") }
        
        if ($authResults -match "(?i)dmarc=pass") { $positiveFactors.Add("DMARC Pass") }
        elseif ($authResults -match "(?i)dmarc=fail") { $riskFactors.Add("DMARC Fail") }
        else { $neutralFactors.Add("DMARC None/Error") }

        if ($authResults -match "(?i)compauth=fail") { $riskFactors.Add("CompAuth Fail") }
        elseif ($authResults -match "(?i)compauth=pass") { $positiveFactors.Add("CompAuth Pass") }
    }
    else {
        $riskFactors.Add("Authentication-Results Header Missing")
    }
    
    # Category Analysis from Forefront
    if ($forefrontReport) {
        if ($forefrontReport -match "(?i)CAT:SPAM") { $riskFactors.Add("Forefront CAT:SPAM") }
        if ($forefrontReport -match "(?i)CAT:PHSH") { $riskFactors.Add("Forefront CAT:PHISHING") }
        if ($forefrontReport -match "(?i)CAT:MALW") { $riskFactors.Add("Forefront CAT:MALWARE") }
        if ($forefrontReport -match "(?i)SFV:SPM") { $riskFactors.Add("Forefront SFV:SPM (Marked as Spam)") }
        if ($forefrontReport -match "(?i)CAT:BULK") { $neutralFactors.Add("Forefront CAT:BULK") }
    }
    
    # Delivery Destination
    $deliveryDestText = "Unknown"
    if ($deliveryInfo -match "(?i)dest:([A-Z])") {
        $destCode = $matches[1].ToUpper()
        $deliveryDestText = switch ($destCode) {
            "I" { "Inbox" }
            "J" { "Junk Email Folder" }
            "D" { "Deleted" }
            "Q" { "Quarantined" }
            "S" { "Skipped" }
            default { "Unknown ($destCode)" }
        }
        if ($destCode -eq "J" -or $destCode -eq "D" -or $destCode -eq "Q") { $riskFactors.Add("Delivered to $deliveryDestText") }
        elseif ($destCode -eq "I") { $positiveFactors.Add("Delivered to Inbox") }
    }
     
    # Overall Verdict
    Write-Host "  MESSAGE VERDICT:" -ForegroundColor Yellow
    Write-Host "  $(New-Object System.String -ArgumentList @('─', 50))" -ForegroundColor Yellow
    
    if ($sclIntValue -ge 9 -or ($riskFactors -match "DMARC Fail" -and $riskFactors -match "CompAuth Fail")) {
        Write-Host "  🚨 HIGH RISK / SPAM DETECTED" -ForegroundColor Red -BackgroundColor Black
        Write-Host "     This message shows strong indicators of being spam or malicious." -ForegroundColor Red
    }
    elseif ($sclIntValue -ge 5 -or $riskFactors.Count -gt $positiveFactors.Count + $neutralFactors.Count) {
        Write-Host "  ⚠️  POTENTIAL RISK / LIKELY SPAM" -ForegroundColor Yellow
        Write-Host "     This message has several characteristics of spam or unwanted mail." -ForegroundColor Yellow
    }
    elseif ($positiveFactors.Count -gt $riskFactors.Count) {
        Write-Host "  ✅ LIKELY LEGITIMATE" -ForegroundColor Green
        Write-Host "     This message appears to be legitimate based on key checks." -ForegroundColor Green
    }
    else {
        Write-Host "  🔎 MIXED RESULTS / CAUTION ADVISED" -ForegroundColor White
        Write-Host "     Review details carefully. Some checks passed, others raised concerns." -ForegroundColor White
    }
    
    Write-Host "`n  DELIVERY INFORMATION:" -ForegroundColor Cyan
    Write-Host "  Final Destination (best guess): $deliveryDestText" -ForegroundColor White
    
    Write-Host "`n  RISK FACTORS IDENTIFIED:" -ForegroundColor Red
    if ($riskFactors.Count -eq 0) {
        Write-Host "    None significant." -ForegroundColor Green
    }
    else {
        foreach ($risk in $riskFactors) {
            Write-Host "    ❌ $risk" -ForegroundColor Red
        }
    }
    
    Write-Host "`n  POSITIVE / NEUTRAL INDICATORS:" -ForegroundColor Green
    if ($positiveFactors.Count -eq 0 -and $neutralFactors.Count -eq 0) {
        Write-Host "    Few positive or neutral indicators found." -ForegroundColor Yellow
    }
    else {
        foreach ($positive in $positiveFactors) {
            Write-Host "    ✅ $positive" -ForegroundColor Green
        }
        foreach ($neutral in $neutralFactors) {
            Write-Host "    ⚪ $neutral" -ForegroundColor Gray # Using Gray for neutral
        }
    }
    
    # Recommendations
    Write-Host "`n  RECOMMENDATIONS:" -ForegroundColor Yellow
    if ($sclIntValue -ge 5 -or ($riskFactors -match "FAIL" -and $riskFactors.Count -gt 1) ) {
        Write-Host "    • High Risk: If this is unexpected, treat with extreme caution. Do not click links or open attachments." -ForegroundColor Red
        Write-Host "    • If legitimate but misidentified (false positive), sender should review their SPF, DKIM, DMARC records, and sending practices." -ForegroundColor Yellow
        Write-Host "    • Consider adding sender to safe sender list if trusted and consistently misclassified, but only after verification." -ForegroundColor Yellow
    }
    elseif ($riskFactors.Count -gt 0) {
        Write-Host "    • Caution: Some risk factors identified. Verify sender and content if message is unexpected." -ForegroundColor Yellow
        Write-Host "    • If sender is known, they may need to correct authentication issues (SPF/DKIM/DMARC)." -ForegroundColor Yellow
    }
    else {
        Write-Host "    • Message appears largely legitimate based on technical analysis." -ForegroundColor Green
        Write-Host "    • Always exercise standard security precautions (e.g., be wary of unexpected requests or attachments)." -ForegroundColor Gray
    }
}

function Show-References {
    Write-SubHeader "USEFUL REFERENCES"
    
    $references = @(
        "Microsoft 365 Security Documentation: https://learn.microsoft.com/en-us/microsoft-365/security/",
        "Exchange Online Protection (EOP): https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/eop-about",
        "Anti-spam message headers in Microsoft 365: https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/message-headers-eop-mdo",
        "Spam confidence levels (SCL): https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/spam-confidence-levels",
        "Bulk Complaint Level (BCL) values: https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/bulk-complaint-level-values",
        "How Microsoft 365 uses SPF: https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/how-office-365-uses-spf-to-prevent-spoofing",
        "How Microsoft 365 uses DKIM: https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/how-office-365-uses-dkim-to-validate-outbound-email",
        "How Microsoft 365 uses DMARC: https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/use-dmarc-to-validate-email",
        "RFC 5322 (Internet Message Format): https://tools.ietf.org/html/rfc5322",
        "RFC 7208 (SPF): https://tools.ietf.org/html/rfc7208",
        "RFC 6376 (DKIM): https://tools.ietf.org/html/rfc6376",
        "RFC 7489 (DMARC): https://tools.ietf.org/html/rfc7489"
    )
    
    foreach ($ref in $references) {
        Write-Host "  • $ref" -ForegroundColor DarkGray
    }
}

Clear-Host
# Run all analyses
try {
    # Display usage information if verbose is enabled
    if ($Global:VerboseOutput) {
        Write-Header "VERBOSE ANALYSIS MODE"
        Write-ColoredText "Running in VERBOSE mode with enhanced output" "Magenta"
        Write-ColoredText "You'll see additional technical details and explanations for each header" "Magenta"
        Write-Host "`n"
    }
    
    Analyze-Authentication
    Analyze-SpamFiltering # Includes SCL, Forefront, MS-Antispam, ATP properties, Mailbox Delivery
    Analyze-SafeAttachments
    Analyze-SafeLinks
    Analyze-TransportRules
    Analyze-GeneralHeaders
    Show-Summary
    Show-References
    
    Write-Header "ANALYSIS COMPLETE"
    Write-ColoredText "Report generated successfully!" "Green"
    
    # Add parameter hints
    Write-Host "`nTIP: For even more detailed analysis, run with the -Verbose and -ShowRawHeaders parameters:" -ForegroundColor Gray
    Write-Host "     .\email-header-report.ps1 -HeaderFilePath $HeaderFilePath -Verbose -ShowRawHeaders" -ForegroundColor DarkGray
}
catch {
    Write-Error "An error occurred during analysis: $($_.Exception.Message)"
    Write-Warning "Script execution halted due to an error. Some parts of the report may be missing."
    exit 1
}

