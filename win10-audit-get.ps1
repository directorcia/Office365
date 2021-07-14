param(                         
    [switch]$debug = $false       ## if -debug create a log file
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Report Local Windows 10 machine Audit log settings

Source - https://github.com/directorcia/Office365/blob/master/win10-audit-get.ps1

Prerequisites = 1
1. Ensure aipservice module installed or updated

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host
if ($debug) {
    write-host "Script activity logged at ..\win10-audit-get.txt"
    start-transcript "..\win10-audit-get.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}

write-host -foregroundcolor $systemmessagecolor "Windows 10 Get Audit Log Settings script started`n"

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $output = invoke-expression "auditpol /get /category:*"

    $value ="Credential Validation"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success and Failure") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Account Logon Audit Credential Validation (Device) =",$outputmatch.trim()

    $value ="Kerberos Authentication Service"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "No Auditing") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Account Logon Audit Kerberos Authentication Service (Device) =",$outputmatch.trim()

    $value ="Account Lockout"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Failure") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Account Logon Logoff Audit Account Lockout (Device) =",$outputmatch.trim()

    $value ="Group Membership"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Account Logon Logoff Audit Group Membership (Device) =",$outputmatch.trim()

    $value ="  Logon"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success and Failure") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Account Logon Logoff Audit Logon (Device) =",$outputmatch.trim()

    $value ="Other Logon/Logoff Events"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success and Failure") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Audit Other Logon Logoff Events (Device) =",$outputmatch.trim()

    $value ="Special Logon"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Audit Special Logon (Device) =",$outputmatch.trim()

    $value ="Security Group Management"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Audit Security Group Management (Device) =",$outputmatch.trim()

    $value ="User Account Management"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success and Failure") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Audit User Account Management (Device) =",$outputmatch.trim()

    $value ="Plug and Play Events"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Detailed Tracking Audit PNP Activity (Device) =",$outputmatch.trim()

    $value ="Process Creation"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Detailed Tracking Audit Process Creation (Device) =",$outputmatch.trim()

    $value ="Detailed File Share"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Failure") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Object Access Audit Detailed File Share (Device) =",$outputmatch.trim()

    $value ="  File Share"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success and Failure") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Audit File Share Access (Device) =",$outputmatch.trim()

    $value ="Other Object Access Events"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success and Failure") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Object Access Audit Other Object Access Events (Device) =",$outputmatch.trim()

    $value ="Removable Storage"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success and Failure") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Object Access Audit Removable Storage (Device) =",$outputmatch.trim()

    $value ="Authentication Policy Change"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Audit Authentication Policy Change (Device) =",$outputmatch.trim()

    $value ="MPSSVC Rule-Level Policy Change"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success and Failure") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Policy Change Audit MPSSVC Rule Level Policy Change (Device) =",$outputmatch.trim()

    $value ="Other Policy Change Events"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Failure") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Policy Change Audit Other Policy Change Events (Device) =",$outputmatch.trim()

    $value ="Audit Policy Change"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Audit Changes to Audit Policy (Device) =",$outputmatch.trim()

    $value ="  Sensitive Privilege Use"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success and Failure") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Privilege Use Audit Sensitive Privilege Use (Device) =",$outputmatch.trim()

    $value ="Other System Events"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success and Failure") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "System Audit Other System Events (Device) =",$outputmatch.trim()

    $value ="Security State Change"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }

    write-host -foregroundcolor $displaycolor "System Audit Security State Change (Device) =",$outputmatch.trim()
    $value ="Security System Extension"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "Audit Security System Extension (Device) =",$outputmatch.trim()

    $value ="System Integrity"
    $outputmatch = ($output -match $value) -replace $value,""
    if ($outputmatch.trim() -eq "Success and Failure") {
        $displaycolor = $processmessagecolor
    }
    else {
        $displaycolor = $errormessagecolor
    }
    write-host -foregroundcolor $displaycolor "System Audit System Integrity (Device) =",$outputmatch.trim()
}
else {
    write-host -foregroundcolor $warningmessagecolor "Audit Policy results only available when script run as Administrator`n"
}
write-host -foregroundcolor $systemmessagecolor "`nWindows 10 Get Audit Log Settings script finished"
if ($debug) {
    Stop-Transcript | Out-Null
}