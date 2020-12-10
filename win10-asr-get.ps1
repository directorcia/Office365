<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Source - https://github.com/directorcia/Office365/blob/master/win10-asr-get.ps1

Description - Report Device Attack Surface Reduction (ASR) settings

Prerequisites -

References:

- ASR Overview - https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/overview-attack-surface-reduction
- Reduce attack surfaces with attack surface reduction rules - https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/attack-surface-reduction
- ASR FAQ - https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/attack-surface-reduction-faq


More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor="red"

Clear-Host
write-host -foregroundcolor $systemmessagecolor "Script started`n"

$asrrules = @()
$asrrules += [PSCustomObject]@{
    Name = "Block executable content from email client and webmail";
    GUID = "BE9BA2D9-53EA-4CDC-84E5-9B1EEEE46550"
    ## Reference - https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/attack-surface-reduction#block-executable-content-from-email-client-and-webmail
}
$asrrules += [PSCustomObject]@{
    Name = "Block all Office applications from creating child processes";
    GUID = "D4F940AB-401B-4EFC-AADC-AD5F3C50688A"
    ## Reference - https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/attack-surface-reduction#block-all-office-applications-from-creating-child-processes
}
$asrrules += [PSCustomObject]@{
    Name = "Block Office applications from creating executable content";
    GUID = "3B576869-A4EC-4529-8536-B80A7769E899"
    ## Reference - https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/attack-surface-reduction#block-office-applications-from-creating-executable-content
}
$asrrules += [PSCustomObject]@{
    Name = "Block Office applications from injecting code into other processes";
    GUID = "75668C1F-73B5-4CF0-BB93-3ECF5CB7CC84"
    ## Reference - https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/attack-surface-reduction#block-office-applications-from-injecting-code-into-other-processes
}
$asrrules += [PSCustomObject]@{
    Name = "Block JavaScript or VBScript from launching downloaded executable content";
    GUID = "D3E037E1-3EB8-44C8-A917-57927947596D"
    ## Reference - https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/attack-surface-reduction#block-javascript-or-vbscript-from-launching-downloaded-executable-content
}
$asrrules += [PSCustomObject]@{
    Name = "Block execution of potentially obfuscated scripts";
    GUID = "5BEB7EFE-FD9A-4556-801D-275E5FFC04CC"
    ## Reference - https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/attack-surface-reduction#block-execution-of-potentially-obfuscated-scripts
}
$asrrules += [PSCustomObject]@{
    Name = "Block Win32 API calls from Office macros";
    GUID = "92E97FA1-2EDF-4476-BDD6-9DD0B4DDDC7B"
    ## Reference - https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/attack-surface-reduction#block-win32-api-calls-from-office-macros
}
$asrrules += [PSCustomObject]@{
    Name = "Block executable files from running unless they meet a prevalence, age, or trusted list criterion";
    GUID = "01443614-cd74-433a-b99e-2ecdc07bfc25"
    ## Reference - https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/attack-surface-reduction#block-executable-files-from-running-unless-they-meet-a-prevalence-age-or-trusted-list-criterion
}
$asrrules += [PSCustomObject]@{
    Name = "Use advanced protection against ransomware";
    GUID = "c1db55ab-c21a-4637-bb3f-a12568109d35"
    ## reference - https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/attack-surface-reduction#use-advanced-protection-against-ransomware
}
$asrrules += [PSCustomObject]@{
    Name = "Block credential stealing from the Windows local security authority subsystem (lsass.exe)";
    GUID = "9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2"
    ## https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/attack-surface-reduction#block-credential-stealing-from-the-windows-local-security-authority-subsystem
}
$asrrules += [PSCustomObject]@{
    Name = "Block process creations originating from PSExec and WMI commands";
    GUID = "d1e49aac-8f56-4280-b9ba-993a6d77406c"
    ## https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/attack-surface-reduction#block-process-creations-originating-from-psexec-and-wmi-commands
}
$asrrules += [PSCustomObject]@{
    Name = "Block untrusted and unsigned processes that run from USB";
    GUID = "b2b3f03d-6a65-4f7b-a9c7-1c7ef74a9ba4"
    ## https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/attack-surface-reduction#block-untrusted-and-unsigned-processes-that-run-from-usb
}
$asrrules += [PSCustomObject]@{
    Name = "Block Office communication application from creating child processes";
    GUID = "26190899-1602-49e8-8b27-eb1d0a1ce869"
    ## Reference - https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/attack-surface-reduction#block-office-communication-application-from-creating-child-processes
}
$asrrules += [PSCustomObject]@{
    Name = "Block Adobe Reader from creating child processes";
    GUID = "7674ba52-37eb-4a4f-a9a1-f0f9a1619a2c"
    ## Reference - https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/attack-surface-reduction#block-adobe-reader-from-creating-child-processes
}
$asrrules += [PSCustomObject]@{
    Name = "Block persistence through WMI event subscription";
    GUID = "e6db77e5-3df2-4cf1-b95a-636979351e5b"
    ## Reference - https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-atp/attack-surface-reduction#block-persistence-through-wmi-event-subscription
}

$enabledvalues = "Not Enabled", "Enabled", "Audit"
$displaycolor = $errormessagecolor, $processmessagecolor, $warningmessagecolor
## https://docs.microsoft.com/en-us/powershell/module/defender/?view=win10-ps
$results = Get-MpPreference
write-host -ForegroundColor Gray -backgroundcolor blue "Attack Surface Reduction Rules`n"
$count = 0 

if (-not [string]::isnullorempty($results.AttackSurfaceReductionRules_ids)) {
    foreach ($id in $asrrules.GUID) {
    ##foreach ($id in $results.AttackSurfaceReductionRules_Ids) {
        $enabled = $results.AttackSurfaceReductionRules_actions[$count]
        switch ($id) {
            "BE9BA2D9-53EA-4CDC-84E5-9B1EEEE46550" {$index=0;break}
            "D4F940AB-401B-4EFC-AADC-AD5F3C50688A" {$index=1;break}
            "3B576869-A4EC-4529-8536-B80A7769E899" {$index=2;break}
            "75668C1F-73B5-4CF0-BB93-3ECF5CB7CC84" {$index=3;break}
            "D3E037E1-3EB8-44C8-A917-57927947596D" {$index=4;break}
            "5BEB7EFE-FD9A-4556-801D-275E5FFC04CC" {$index=5;break}
            "92E97FA1-2EDF-4476-BDD6-9DD0B4DDDC7B" {$index=6;break}
            "01443614-cd74-433a-b99e-2ecdc07bfc25" {$index=7;break}
            "c1db55ab-c21a-4637-bb3f-a12568109d35" {$index=8;break}
            "9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2" {$index=9;break}
            "d1e49aac-8f56-4280-b9ba-993a6d77406c" {$index=10;break}
            "b2b3f03d-6a65-4f7b-a9c7-1c7ef74a9ba4" {$index=11;break}
            "26190899-1602-49e8-8b27-eb1d0a1ce869" {$index=12;break}
            "7674ba52-37eb-4a4f-a9a1-f0f9a1619a2c" {$index=13;break}
            "e6db77e5-3df2-4cf1-b95a-636979351e5b" {$index=14;break}
        }
        switch ($enabled) {
            0 {write-host -foregroundcolor $displaycolor[$enabled] $asrrules[$index].name"="$enabledvalues[$enabled]; break}
            1 {write-host -foregroundcolor $displaycolor[$enabled] $asrrules[$index].name"="$enabledvalues[$enabled]; break}
            2 {write-host -foregroundcolor $displaycolor[$enabled] $asrrules[$index].name"="$enabledvalues[$enabled]; break}
            default {write-host -foregroundcolor $errormessagecolor $asrrules[$index].name"= Not found"; break}
        }
        $count++
    }
}
else {
    write-host -foregroundcolor $errormessagecolor $asrrules.count"ASR rules empty"
}

write-host -foregroundcolor $systemmessagecolor "`nScript completed`n"