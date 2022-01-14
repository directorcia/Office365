param(                        
    [switch]$debug = $false    ## if -debug parameter don't prompt for input
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Perform security tests in your environment
Source - https://github.com/directorcia/Office365/blob/master/mde-offboard-check.ps1

Prerequisites = Windows 10

More scripts at www.ciaopsacademy.com

#>

#Region Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"
#EndRegion Variables

<#          Main                #>
Clear-Host
if ($debug) {       # If -debug command line option specified record log file in parent
    write-host -foregroundcolor $processmessagecolor "Create log file ..\mde-offboard-check.txt`n"
    Start-transcript "..\mde-offboard-check.txt" | Out-Null                                   ## Log file created in current directory that is overwritten on each run
}
write-host -foregroundcolor $systemmessagecolor "Microsoft Defender for Endpoint Offboarding check script started`n"
if (-not $debug) {
    Write-host -foregroundcolor $warningmessagecolor "    * use the -debug parameter on the command line to create an execution log file for this script`n"
}

try {
    $ob = Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows Advanced Threat Protection\Status\" -ErrorAction silentlycontinue    
}
catch {
    $ob=$null
} 
if (-not [string]::isnullorempty($ob.OnboardingState)) {
    switch ($ob.OnboardingState) {
        0 {write-host -foregroundcolor $processmessagecolor "Defender for Endpoint = Offboarded";break}
        1 {write-host -foregroundcolor $processmessagecolor "Defender for Endpoint = Onboarded";break}
        default {write-host -foregroundcolor $errormessagecolor "Defender for Endpoint = Unknown value";break}
    }
} else {
    write-host -foregroundcolor $errormessagecolor "Defender for Endpoint = Status not found"
}
write-host -foregroundcolor $systemmessagecolor "`nScript completed`n"

if ($debug){
    Stop-Transcript | Out-Null
}