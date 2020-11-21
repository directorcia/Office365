<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Source - https://github.com/directorcia/Office365/blob/master/win10-def-get.ps1

Description - Report Windows Defender configuration

Prerequisites = 

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor="red"

Clear-Host
write-host -foregroundcolor $systemmessagecolor "Script started`n"

write-host -ForegroundColor Gray -backgroundcolor blue "Latest signature and engine versions"
## https://docs.microsoft.com/en-us/previous-versions/windows/desktop/defender/msft-mpcomputerstatus#properties
$localdefender = Get-MpComputerStatus
write-host -foregroundcolor $processmessagecolor "Read latest version from web page - https://www.microsoft.com/en-us/wdsi/defenderupdates"
$info=invoke-webrequest -Uri "https://www.microsoft.com/en-us/wdsi/defenderupdates" -UseBasicParsing -DisableKeepAlive
write-host -foregroundcolor $processmessagecolor "Find values`n"
$check = $info.RawContent -match '<li>Version: <span>.*'
$ver = $Matches.values
$ver=$ver.replace("<li>Version: <span>","")
$version=$ver.replace("</span></li>","").trim()

if ($localdefender.AntispywareSignatureVersion -match $version ) {
    write-host -foregroundcolor $processmessagecolor "Version:",$localdefender.AntispywareSignatureVersion
}
else {
    for($i = 0; $i -lt $version.length; $i++) {
        if ($version[$i] -notmatch $localdefender.AntispywareSignatureVersion[$i]) {
            if (-not $skip) {
                if ([int]::Parse($version[$i]) -lt [int]::Parse($localdefender.AntispywareSignatureVersion[$i])) {
                    $current = $true
                    $skip = $true
                } else {
                    $current = $false
                    $skip = $true
                }
            }
        }
    }
    if ($current) {
        write-host -foregroundcolor $processmessagecolor "Local version =",$localdefender.AntispywareSignatureVersion
        write-host -foregroundcolor gray "    is more current that reported latest version:",$version    
    } else {
        write-host -foregroundcolor $errormessagecolor "Local version =",$localdefender.AntispywareSignatureVersion
        write-host -forgroundcolor $errormessagecolor "    is less current that web latest version:",$version
    }
}

$check = $info.RawContent -match '<li>Engine version: <span>.*'
$ver = $Matches.values
$ver=$ver.replace("<li>Engine Version: <span>","")
$engine=$ver.replace("</span></li>","").trim()
if ($localdefender.AMEngineVersion -match $engine ) {
    write-host -foregroundcolor $processmessagecolor "Engine version =",$localdefender.AMEngineVersion
}
else {
    write-host -foregroundcolor $errormessagecolor "Engine version =",$localdefender.AMEngineVersion,"["$engine"]"
}

$check = $info.RawContent -match '<li>Platform version: <span>.*'
$ver = $Matches.values
$ver=$ver.replace("<li>Platform Version: <span>","")
$platform=$ver.replace("</span></li>","").trim()
if ($localdefender.AMServiceVersion -like $platform ) {
    write-host -foregroundcolor $processmessagecolor "Platform version =",$localdefender.AMServiceVersion
}
else {
    write-host -foregroundcolor $errormessagecolor "Platform version =",$localdefender.AMServiceVersion,"["$platform"]"
}

$check = $info.RawContent -match '<li>Released: <span id=.*'
$ver = $Matches.values
$ver=$ver.replace('<li>Released: <span id="dateofrelease">',"")
$release=$ver.replace("</span></li>","").trim()
write-host -foregroundcolor $processmessagecolor "Released (UTC) =",$release
write-host -foregroundcolor gray "    Last local update:",$localdefender.AntivirusSignatureLastUpdated
write-host -foregroundcolor $processmessagecolor "Anti-Malware Mode =",$localdefender.AMRunningMode
write-host -foregroundcolor $processmessagecolor "Anti-Malware Service enabled =",$localdefender.AMServiceEnabled
write-host -foregroundcolor $processmessagecolor "Anti-Spyware Service enabled =",$localdefender.AntispywareEnabled
write-host -foregroundcolor $processmessagecolor "Anti-Virus Service enabled =",$localdefender.AntivirusEnabled
write-host -foregroundcolor $processmessagecolor "Behavior Monitoring enabled =",$localdefender.BehaviorMonitorEnabled
write-host -foregroundcolor $processmessagecolor "Scan all downloaded files and attachments =",$localdefender.IoavProtectionEnabled

if ($localdefender.IsTamperProtected) {
    write-host -foregroundcolor $processmessagecolor "Is tamper protected enabled =",$localdefender.IsTamperProtected
} else {
    write-host -foregroundcolor $errormessagecolor "Is tamper protected enabled =",$localdefender.IsTamperProtected
}

write-host -foregroundcolor $processmessagecolor "NRI Engine enabled =",$localdefender.NISEnabled
write-host -foregroundcolor $processmessagecolor "On Access Protection enabled =",$localdefender.OnAccessProtectionEnabled
write-host -foregroundcolor $processmessagecolor "Real Time Protection enabled =",$localdefender.RealTimeProtectionEnabled

write-host -foregroundcolor $systemmessagecolor "`nScript completed`n"