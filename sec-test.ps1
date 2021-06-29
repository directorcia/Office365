param(                        
    [switch]$debug = $false    ## if -debug parameter don't prompt for input
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - 
Source - 

Prerequisites = Windows 10

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

Clear-Host
if ($debug) {
    write-host -foregroundcolor $processmessagecolor "Create log file ..\win10-sec-test.txt`n"
    Start-transcript "..\win10-sec-test.txt" | Out-Null                                   ## Log file created in current directory that is overwritten on each run
}

write-host -foregroundcolor $systemmessagecolor "Windows 10 security test script started`n"
write-host -ForegroundColor white -backgroundcolor blue "--- Download EICAR file ---"
$dldetect=$true
write-host -foregroundcolor $processmessagecolor "Download eicar.com.txt file to current directory"
Invoke-WebRequest -Uri https://secure.eicar.org/eicar.com.txt -OutFile .\eicar.com.txt
write-host -foregroundcolor $processmessagecolor "Verify eicar.com.txt file in current directory"
try {
    read-content .\eicar.com.txt
}
catch {
    write-host -foregroundcolor $processmessagecolor "eicar.com.txt file download not found"
    $dldetect=$false
}
if ($dldetect) {
    write-host -foregroundcolor $warningmessagecolor "eicar.com.txt file download found"
    $dlexist = $true
    try {
        $dlsize = (Get-ChildItem ".\eicar.com.txt").Length
    }
    catch {
        $dlexist = $false
        write-host -foregroundcolor $processmessagecolor "eicar.com.txt download not found"
    }
    if ($dlexist) {
        if ($dlsize -ne 0) {
            write-host -foregroundcolor $errormessagecolor "eicar.com.txt download file length > 0"
        }
    }

}
write-host -ForegroundColor white -backgroundcolor blue "`n--- Create EICAR file in current directory ---"
set-content .\eicar1.com.txt:EICAR ‘X5O!P%@AP[4\PZX54(P^)7CC)7}$EICAR-STANDARD-ANTIVIRUS-TEST-FILE!$H+H*’ 
write-host -foregroundcolor $processmessagecolor "eicar1.com.txt created"
$crdetect = $false
write-host -foregroundcolor $processmessagecolor "Check Windows Defender logs for eicar1 report"
$results = get-mpthreatdetection | sort-object initialdetectiontime -Descending
$item = 0
foreach ($result in $results) {
    if ($result.actionsuccess -and ($result.resources -match "eicar1")) {
        ++$item
        write-host "`nItem =",$item
        write-host "Initial detection time =",$result.initialdetectiontime
        write-host "Process name =",$result.processname
        write-host -foregroundcolor $processmessagecolor "Resource = ",$result.resources
        $crdetect = $true
    }
}
if ($crdetect) {
    write-host -foregroundcolor $processmessagecolor "`nEICAR file creation detected"
}
else {
    write-host -foregroundcolor $errormessagecolor "`nEICAR file creation not detected"
}
$crdtect = $true
try {
    $fileproperty = get-itemproperty .\eicar1.com.txt
}
catch {
    write-host -foregroundcolor $processmessagecolor "eicar1.com.txt file not detected"
    $crdtect = $false
}
if ($crdetect) {
    if ($fileproperty.Length -eq 0) {
        write-host -foregroundcolor $processmessagecolor "eicar1.com.txt detected with file size = 0"
    }
    else {
        write-host -foregroundcolor $errormessagecolor "eicar1.com.txt detected but file size is not 0"
    }
}
write-host -ForegroundColor white -backgroundcolor blue "`n--- Attempt LSASS process dump ---"
$result = test-path ".\procdump.exe"
$procdump = $true
if (-not $result) {
    write-host -foregroundcolor $warningmessagecolor "SysInternals procdump.exe not found in current directory"
    do {
        $result = Read-host -prompt "Download SysInternals procdump (Y/N)?"
    } until (-not [string]::isnullorempty($result))
    if ($result -eq 'Y' -or $result -eq 'y') {
        write-host -foregroundcolor $processmessagecolor "Download procdump.zip to current directory"
        invoke-webrequest -uri https://download.sysinternals.com/files/Procdump.zip -outfile .\procdump.zip
        write-host -foregroundcolor $processmessagecolor "Expand procdump.zip file to current directory"
        Expand-Archive -LiteralPath .\procdump.zip -DestinationPath .\
        $result = test-path ".\procdump.exe"
        if ($result) {
            write-host -foregroundcolor $processmessagecolor "procdump.exe found in current directory"
        }
        else {
            write-host -foregroundcolor $errormessagecolor "procdump.exe not found in current directory"
            $procdump = $false
        }
    }
    else {
        $procdump = $false
    }
}
if ($procdump) {
    $accessdump = $true
    try {
        .\procdump.exe -ma lsass.exe lsass.dmp    
    }
    catch {
        if ($error[0] -match "Access is denied") {
            write-host -foregroundcolor $processmessagecolor "Access denied - Unable to process dump"
            $accessdump = $false
        }
        else {
            write-host -foregroundcolor $processmessagecolor $error[0]
        }
    }
    if ($accessdump) {
        write-host -foregroundcolor $errormessagecolor "Able to process dump or other error"
    }
}
write-host -ForegroundColor white -backgroundcolor blue "`n--- Generate failed login ---"
do {
    $username = Read-host -prompt "Enter use email address to generate failed login"
} until (-not [string]::isnullorempty($result))
$password="1"
$URL = "https://login.microsoft.com"
$BodyParams = @{'resource' = 'https://graph.windows.net'; 'client_id' = '1b730954-1685-4b74-9bfd-dac224a7b894' ; 'client_info' = '1' ; 'grant_type' = 'password' ; 'username' = $username ; 'password' = $password ; 'scope' = 'openid'}
$PostHeaders = @{'Accept' = 'application/json'; 'Content-Type' =  'application/x-www-form-urlencoded'}
try {
    $webrequest = Invoke-WebRequest $URL/common/oauth2/token -Method Post -Headers $PostHeaders -Body $BodyParams -ErrorVariable RespErr
} 
catch {
    switch -wildcard ($RespErr)
    {
        "*AADSTS50126*" {write-host -foregroundcolor $processmessagecolor "Error validating credentials due to invalid username or password as expected"; break}
        "*AADSTS50034*" {write-host -foregroundcolor $warningmessagecolor "User $username doesn't exist"; break}
        "*AADSTS50053*" {write-host -foregroundcolor $warningmessagecolor "User $username appears to be locked"; break}
        "*AADSTS50057*" {write-host -foregroundcolor $warningmessagecolor "User $username appears to be disabled"; break}
        default {write-host -foregroundcolor $warningmessagecolor "Unknow error for user $username"}
    }
}
if ($debug) {
    Stop-Transcript | Out-Null
}