param(                        
    [switch]$debug = $false    ## if -debug parameter don't prompt for input
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Perform security tests in your environment
Source - https://github.com/directorcia/Office365/blob/master/sec-test.ps1
Documentation - https://blog.ciaops.com/2021/06/29/is-security-working-powershell-script/
Resources - https://demo.wd.microsoft.com/

Prerequisites = Windows 10, OFfice, valid Microsoft 365 login, endpoint security

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

function downloadfile() {
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
}

function createfile() {
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
            write-host -foregroundcolor $processmessagecolor "Removing file .\EICAR1.COM.TXT"
            Remove-Item .\eicar1.com.txt
        }
        else {
            write-host -foregroundcolor $errormessagecolor "eicar1.com.txt detected but file size is not 0"
        }
    }
}

function inmemorytest(){
    write-host -ForegroundColor white -backgroundcolor blue "`n--- In memory test ---"
    $memdetect = $false
    $errorfile = ".\sec-test-$(get-date -f yyyyMMddHHmmss).txt"     # unique output file
    $s1 = ‘AMSI Test Sample: 7e72c3ce'             # first half of EICAR string
    $s2 = '-861b-4339-8740-0ac1484c1386’           # second half of EICAR string
    $s3=($s1+$s2)                                  # combined EICAR string in one variable
    $encodedcommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($s3)) # need to encode command so not detected and block in this script
    write-host -foregroundcolor $processmessagecolor "Launch Powershell child process to output EICAR string to console"
    Start-Process powershell -ArgumentList "-EncodedCommand $encodedcommand" -wait -WindowStyle Hidden -redirectstandarderror $errorfile
    write-host -foregroundcolor $processmessagecolor "Attempt to read output file created by child process"
    try {
        get-content $errorfile -ErrorAction Stop        # look at child process error output
    }
    catch {     # if unable to open file this is because EICAR strng found in there
        write-host -foregroundcolor $processmessagecolor "In memory test SUCCEEDED"
        write-host -foregroundcolor $processmessagecolor "Removing file $errorfile"
        remove-item $errorfile      # remove child process error output file
        $memdetect = $true          # set detection state = found
    }
    if (-not $memdetect) {
        write-host -foregroundcolor $errormessagecolor "In memory test FAILED. Recommended action = review file $errorfile"
    }
}

function processdump() {
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
            Expand-Archive -LiteralPath .\procdump.zip -DestinationPath .\ -Force
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
            write-host -nonewline -foregroundcolor $processmessagecolor "Attempt process dump in current user context = "
            $result = .\procdump.exe -mm lsass.exe lsass.dmp -accepteula    
        }
        catch {
            if ($error[0] -match "Access is denied") {
                write-host -foregroundcolor $processmessagecolor "Access denied - Unable to process dump in current user context"
                $accessdump = $false
            }
            else {
                write-host -foregroundcolor $processmessagecolor $error[0]
            }
        }
        if ($result -match "Access is denied") {
            write-host -foregroundcolor $processmessagecolor "Access denied - Unable to process dump in current user context"
            $accessdump = $false
        }
        try {
            write-host -nonewline -foregroundcolor $processmessagecolor "Attempt process dump in admin context = "
            $error.Clear()      # Clear any existing errors
            start-process -filepath ".\procdump.exe" -argumentlist "-mm -o lsass.exe lsass.dmp" -verb runas -wait -WindowStyle Hidden
        }
        catch {
            if ($error[0] -match "Access is denied") {
                write-host -foregroundcolor $processmessagecolor "Access denied - Unable to process dump in admin context"
                $accessdump = $false
            }
        }
        $result = test-path ".\lsass.dmp"
        if ($result) {
            write-host -foregroundcolor $errormessagecolor "Dump file found"
            $accessdump = $true
            write-host -foregroundcolor $processmessagecolor "Removing dump file .\LSASS.DMP"
            Remove-Item ".\lsass.dmp"
        }
        if ($accessdump) {
            write-host -foregroundcolor $errormessagecolor "Able to process dump or other error - test has FAILED"
        }
    }
}

function mimikatztest() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- Mimikatz test ---"
    $errorfile = ".\sec-test-$(get-date -f yyyyMMddHHmmss).txt"     # unique output file
    $s1 = 'invoke-'             # first half of command
    $s2 = 'mimikatz’           # second half of command
    $s3=($s1+$s2)                                  # combined EICAR string in one variable
    $encodedcommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($s3)) # need to encode command so not detected and block in this script
    write-host -foregroundcolor $processmessagecolor "Launch Powershell child process to output Mimikatz command string to console"
    Start-Process powershell -ArgumentList "-EncodedCommand $encodedcommand" -wait -WindowStyle Hidden -redirectstandarderror $errorfile
    write-host -foregroundcolor $processmessagecolor "Attempt to read output file created by child process"
    try {
        $result = get-content $errorfile -ErrorAction Stop        # look at child process error output
    }
    catch {     # if unable to open file this is because EICAR strng found in there
        write-host -foregroundcolor $errormessagecolor "[ERROR] Output file not found"
    }
    if ($result -match "This script contains malicious content and has been blocked by your antivirus software") {
        write-host -ForegroundColor $processmessagecolor "Malicious content and has been blocked by your antivirus software"
        remove-item $errorfile      # remove child process error output file
    }
    else {
        write-host -foregroundcolor $errormessagecolor "Malicious content NOT DETECTED = review file $errorfile"
    }   
}

function failedlogin() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- Generate failed login ---"
    do {
        $username = Read-host -prompt "Enter valid Microsoft 365 email address to generate failed login"
    } until (-not [string]::isnullorempty($username))
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
            default {write-host -foregroundcolor $warningmessagecolor "Unknown error for user $username"}
        }
    }
}

function officechildprocess() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- Office applications creating child processes ---"
    write-host -foregroundcolor $processmessagecolor "Download test Word document to current directory"
    Invoke-WebRequest -Uri https://demo.wd.microsoft.com/Content/TestFile_OfficeChildProcess_D4F940AB-401B-4EFC-AADC-AD5F3C50688A.docm -OutFile .\TestFile_OfficeChildProcess_D4F940AB-401B-4EFC-AADC-AD5F3C50688A.docm
    write-host -foregroundcolor $processmessagecolor "Open document using Word"
    Start-Process winword.exe -ArgumentList ".\TestFile_OfficeChildProcess_D4F940AB-401B-4EFC-AADC-AD5F3C50688A.docm"
    write-host -foregroundcolor $processmessagecolor "Ensure that a Run Time Error is displayed. If a command prompt appears the test has FAILED."
    write-host -foregroundcolor $processmessagecolor "Please close Word once complete."
    pause
    write-host -foregroundcolor $processmessagecolor "Delete .\TestFile_OfficeChildProcess_D4F940AB-401B-4EFC-AADC-AD5F3C50688A.docm"
    remove-item .\TestFile_OfficeChildProcess_D4F940AB-401B-4EFC-AADC-AD5F3C50688A.docm  
}

function officecreateexecutable() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- Office applications creating executables ---"
    write-host -foregroundcolor $processmessagecolor "Download test Word document to current directory"
    Invoke-WebRequest -Uri https://demo.wd.microsoft.com/Content/TestFile_Block_Office_applications_from_creating_executable_content_3B576869-A4EC-4529-8536-B80A7769E899.docm -OutFile .\TestFile_Block_Office_applications_from_creating_executable_content_3B576869-A4EC-4529-8536-B80A7769E899.docm
    write-host -foregroundcolor $processmessagecolor "Open document using Word"
    Start-Process winword.exe -ArgumentList ".\TestFile_Block_Office_applications_from_creating_executable_content_3B576869-A4EC-4529-8536-B80A7769E899.docm"
    write-host -foregroundcolor $processmessagecolor "Ensure that no executable runs. Please close Word once complete."
    pause
    write-host -foregroundcolor $processmessagecolor "Delete TestFile_Block_Office_applications_from_creating_executable_content_3B576869-A4EC-4529-8536-B80A7769E899.docm"
    remove-item .\TestFile_Block_Office_applications_from_creating_executable_content_3B576869-A4EC-4529-8536-B80A7769E899.docm  
}

function scriptlaunch() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- Impede Javascript and VBScript launch executables ---"
    write-host -foregroundcolor $processmessagecolor "Create DLTEST.JS file in current directory"
    set-content -Path .\dltest.js `
'// SCPT:xmlHttpRequest
var xmlHttp = WScript.CreateObject("MSXML2.XMLHTTP");
xmlHttp.open("GET", "https://www.bing.com", false);
xmlHttp.send();

// SCPT:JSRunsFile
var shell = WScript.CreateObject("WScript.Shell");
shell.Run("notepad.exe");'
    write-host -foregroundcolor $processmessagecolor "Execute DLTEST.JS file in current directory"
    start-process .\dltest.js
    write-host -foregroundcolor $processmessagecolor "A Windows Script Host error dialog box should have appeared."
    write-host -foregroundcolor $processmessagecolor "It should read:`n"
    write-host "Error: This script is blocked by IT policy"
    write-host "Code: 800A802E`n"
    write-host -foregroundcolor $warningmessagecolor "If NOTEPAD executed, then the system is vulnerable to this attack`n"
    pause
    write-host -foregroundcolor $processmessagecolor "Delete DLTEST.JS"
    remove-item .\dltest.js  

}

function officemacroimport() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- Block Win32 imports from Macro code in Office ---"
    write-host -foregroundcolor $processmessagecolor "Download test Word document to current directory"
    Invoke-WebRequest -Uri https://demo.wd.microsoft.com/Content/Block_Win32_imports_from_Macro_code_in_Office_92E97FA1-2EDF-4476-BDD6-9DD0B4DDDC7B.docm -OutFile .\Block_Win32_imports_from_Macro_code_in_Office_92E97FA1-2EDF-4476-BDD6-9DD0B4DDDC7B.docm
    write-host -foregroundcolor $processmessagecolor "Open document using Word"
    Start-Process winword.exe -ArgumentList ".\Block_Win32_imports_from_Macro_code_in_Office_92E97FA1-2EDF-4476-BDD6-9DD0B4DDDC7B.docm"
    write-host -foregroundcolor $processmessagecolor "Ensure that no macros runs. Please close Word once complete."
    pause
    write-host -foregroundcolor $processmessagecolor "Delete Block_Win32_imports_from_Macro_code_in_Office_92E97FA1-2EDF-4476-BDD6-9DD0B4DDDC7B.docm"
    remove-item .\Block_Win32_imports_from_Macro_code_in_Office_92E97FA1-2EDF-4476-BDD6-9DD0B4DDDC7B.docm  
}

function psexecwmicreation() {
    
    write-host -ForegroundColor white -backgroundcolor blue "`n--- Block Process Creations originating from PSExec & WMI commands ---"
    write-host -foregroundcolor $processmessagecolor "Create DLTEST.VBS file in current directory"
    set-content -Path .\dltest.vbs `
'on error resume next
set process = GetObject("winmgmts:Win32_Process")
WScript.Echo "Executing notepad"
result = process.Create ("notepad.exe",null,null,processid)
WScript.Echo "Method returned result = " & result
WScript.Echo "Id of new process is " & processid'
    write-host -foregroundcolor $processmessagecolor "Execute DLTEST.VBS file in current directory"
    start-process .\dltest.vbs
    write-host "`n1. NOTEPAD should NOT run."
    write-host "2. A dialog should appear that says - Executing notepad"
    write-host "3. After you press OK button, dialog should say - Method returned result = 2"
    write-host "4. After you press OK button again, should say - Id of new process is"
    write-host "5. There should be NO number displayed in this dialog box"
    write-host "6. Press OK button to end test`n"
    write-host -foregroundcolor $warningmessagecolor "If NOTEPAD executed and/or there is a Process Id number displayed, the test has FAILED`n"
    pause
    write-host -foregroundcolor $processmessagecolor "Delete DLTEST.VBS"
    remove-item .\dltest.vbs  
}

function networkprotection() {
    $npdetect = $false
    write-host -ForegroundColor white -backgroundcolor blue "`n--- Network protection ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL https://smartscreentestratings2.net/"
    try {
        $result = Invoke-WebRequest -Uri https://smartscreentestratings2.net/ 
    }
    catch {
        if ($error[0] -match "The remote name could not be resolved") {
            write-host -foregroundcolor $processmessagecolor "The remote name could not be resolved: 'smartscreentestratings2.net'"
        }
        else {
            write-host -foregroundcolor $errormessagecolor "Unknown error"
        }
        $npdetect=$true
    }
    if (-not $npdetect) {
        write-host -foregroundcolor $errormessagecolor "Navigation permitted - test FAILED"
    }
}

function suspiciouspage() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- Suspicious web page ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL https://demo.smartscreen.msft.net/other/areyousure.html"
    start-process -filepath https://demo.smartscreen.msft.net/other/areyousure.html 
    write-host "`n1. Your default browser should open"
    write-host "2. Your browser should indicate security issues with the page`n"
    pause
}

function phishpage() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- Phishing web page ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL https://demo.smartscreen.msft.net/phishingdemo.html"
    start-process -filepath https://demo.smartscreen.msft.net/phishingdemo.html 
    write-host "`n1. Your default browser should open"
    write-host "2. Your browser should indicate security issues with the page and be reported as unsafe`n"
    pause
}

function downloadblock() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- Block download on reputation ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL https://demo.smartscreen.msft.net/download/malwaredemo/freevideo.exe"
    start-process -filepath https://demo.smartscreen.msft.net/download/malwaredemo/freevideo.exe 
    write-host "`n1. Your default browser should open"
    write-host "2. Your browser should indicate security issues with the page and be reported as unsafe`n"
    write-host -foregroundcolor $warningmessagecolor "You be UNABLE to download and save a file from browser to local workstation`n"
    pause
}

function exploitblock() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- Browser exploit protection ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL https://demo.smartscreen.msft.net/other/exploit.html"
    start-process -filepath https://demo.smartscreen.msft.net/other/exploit.html 
    write-host "`n1. Your default browser should open"
    write-host "2. Your browser should indicate security issues with the page and be reported as unsafe`n"
    pause
}

function maliciousframe() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- Mailcious browser frame protection ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL https://demo.smartscreen.msft.net/other/exploit_frame.html"
    start-process -filepath https://demo.smartscreen.msft.net/other/exploit_frame.html 
    write-host "`n1. Your default browser should open"
    write-host "2. Your browser should indicate security issues with a frame in the page and be reported as unsafe`n"
    pause
}

function unknownprogram() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- Unknown program protection ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL https://demo.smartscreen.msft.net/download/unknown/freevideo.exe"
    start-process -filepath https://demo.smartscreen.msft.net/download/unknown/freevideo.exe 
    write-host "`n1. Your default browser should open"
    write-host "2. Your browser should warn that file blocked because it could harm your device`n"
    write-host -foregroundcolor $warningmessagecolor "You be UNABLE to download and save a file from browser to local workstation`n"
    pause
}

function knownmalicious() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- Known malicious program protection ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL https://demo.smartscreen.msft.net/download/known/knownmalicious.exe"
    start-process -filepath https://demo.smartscreen.msft.net/download/known/knownmalicious.exe 
    write-host "`n1. Your default browser should open"
    write-host "2. Your browser should warn that file blocked because it it is malicious`n"
    write-host -foregroundcolor $warningmessagecolor "You be UNABLE to download and save a file from browser to local workstation`n"
    pause
}

function pua() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- Potentially unwanted application protection ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL http://amtso.eicar.org/PotentiallyUnwanted.exe"
    start-process -filepath http://amtso.eicar.org/PotentiallyUnwanted.exe 
    write-host "`n1. Your default browser should open"
    write-host "2. Should not be able to reach this site or download the file`n"
    write-host -foregroundcolor $warningmessagecolor "You be UNABLE to download and save a file from browser to local workstation`n"
    pause
}

<#          Main                #>
Clear-Host
if ($debug) {       # If -debug command line option specified record log file in parent
    write-host -foregroundcolor $processmessagecolor "Create log file ..\sec-test.txt`n"
    Start-transcript "..\sec-test.txt" | Out-Null                                   ## Log file created in current directory that is overwritten on each run
}
write-host -foregroundcolor $systemmessagecolor "Security test script started`n"

downloadfile            # 1
createfile              # 2
inmemorytest            # 3
processdump             # 4
mimikatztest            # 5
failedlogin             # 6
officechildprocess      # 7
officecreateexecutable  # 8
scriptlaunch            # 9
officemacroimport       # 10
psexecwmicreation       # 11
networkprotection       # 12
suspiciouspage          # 13
phishpage               # 14
downloadblock           # 15
exploitblock            # 16
maliciousframe          # 17
unknownprogram          # 18
knownmalicious          # 19
pua                     # 20

write-host -foregroundcolor $systemmessagecolor "`nSecurity test script completed"
if ($debug) {
    Stop-Transcript | Out-Null
}