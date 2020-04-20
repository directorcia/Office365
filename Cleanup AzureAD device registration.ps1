##This script checks for devices registered to AzureAD and removes them so you can successfully perform an AzureAD join. 
# We recommend you backup your registry prior to running. We take no responisbility for the use of this script.

# Source = https://www.inspiredtechs.com.au/fix-for-azure-ad-join-error-code-8018000a-this-device-is-already-enrolled

$sids = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\EnterpriseResourceManager\Tracked' -name |where-object {$_.Length -gt 25}


Foreach ($sid in $sids){

Write-host "Found a registered device. Would you like to remove the device registration settings for SID: $($sid)?" -ForegroundColor Yellow 
    $Readhost = Read-Host " ( y / n ) " 
    Switch ($ReadHost) 
     { 
       Y {Write-host "Yes, Remove registered device"; $removedevice=$true} 
       N {Write-Host "No, do not remove device registration"; $removedevice=$false} 
       Default {Write-Host "Default, Do not remove device registration"; $removedevice=$false} 
     } 


if ($removedevice -eq $true) {

$enrollmentpath = "HKLM:\SOFTWARE\Microsoft\Enrollments\$($sid)"
$entresourcepath = "HKLM:\SOFTWARE\Microsoft\EnterpriseResourceManager\Tracked\$($sid)"


##Remove device from enrollments in registry

$value1 = Test-Path $enrollmentpath
If ($value1 -eq $true) {

write-host "$($sid) exists and will be removed"

Remove-Item -Path $enrollmentpath -Recurse -confirm:$false
Remove-Item -Path $entresourcepath -Recurse -confirm:$false


} 
Else {Write-Host "The value does not exist, skipping"}



##Cleanup scheduled tasks related to device enrollment and the folder for this SID


Get-ScheduledTask -TaskPath "\Microsoft\Windows\EnterpriseMgmt\$($sid)\*"| Unregister-ScheduledTask -Confirm:$false


$scheduleObject = New-Object -ComObject Schedule.Service
$scheduleObject.connect()
$rootFolder = $scheduleObject.GetFolder("\Microsoft\Windows\EnterpriseMgmt")
$rootFolder.DeleteFolder($sid,$null)

Write-Host "Device registration cleaned up for $($sid). If there is more than 1 device registration, we will continue to the next one."
pause


} else { Write-host "Removal has been cancelled for $($sid)"}


}


write-host "Cleanup of device registration has been completed. Ensure you delete the device registration in AzureAD and you can now join your device."


