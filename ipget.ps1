<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Source - https://github.com/directorcia/Office365/blob/master/ipinf.ps1

Description - Report IP information

Prerequisites = 

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"
write-host -foregroundcolor $processmessagecolor "Getting current public IP information `n"
$ipinfo = Invoke-RestMethod http://ipinfo.io/json 

write-host -foregroundcolor $processmessagecolor "IP address =",$ipinfo.ip 
write-host -foregroundcolor $processmessagecolor "Hostname   =",$ipinfo.hostname 
write-host -foregroundcolor $processmessagecolor "City       =",$ipinfo.city 
write-host -foregroundcolor $processmessagecolor "Region     =",$ipinfo.region 
write-host -foregroundcolor $processmessagecolor "Country    =",$ipinfo.country 
write-host -foregroundcolor $processmessagecolor "Location   =",$ipinfo.loc 
write-host -foregroundcolor $processmessagecolor "Org        =",$ipinfo.org

write-host -foregroundcolor $systemmessagecolor "`nScript completed"