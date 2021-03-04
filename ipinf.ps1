<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Source - https://github.com/directorcia/Office365/blob/master/ipget.ps1

Description - Report IP information

Prerequisites = 

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"
$ip2check = read-host -prompt "Enter IP address to check"
$url = "http://api.ipstack.com/$($ip2check)?access_key=1c0181a05c46d0e0c04cc1be36bd8dcd"
$ipinfo = (invoke-webrequest -method GET -ContentType "application/json" -uri $url  -UseBasicParsing).content | convertfrom-json
write-host -foregroundcolor $processmessagecolor "`nIP address =",$ipinfo.ip 
write-host -foregroundcolor $processmessagecolor "Type       =",$ipinfo.type 
write-host -foregroundcolor $processmessagecolor "City       =",$ipinfo.city 
write-host -foregroundcolor $processmessagecolor "Region     =",$ipinfo.region_name 
write-host -foregroundcolor $processmessagecolor "Country    =",$ipinfo.country_name 
write-host -foregroundcolor $processmessagecolor "Location   = (lat)$($ipinfo.latitude), (long)$($ipinfo.longitude)"

write-host -foregroundcolor $systemmessagecolor "`nScript completed"