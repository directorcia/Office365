param(                           
    [switch]$debug = $false      ## if -debug parameter capture log information
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Connect to the Microsoft Graph and get meetign attendees using the Microsoft Graph
Documentation - https://blog.ciaops.com/2023/05/25/get-teams-meeting-attendees-via-powershell-and-the-microsoft-graph/
Source - https://github.com/directorcia/Office365/blob/master/tms-attend-get.ps1
Reference - 

Prerequisites = 1
1. MSGraph PowerShell module installed

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

$scopes = "OnlineMeetingArtifact.Read.All, OnlineMeetings.Read" # required to read meeting data

#region Change to suit environment
$tenantid = "tenant.onmicrosoft.com"                       # e.g. tenant.onmicrosoft.com
$meetingjoinurl = 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_MWY1YWZlNjItNmViNS11MDY4LWI5YjMtMjliYzk5MDY2YDM3%40thread.v2/0?context=%7b%22Tid%22%3a%225243d63d-7632-4d07-a77e-de0fea1ba774%22%2c%22Oid%22%3a%22b75e7296-a058-7074-acb8-6021a3dca444%22%7d'
$useremail = "user@domain.com"                              # user who created the meeting
#endregion

Clear-Host
if ($debug) {
    write-host -foregroundcolor $processmessagecolor "Script activity logged at ..\tms-attend-get.txt"
    start-transcript "..\tms-attend-get.txt"
}

write-host -foregroundcolor $systemmessagecolor "Get Teams meeting attendees - start script`n"

# Connect to MSGraph
write-host -foregroundcolor $processmessagecolor "Connect to the Microsoft Graph"
Select-MgProfile beta
Connect-MgGraph -Tenant $tenantid -Scopes $scopes

# Get user GUID from email address
write-host -foregroundcolor $processmessagecolor "Get meeting creator GUID"
$userid = (get-mguser -UserId $useremail).Id

# Get the meetign id
write-host -foregroundcolor $processmessagecolor "Get meeting id"
$meetingid = Get-MgUserOnlineMeeting -Filter "JoinWebUrl eq '$meetingjoinurl'" -UserId $userid | select-object -ExpandProperty Id

# Get report info from meeting id
write-host -foregroundcolor $processmessagecolor "Get meeting report id"
$reportinfo = Get-MgUserOnlineMeetingAttendanceReport -OnlineMeetingId $meetingid -UserId $userid

# Get meeting report from report id 
write-host -foregroundcolor $processmessagecolor "Get meeting report details`n"
$report = Get-MgUserOnlineMeetingAttendanceReportAttendanceRecord -UserId $userid -OnlineMeetingId $meetingid -MeetingAttendanceReportId $reportinfo.id

# Display attendee list
write-host "Meeting attendees"
write-host "-----------------"
$report.emailaddress

write-host -foregroundcolor $systemmessagecolor "`nGet Teams meeting attendees - end script"
if ($debug) {
    stop-transcript | Out-Null
}
