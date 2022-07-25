param(                         ## if no parameter used then login without MFA and use interactive mode
    [switch]$debug = $false    ## if -debug parameter capture log information
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Script designed to get Microsoft Teams configuration information for a tenant
Source - https://github.com/directorcia/office365/blob/master/o365-tms-import.ps1

Prerequisites = 2
1. Ensure connected to Teams Online - Use the script https://github.com/directorcia/Office365/blob/master/o365-connect-tms.ps1
2. Ensure CSV import file and path are define in variable $importfile. Default is a file called teamsimportdata.csv in current directory

Data file (CSV) column format:

Row 1 (Headings)= TeamsName, TeamType, ChannelName, Owners, Members
Rows 2 - ? = Data to be imported

Columns:
--------
TeamsName = Name you want the new Teams to be called
TeamType = Set to Public to allow all users in your organization to join the group by default. Set to Private to require that an owner approve the join request.
ChannelName = Name you want new channles to be called
Owners = User principal names (i.e. user@domain.com) that you wish to be owners of the Team
Members = User principal names (i.e. user@domain.com) that you wish to be members of the Team

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

$importfile = ".\o365-tms-import.csv"     ## CSV file that contains data to be imported. Change to suit yoru needs

function Create-Channel
{   
   param (   
             $ChannelName,$GroupId
         )   
    Process
    {
        try
            {
                $teamchannels = $ChannelName -split ";" 
                if($teamchannels)
                {
                    for($i =0; $i -le ($teamchannels.count - 1) ; $i++)
                    {
                        New-TeamChannel -GroupId $GroupId -DisplayName $teamchannels[$i] | Out-Null
                        write-host -foregroundcolor $processmessagecolor "  - Created channel",$teamchannels[$i]
                    }
                }
                Else {
                    write-host -foregroundcolor $warningmessagecolor "No import channels found`n"                
                }
            }
        Catch
            {
            }
    }
}

function Add-Users
{   
    param(   
            $Users,$GroupId,$Currentuser,$Role
          )   
    Process
    {
        
        try{
                $teamusers = $Users -split ";" 
                if($teamusers)
                {
                    for($j =0; $j -le ($teamusers.count - 1) ; $j++)
                    {
                        if($teamusers[$j] -ne $CurrentUsername)
                        {
                            Add-TeamUser -GroupId $GroupId -User $teamusers[$j] -Role $Role | Out-Null
                            write-host -foregroundcolor $processmessagecolor "  - Added user",$teamusers[$j],"with role =",$role
                        }
                    }                    
                }
                Else {
                    write-host -foregroundcolor $warningmessagecolor "No import users found`n"                
                }
            }
        Catch
            {
            }
        }
}

function Create-NewTeam
{   
   param (   
             $ImportPath,$Currentuser
         )   
  Process
    {
        $teams = Import-Csv -Path $ImportPath
        foreach($team in $teams)
        {
            $getteam = get-team | where-object { $_.displayname -eq $team.TeamsName}
            If($getteam -eq $null)
            {
                Write-Host "`nStart creating the team: " $team.TeamsName
                $group = New-Team -MailNickName $team.TeamsName -displayname $team.TeamsName -Visibility $team.TeamType
                Write-Host "Creating channels..."
                Create-Channel -ChannelName $team.ChannelName -GroupId $group.GroupId
                Write-Host "Adding team members..."
                Add-Users -Users $team.Members -GroupId $group.GroupId -Currentuser $currentuser -Role Member
                Write-Host "Adding team owners..."
                Add-Users -Users $team.Owners -GroupId $group.GroupId -Currentuser $currentuser -Role Owner
                Write-Host "Completed creating the team: " $team.TeamsName
                $team=$null
            }
            else {
                Write-Host -foregroundcolor $warningmessagecolor "Team $team already exists. No changes will be made"
            }
         }
    }
}

<#      Start Main      #>

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host
if ($debug) {
    write-host -ForegroundColor $processmessagecolor "Script activity logged at ..\o365-tms-import.txt"
    start-transcript "..\o365-tms-import.txt" | Out-Null
}
write-host -foregroundcolor $systemmessagecolor "Start - Microsoft Teams import from CSV`n"
write-host -ForegroundColor $processmessagecolor "Logging = ",$debug

write-host -ForegroundColor $processmessagecolor "Check for Teams PowerShell module"
if (get-module -listavailable -name MicrosoftTeams) {    ## Has the Teams PowerShell module been loaded?
    write-host -ForegroundColor $processmessagecolor "Teams PowerShell found"
}
else {
    write-host -ForegroundColor yellow -backgroundcolor $errormessagecolor "[001] - Teams PowerShell module not installed. Please install and re-run script`n"
    if ($debug) {
        Stop-Transcript                 ## Terminate transcription
    }
    exit 1                          ## Terminate script
}

if (-not (Test-Path $importfile)) {    # test for import CSV file
    write-host -foregroundcolor -backgroundcolor $errormessagecolor "[002] - CSV file $importfile not found`n"
    if ($debug) {
        Stop-Transcript                 ## Terminate transcription
    }
    exit 2                          ## Terminate script    
}

Write-host -ForegroundColor $processmessagecolor "Connect to Microsoft Teams"
$connect = Connect-MicrosoftTeams

Create-NewTeam -ImportPath $importfile -currentuser $connect.account.id

Write-host -ForegroundColor $systemmessagecolor "`nFinish - Microsoft Teams import from CSV"

if ($debug) {
    Stop-Transcript | Out-Null
}

