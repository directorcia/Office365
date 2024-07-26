<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Connect to the Microsoft Graph

Source - https://github.com/directorcia/Office365/blob/master/graph-globaladmins-get.ps1

Prerequisites = 1
1. Graph module installed

More scripts available by joining http://www.ciaopspatron.com

#>

import-module Microsoft.Graph.Identity.DirectoryManagement

Connect-MgGraph -Scopes "RoleManagement.Read.Directory","User.Read.All"

$globalAdmins = Get-MgDirectoryRole | Where-Object { $_.displayName -eq "Global Administrator" }
$globalAdminUsers = Get-MgDirectoryRoleMember -DirectoryRoleId $globalAdmins.id

$globaladminsummary = @()
foreach ($adminuser in $globalAdminUsers) {
    $user = Get-MgUser -userId $adminuser.Id
    $globaladminSummary += [pscustomobject]@{       
        Id                = $adminuser.Id 
        UserPrincipalName = $user.UserPrincipalName
        DisplayName       = $user.DisplayName
    }
}

$globaladminsummary
