<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Connect to the Microsoft Graph

Source - https://github.com/directorcia/Office365/blob/master/graph-globaladmins-get.ps1
Description - https://github.com/directorcia/Office365/wiki/Report-Global-Admins

Prerequisites = 1
1. Graph module installed

More scripts available by joining http://www.ciaopspatron.com

#>

param(
    [switch]$csv = $false,

    [ValidateNotNullOrEmpty()]
    [string]$OutputFile = "..\graph-globaladmins.csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$globalAdminRoleTemplateId = "62e90394-69f5-4237-9190-012177145e10"

function Get-GraphPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName
    )

    if ($null -eq $InputObject) {
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($PropertyName)) {
            return $InputObject[$PropertyName]
        }

        return $null
    }

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -ne $property) {
        return $property.Value
    }

    $additionalProperties = $InputObject.PSObject.Properties['AdditionalProperties']
    if ($null -ne $additionalProperties -and $additionalProperties.Value -is [System.Collections.IDictionary]) {
        if ($additionalProperties.Value.Contains($PropertyName)) {
            return $additionalProperties.Value[$PropertyName]
        }
    }

    return $null
}

function Get-AllGraphItems {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )

    $items = [System.Collections.Generic.List[object]]::new()
    $nextLink = $Uri

    while (-not [string]::IsNullOrWhiteSpace($nextLink)) {
        $response = Invoke-MgGraphRequest -Uri $nextLink -Method GET
        foreach ($item in @($response.value)) {
            $items.Add($item)
        }

        $nextLink = [string](Get-GraphPropertyValue -InputObject $response -PropertyName '@odata.nextLink')
    }

    return $items
}

try {
    Connect-MgGraph -Scopes "RoleManagement.Read.Directory", "User.Read.All" -NoWelcome | Out-Null

    # Use role template id for resilience to display name changes.
    $roleLookupUri = "https://graph.microsoft.com/v1.0/directoryRoles?`$filter=roleTemplateId eq '$globalAdminRoleTemplateId'"
    $globalAdminRole = (Get-AllGraphItems -Uri $roleLookupUri | Select-Object -First 1)
    if ($null -eq $globalAdminRole) {
        throw "Unable to locate the Global Administrator directory role in this tenant."
    }

    $memberLookupUri = "https://graph.microsoft.com/v1.0/directoryRoles/$($globalAdminRole.id)/members?`$select=id,displayName,userPrincipalName"
    $globalAdminMembers = Get-AllGraphItems -Uri $memberLookupUri

    $globalAdminSummary = foreach ($member in @($globalAdminMembers)) {
        $principalType = [string](Get-GraphPropertyValue -InputObject $member -PropertyName '@odata.type')
        $displayName = [string](Get-GraphPropertyValue -InputObject $member -PropertyName 'displayName')
        $userPrincipalName = [string](Get-GraphPropertyValue -InputObject $member -PropertyName 'userPrincipalName')
        $memberId = [string](Get-GraphPropertyValue -InputObject $member -PropertyName 'id')

        if ($principalType -eq "#microsoft.graph.user") {
            if ([string]::IsNullOrWhiteSpace($userPrincipalName) -or [string]::IsNullOrWhiteSpace($displayName)) {
                $userLookupUri = "https://graph.microsoft.com/v1.0/users/$memberId?`$select=id,displayName,userPrincipalName"
                $user = Invoke-MgGraphRequest -Uri $userLookupUri -Method GET

                $displayName = [string](Get-GraphPropertyValue -InputObject $user -PropertyName 'displayName')
                $userPrincipalName = [string](Get-GraphPropertyValue -InputObject $user -PropertyName 'userPrincipalName')
                $memberId = [string](Get-GraphPropertyValue -InputObject $user -PropertyName 'id')
            }

            [pscustomobject]@{
                Id                = $memberId
                DisplayName       = $displayName
                UserPrincipalName = $userPrincipalName
                PrincipalType     = "User"
            }
        }
        else {
            [pscustomobject]@{
                Id                = $memberId
                DisplayName       = $displayName
                UserPrincipalName = $null
                PrincipalType     = if ([string]::IsNullOrWhiteSpace([string]$principalType)) { "Unknown" } else { $principalType }
            }
        }
    }

    $globalAdminSummary = $globalAdminSummary | Sort-Object DisplayName
    $globalAdminSummary | Format-Table DisplayName, UserPrincipalName, PrincipalType, Id -AutoSize

    if ($csv) {
        $globalAdminSummary | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
        Write-Host "CSV output saved to $OutputFile"
    }
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
finally {
    try {
        Disconnect-MgGraph | Out-Null
    }
    catch {
        # Ignore disconnect failures.
    }
}
